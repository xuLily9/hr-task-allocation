"""
Local web server for the Human–Robot Allocation Explorer
=========================================================

Serves index.html and solves with the SAME model as allocate.py (it imports
its functions), so there is only one model to maintain.

Two ways to use the page:
  * the built-in worked example (presets A–D, sliders), and
  * your own project: load a project workbook (see project_template.xlsx) with
    "Load project Excel", adjust crew, rules and weights, and download the
    result workbook.

Put app.py, allocate.py and index.html in one folder, then (inside ortools-env):
    pip install ortools openpyxl      # once
    python app.py
and open http://localhost:8000. Stop with Ctrl+C.
"""

import json
import os
import tempfile
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import allocate as core   # the model and solvers live in allocate.py

PORT = 8000
HERE = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix="hr_alloc_")
PROJECT = {}   # the loaded project: path, name, settings, weights, tasks
RESULT_PATH = os.path.join(TMP, "result.xlsx")


def run_solver(settings, tasks):
    if core.cp_model is not None:
        return core.solve_cpsat(settings, tasks)
    if len(tasks) <= 7:
        res = core.solve_exhaustive(settings, tasks)
        res["engine"] = "exhaustive search (ortools not installed)"
        return res
    try:
        return core.solve_with_milp(settings, tasks)
    except ImportError:
        raise RuntimeError("ortools is needed for more than 7 tasks: pip install ortools")


def pack(res):
    out = {"status": res["status"], "seconds": res["seconds"], "engine": res["engine"]}
    if "items" in res:
        out.update({"items": [dict(it, task=i) for i, it in enumerate(res["items"])],
                    "makespan": res["makespan"], "costInt": res["cost_int"],
                    "expo": res["exposure"], "obj": res["obj"]})
    return out


# ------------------------------------------------------------- built-in example
def solve_example(cfg):
    """cfg comes from the page: tasks with dur / exp / cost (x1000) per option, crew and rules."""
    nW, nR = cfg["nW"], cfg["nR"]
    tasks = []
    for t in cfg["tasks"]:
        opts = {}
        for k, d in t["dur"].items():
            if d is None or k not in t["cost"]:
                continue
            if (core.USES_WORKER[k] and nW == 0) or (core.USES_ROBOT[k] and nR == 0):
                continue
            opts[k] = {"dur": int(d), "exp": int(t["exp"][k] or 0), "cost_int": int(t["cost"][k])}
        if not opts:
            return {"status": "INFEASIBLE", "seconds": 0, "engine": "CP-SAT"}
        tasks.append({"id": t["id"], "preds": t["preds"], "opts": opts})
    settings = {"nW": nW, "nR": nR, "lam": cfg["lam"] / core.SCALE, "zone": cfg["zone"],
                "cap": cfg["cap"], "time_limit": 10}
    return pack(run_solver(settings, tasks))


# ------------------------------------------------------------- your own project
def load_project(data, name):
    path = os.path.join(TMP, "project.xlsx")
    with open(path, "wb") as f:
        f.write(data)
    try:
        settings, weights, tasks = core.read_project(path, filter_crew=False)
    except core.InputError:
        raise
    except Exception:
        raise core.InputError(f"{name} could not be opened as an Excel workbook. Save it as .xlsx, "
                              "starting from project_template.xlsx, and try again.")
    PROJECT.clear()
    if os.path.exists(RESULT_PATH):
        os.remove(RESULT_PATH)
    PROJECT.update(path=path, name=name, settings=settings, weights=weights)
    return {"name": name, "nTasks": len(tasks), "hasPanels": settings["has_panels"],
            "panels": settings["panels"], "perPanel": settings["per_panel_tasks"],
            "settings": {k: settings[k] for k in ("nW", "nR", "lam", "zone", "cap")},
            "weights": weights}


def solve_project(cfg):
    if not PROJECT:
        raise RuntimeError("No project loaded yet.")
    settings, _, tasks = core.read_project(PROJECT["path"], filter_crew=False, panels=cfg.get("panels"))
    settings.update(nW=int(cfg["nW"]), nR=int(cfg["nR"]), lam=float(cfg["lam"]),
                    zone=bool(cfg["zone"]), cap=cfg["cap"])
    w = {k: float(cfg["weights"][k]) for k in core.CRIT}
    total = sum(w.values()) or 1
    weights = {k: v / total for k, v in w.items()}

    for t in tasks:   # drop options the current crew cannot staff
        t["opts"] = {k: o for k, o in t["opts"].items()
                     if not ((core.USES_WORKER[k] and settings["nW"] == 0) or (core.USES_ROBOT[k] and settings["nR"] == 0))}
        if not t["opts"]:
            return {"status": "INFEASIBLE", "seconds": 0, "engine": "CP-SAT",
                    "message": f"Task {t['id']} has no option the current crew can do.", "tasks": []}
    core.build_costs(settings, weights, tasks)
    res = run_solver(settings, tasks)
    ready = os.path.exists(RESULT_PATH)
    if cfg.get("write", True):     # scenario comparisons do not overwrite the downloadable result
        try:
            core.write_result(PROJECT["path"], RESULT_PATH, settings, weights, tasks, res, "not run")
            ready = True
        except SystemExit:
            ready = False
    out = pack(res)
    out["resultReady"] = ready
    out["tasks"] = [{"id": t["id"], "name": t["name"], "preds": t["preds"],
                     "opts": {k: {"dur": o["dur"], "exp": o["exp"], "cost": o["cost_int"] / core.SCALE}
                              for k, o in t["opts"].items()}} for t in tasks]
    return out


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json", extra=None):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(os.path.join(HERE, "index.html"), "rb") as f:
                self._send(200, f.read(), "text/html; charset=utf-8")
        elif self.path == "/health":
            self._send(200, json.dumps({"cpsat": core.cp_model is not None, "project": bool(PROJECT)}))
        elif self.path.startswith("/project/result.xlsx"):
            if not os.path.exists(RESULT_PATH):
                return self._send(404, json.dumps({"error": "no result yet"}))
            name = os.path.splitext(PROJECT.get("name", "project"))[0] + "_result.xlsx"
            with open(RESULT_PATH, "rb") as f:
                self._send(200, f.read(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           {"Content-Disposition": f'attachment; filename="{name}"'})
        else:
            self._send(404, json.dumps({"error": "not found"}))

    def do_POST(self):
        try:
            n = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(n)
            if self.path == "/solve":
                out = solve_example(json.loads(body))
            elif self.path.startswith("/project/load"):
                name = self.headers.get("X-File-Name", "project.xlsx")
                out = load_project(body, name)
            elif self.path == "/project/solve":
                out = solve_project(json.loads(body))
            else:
                return self._send(404, json.dumps({"error": "not found"}))
            self._send(200, json.dumps(out))
        except core.InputError as exc:
            self._send(400, json.dumps({"error": str(exc)}))
        except SystemExit as exc:      # e.g. openpyxl missing
            self._send(500, json.dumps({"error": str(exc)}))
        except Exception as exc:       # report the problem to the page instead of crashing
            self._send(500, json.dumps({"error": f"{type(exc).__name__}: {exc}"}))

    def log_message(self, *args):  # keep the terminal quiet
        pass


if __name__ == "__main__":
    if core.cp_model is None:
        print("Note: ortools not found, so small cases use exhaustive search. "
              "For CP-SAT, activate ortools-env first: source ortools-env/bin/activate")
    url = f"http://localhost:{PORT}"
    print(f"Human–Robot Allocation Explorer running at {url}  (Ctrl+C to stop)")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
