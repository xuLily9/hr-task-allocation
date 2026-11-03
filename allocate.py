"""
Human–robot task allocation for any project
============================================

Reads a project workbook (see project_template.xlsx), builds the safety-aware
allocation and scheduling model, solves it with OR-Tools CP-SAT and writes the
result to a new workbook with a Results sheet and a 5-minute Timeline sheet.

    python allocate.py project_template.xlsx
    python allocate.py my_project.xlsx --out my_result.xlsx
    python allocate.py project_template.xlsx --zone yes --cap 40    (override Settings)

Model
    min  sum_{i,k} C_ik x_ik + lambda * makespan
    s.t. one option per task; infeasible options excluded; precedence;
         each worker and each robot does one task at a time (HR uses one of each);
         zone rule (W and R never overlap) and exposure cap, if switched on.
    C_ik = sum_j w_j s_ijk over C1–C7:
         C1 = min(1, duration / reference duration), C2 = min(1, cost / reference cost),
         C3–C7 = (rubric rating 1–5 - 1) / 4, rated per task–option pair for the project.

Experts, once:   Weights (BWM), hard rules, and the scoring rubric for C3–C7.
Each project:    Settings, Tasks, and Options (duration, exposure, rubric ratings per task–option pair).

Without ortools (or for up to 7 tasks as a cross-check) an exhaustive search is used.
"""

import argparse
import itertools
import os
import sys
import time


try:
    from ortools.sat.python import cp_model
except ImportError:
    cp_model = None

MODES = ["W", "R", "HR"]
USES_WORKER = {"W": True, "R": False, "HR": True}
USES_ROBOT = {"W": False, "R": True, "HR": True}
CRIT = ["C1", "C2", "C3", "C4", "C5", "C6", "C7"]
SCALE = 1000


class InputError(Exception):
    pass


def _xl():
    """Import openpyxl only when Excel files are read or written (app.py does not need it)."""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        sys.exit("openpyxl is needed to read and write Excel files: pip install openpyxl")
    return openpyxl, Alignment, Font, PatternFill


# --------------------------------------------------------------------------- reading
def read_project(path, filter_crew=True, panels=None):
    openpyxl = _xl()[0]
    wb = openpyxl.load_workbook(path, data_only=False)
    for name in ("Settings", "Weights", "Tasks", "Options"):
        if name not in wb.sheetnames:
            raise InputError(f"Sheet '{name}' is missing. Start from project_template.xlsx.")

    s = {str(r[0]).strip(): r[1] for r in wb["Settings"].iter_rows(min_row=2, values_only=True) if r[0]}
    def num(key, cast=float, minimum=None, blank_ok=False):
        v = s.get(key)
        if v in (None, ""):
            if blank_ok:
                return None
            raise InputError(f"Settings: '{key}' is empty.")
        try:
            v = cast(v)
        except (TypeError, ValueError):
            raise InputError(f"Settings: '{key}' must be a number, found {v!r}.")
        if minimum is not None and v < minimum:
            raise InputError(f"Settings: '{key}' must be at least {minimum}.")
        return v
    settings = {
        "nW": num("Workers", int, 1), "nR": num("Robots", int, 0),
        "rate": {"W": num("Cost rate W"), "R": num("Cost rate R"), "HR": num("Cost rate HR")},
        "ref_dur": num("Reference duration", float, 1e-9), "ref_cost": num("Reference cost", float, 1e-9),
        "lam": num("Lambda", float, 0),
        "zone": str(s.get("Zone rule", "No")).strip().lower() in ("yes", "y", "true", "1"),
        "cap": num("Exposure cap", float, 0, blank_ok=True),
        "time_limit": num("Time limit", float, 1, blank_ok=True) or 30,
        "panels": num("Panels", int, 1, blank_ok=True) or 1,
    }

    weights = {}
    for r in wb["Weights"].iter_rows(min_row=2, values_only=True):
        if r[0] in CRIT:
            if r[2] in (None, "") or float(r[2]) < 0:
                raise InputError(f"Weights: weight of {r[0]} is empty or negative.")
            weights[r[0]] = float(r[2])
    if set(weights) != set(CRIT):
        raise InputError(f"Weights: need C1–C7, found {sorted(weights)}.")
    total = sum(weights.values())
    if total <= 0:
        raise InputError("Weights: the weights sum to zero.")
    weights = {k: v / total for k, v in weights.items()}

    # ---- Tasks: a task can be repeated for every panel ("Per panel" = Yes)
    ws = wb["Tasks"]
    head = [str(c.value or "").strip().lower() for c in ws[1]]
    def col(*names):
        for nm in names:
            for j, h in enumerate(head):
                if h.startswith(nm):
                    return j
        return None
    c_per, c_prev = col("per panel"), col("after previous panel")
    templ, tindex = [], {}
    for n, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not r[0]:
            continue
        tid = str(r[0]).strip()
        if tid in tindex:
            raise InputError(f"Tasks row {n}: task ID {tid} appears twice.")
        split = lambda v: [p.strip() for p in str(v).split(",") if p.strip()] if v not in (None, "") else []
        per = c_per is not None and str(r[c_per] or "").strip().lower() in ("yes", "y", "true", "1")
        prev = split(r[c_prev]) if c_prev is not None else []
        tindex[tid] = len(templ)
        templ.append({"id": tid, "name": r[1] or "", "pred_ids": split(r[2]), "per": per, "prev": prev, "opts": {}})
    if not templ:
        raise InputError("Tasks: the sheet has no tasks.")
    for t in templ:
        for p in t["pred_ids"] + t["prev"]:
            if p not in tindex:
                raise InputError(f"Tasks: {t['id']} lists predecessor {p!r}, which is not a task ID.")
        for p in t["prev"]:
            if not (t["per"] and templ[tindex[p]]["per"]):
                raise InputError(f"Tasks: 'After previous panel' of {t['id']} must name a per-panel task, on a per-panel task.")

    for n, r in enumerate(wb["Options"].iter_rows(min_row=2, values_only=True), start=2):
        if not r[0]:
            continue
        tid, k = str(r[0]).strip(), str(r[1] or "").strip().upper()
        if tid not in tindex:
            raise InputError(f"Options row {n}: task {tid} is not in the Tasks sheet.")
        if k not in MODES:
            raise InputError(f"Options row {n}: option must be W, R or HR, found {r[1]!r}.")
        t = templ[tindex[tid]]
        if k in t["opts"]:
            raise InputError(f"Options row {n}: {tid} {k} appears twice.")
        try:
            d = int(round(float(r[2])))
            e = int(round(float(r[3] or 0)))
            rating = [float(v) for v in r[4:9]]
        except (TypeError, ValueError):
            raise InputError(f"Options row {n} ({tid} {k}): duration, exposure and the five ratings must all be numbers.")
        if d <= 0:
            raise InputError(f"Options row {n} ({tid} {k}): duration must be positive.")
        if any(v < 1 or v > 5 for v in rating):
            raise InputError(f"Options row {n} ({tid} {k}): ratings must be between 1 and 5.")
        if filter_crew and ((USES_WORKER[k] and settings["nW"] == 0) or (USES_ROBOT[k] and settings["nR"] == 0)):
            continue
        t["opts"][k] = {"dur": d, "exp": e, "rating": rating}
    for t in templ:
        if not t["opts"]:
            raise InputError(f"Task {t['id']}: no feasible option. Add a row in Options, or add workers or robots.")

    # ---- expand per-panel tasks
    has_panels = any(t["per"] for t in templ)
    n_pan = int(panels if panels is not None else settings["panels"]) if has_panels else 1
    if has_panels and not 1 <= n_pan <= 26:
        raise InputError("Settings: 'Panels' must be between 1 and 26.")
    settings["panels"], settings["has_panels"] = n_pan, has_panels
    settings["per_panel_tasks"] = sum(t["per"] for t in templ)
    letters = [chr(65 + i) for i in range(n_pan)]
    first_per = next((i for i, t in enumerate(templ) if t["per"]), len(templ))
    order = ([(t, None) for t in templ[:first_per] if not t["per"]]
             + [(t, L) for L in letters for t in templ if t["per"]]
             + [(t, None) for t in templ[first_per:] if not t["per"]])
    eid = lambda t, L: t["id"] + L if (t["per"] and L) else t["id"]
    tasks, index = [], {}
    for t, L in order:
        index[eid(t, L)] = len(tasks)
        tasks.append({"id": eid(t, L), "name": t["name"] + (f", panel {L}" if L else ""),
                      "opts": {k: dict(o) for k, o in t["opts"].items()}, "_t": t, "_L": L})
    for e in tasks:
        t, L = e["_t"], e["_L"]
        ids = []
        for p in t["pred_ids"]:
            pt = templ[tindex[p]]
            if pt["per"]:
                ids += [eid(pt, L)] if L else [eid(pt, M) for M in letters]
            else:
                ids.append(p)
        if L and letters.index(L) > 0:
            ids += [eid(templ[tindex[p]], letters[letters.index(L) - 1]) for p in t["prev"]]
        e["pred_ids"] = list(dict.fromkeys(ids))
        e["preds"] = [index[p] for p in e["pred_ids"]]
        del e["_t"], e["_L"]
    check_acyclic(tasks)
    return settings, weights, tasks


def check_acyclic(tasks):
    state = [0] * len(tasks)
    def visit(i, path):
        if state[i] == 1:
            raise InputError("Tasks: the predecessors form a loop: " + " → ".join(tasks[j]["id"] for j in path + [i]))
        if state[i] == 0:
            state[i] = 1
            for p in tasks[i]["preds"]:
                visit(p, path + [i])
            state[i] = 2
    for i in range(len(tasks)):
        visit(i, [])


# --------------------------------------------------------------------------- scores and cost
def build_costs(settings, weights, tasks):
    """Fixed scoring rules, so a weight means the same thing in every project."""
    for t in tasks:
        for k, o in t["opts"].items():
            s = {"C1": min(1.0, o["dur"] / settings["ref_dur"]),
                 "C2": min(1.0, settings["rate"][k] * o["dur"] / settings["ref_cost"])}
            for j, c in enumerate(CRIT[2:]):
                s[c] = (o["rating"][j] - 1) / 4
            o["scores"] = s
            o["cost"] = sum(weights[c] * s[c] for c in CRIT)
            o["cost_int"] = round(SCALE * round(o["cost"], 6))   # strip float noise before rounding


# --------------------------------------------------------------------------- CP-SAT
def solve_cpsat(settings, tasks):
    nW, nR, zone, cap = settings["nW"], settings["nR"], settings["zone"], settings["cap"]
    lam = round(SCALE * settings["lam"])
    m = cp_model.CpModel()
    horizon = sum(max(o["dur"] for o in t["opts"].values()) for t in tasks)
    start = [m.new_int_var(0, horizon, f"s{i}") for i in range(len(tasks))]
    end = [m.new_int_var(0, horizon, f"e{i}") for i in range(len(tasks))]
    x, iv, yw, ivw, yr, ivr = {}, {}, {}, {}, {}, {}
    for i, t in enumerate(tasks):
        for k, o in t["opts"].items():
            d = o["dur"]
            x[i, k] = m.new_bool_var(f"x{i}{k}")
            iv[i, k] = m.new_optional_interval_var(start[i], d, end[i], x[i, k], f"iv{i}{k}")
            if USES_WORKER[k]:
                for w in range(nW):
                    yw[i, k, w] = m.new_bool_var(f"yw{i}{k}{w}")
                    ivw[i, k, w] = m.new_optional_interval_var(start[i], d, end[i], yw[i, k, w], f"ivw{i}{k}{w}")
                m.add(sum(yw[i, k, w] for w in range(nW)) == x[i, k])
            if USES_ROBOT[k]:
                for r in range(nR):
                    yr[i, k, r] = m.new_bool_var(f"yr{i}{k}{r}")
                    ivr[i, k, r] = m.new_optional_interval_var(start[i], d, end[i], yr[i, k, r], f"ivr{i}{k}{r}")
                m.add(sum(yr[i, k, r] for r in range(nR)) == x[i, k])
        m.add_exactly_one(x[i, k] for k in t["opts"])
        for p in t["preds"]:
            m.add(start[i] >= end[p])
    for w in range(nW):
        m.add_no_overlap([v for (i, k, ww), v in ivw.items() if ww == w])
    for r in range(nR):
        m.add_no_overlap([v for (i, k, rr), v in ivr.items() if rr == r])
    if zone:
        m.add_no_overlap([v for (i, k), v in iv.items() if k in ("W", "R")])
    if cap is not None:
        m.add(sum(tasks[i]["opts"][k]["exp"] * v for (i, k), v in x.items()) <= int(cap))
    makespan = m.new_int_var(0, horizon, "makespan")
    m.add_max_equality(makespan, end)
    m.minimize(sum(tasks[i]["opts"][k]["cost_int"] * v for (i, k), v in x.items()) + lam * makespan)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = settings["time_limit"]
    solver.parameters.num_workers = 8
    t0 = time.perf_counter()
    status = solver.solve(m)
    secs = time.perf_counter() - t0
    name = solver.status_name(status)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {"status": name, "seconds": secs, "engine": "CP-SAT"}
    items = []
    for i, t in enumerate(tasks):
        k = next(k for k in t["opts"] if solver.value(x[i, k]))
        w = next((ww for ww in range(nW) if (i, k, ww) in yw and solver.value(yw[i, k, ww])), -1)
        r = next((rr for rr in range(nR) if (i, k, rr) in yr and solver.value(yr[i, k, rr])), -1)
        items.append({"mode": k, "s": solver.value(start[i]), "e": solver.value(end[i]), "w": w, "r": r})
    return finish({"status": name, "seconds": secs, "engine": "CP-SAT", "items": items}, settings, tasks)


# --------------------------------------------------------------------------- exhaustive search
def topo_orders(tasks):
    n = len(tasks)
    def rec(done, order):
        if len(order) == n:
            yield list(order)
            return
        for i in range(n):
            if i not in done and all(p in done for p in tasks[i]["preds"]):
                done.add(i); order.append(i)
                yield from rec(done, order)
                order.pop(); done.discard(i)
    yield from rec(set(), [])


def schedule(tasks, modes, order, nW, nR, zone):
    wl, rl, placed, allv = [[] for _ in range(nW)], [[] for _ in range(nR)], [None] * len(tasks), []
    ov = lambda a, b: a[0] < b[1] and b[0] < a[1]
    for i in order:
        k = modes[i]; d = tasks[i]["opts"][k]["dur"]
        ready = max([placed[p]["e"] for p in tasks[i]["preds"]], default=0)
        for s in sorted({ready, *[v["e"] for v in allv if v["e"] > ready]}):
            itv = (s, s + d)
            w = r = -1
            if USES_WORKER[k]:
                w = next((j for j, l in enumerate(wl) if all(not ov(itv, (v["s"], v["e"])) for v in l)), -1)
                if w < 0: continue
            if USES_ROBOT[k]:
                r = next((j for j, l in enumerate(rl) if all(not ov(itv, (v["s"], v["e"])) for v in l)), -1)
                if r < 0: continue
            if zone and k in ("W", "R"):
                other = "R" if k == "W" else "W"
                if any(v["mode"] == other and ov(itv, (v["s"], v["e"])) for v in allv): continue
            rec = {"mode": k, "s": s, "e": s + d, "w": w, "r": r}
            if w >= 0: wl[w].append(rec)
            if r >= 0: rl[r].append(rec)
            placed[i] = rec; allv.append(rec)
            break
        else:
            return None
    return placed


def solve_exhaustive(settings, tasks):
    t0 = time.perf_counter()
    lam = round(SCALE * settings["lam"]); cap = settings["cap"]
    orders = list(topo_orders(tasks))
    best = None
    for modes in itertools.product(*[list(t["opts"]) for t in tasks]):
        expo = sum(t["opts"][k]["exp"] for t, k in zip(tasks, modes))
        if cap is not None and expo > cap:
            continue
        c = sum(t["opts"][k]["cost_int"] for t, k in zip(tasks, modes))
        for o in orders:
            items = schedule(tasks, modes, o, settings["nW"], settings["nR"], settings["zone"])
            if items is None:
                continue
            ms = max(v["e"] for v in items)
            obj = c + lam * ms
            if best is None or obj < best[0]:
                best = (obj, items)
    secs = time.perf_counter() - t0
    if best is None:
        return {"status": "INFEASIBLE", "seconds": secs, "engine": "exhaustive search"}
    return finish({"status": "OPTIMAL", "seconds": secs, "engine": "exhaustive search", "items": best[1]}, settings, tasks)


def finish(res, settings, tasks):
    items = res["items"]
    res["makespan"] = max(v["e"] for v in items)
    res["cost_int"] = sum(tasks[i]["opts"][v["mode"]]["cost_int"] for i, v in enumerate(items))
    res["cost"] = res["cost_int"] / SCALE   # same rounding as the solver objective
    res["exposure"] = sum(tasks[i]["opts"][v["mode"]]["exp"] for i, v in enumerate(items))
    res["obj"] = res["cost_int"] + round(SCALE * settings["lam"]) * res["makespan"]
    return res


# --------------------------------------------------------------------------- MILP fallback
def assign_lanes(plan, nW, nR):
    """Give each task a specific worker / robot (interval colouring by start time)."""
    order = sorted(range(len(plan)), key=lambda i: plan[i][1])
    wfree, rfree = [0] * nW, [0] * nR
    items = [None] * len(plan)
    for i in order:
        k, s, e = plan[i]
        w = r = -1
        if USES_WORKER[k]:
            w = next(j for j in range(nW) if wfree[j] <= s); wfree[w] = e
        if USES_ROBOT[k]:
            r = next(j for j in range(nR) if rfree[j] <= s); rfree[r] = e
        items[i] = {"mode": k, "s": s, "e": e, "w": w, "r": r}
    return items


def solve_with_milp(settings, tasks):
    from verify_milp import solve_milp      # needs scipy; kept in a separate file
    t0 = time.perf_counter()
    r = solve_milp(settings, tasks, time_limit=settings["time_limit"] * 10)
    secs = time.perf_counter() - t0
    if "plan" not in r:
        return {"status": r["status"], "seconds": secs, "engine": "HiGHS MILP"}
    status = "OPTIMAL" if r["status"] == "OPTIMAL" else "FEASIBLE"
    res = {"status": status, "seconds": secs, "engine": "HiGHS MILP (fallback)",
           "items": assign_lanes(r["plan"], settings["nW"], settings["nR"])}
    return finish(res, settings, tasks)


# --------------------------------------------------------------------------- writing
FILL_HEX = {"W": "F0B66B", "R": "1F3F66", "HR": "6A5A9E"}
INK = {"W": "1B2433", "R": "FFFFFF", "HR": "FFFFFF"}
A = "Arial"


def write_result(src, out, settings, weights, tasks, res, check):
    openpyxl, Alignment, Font, PatternFill = _xl()
    FILL = {k: PatternFill("solid", fgColor=v) for k, v in FILL_HEX.items()}
    wb = openpyxl.load_workbook(src)
    for name in ("Results", "Timeline"):
        if name in wb.sheetnames:
            del wb[name]
    ws = wb.create_sheet("Results", 0)
    b, f = Font(name=A, bold=True), Font(name=A)
    rows = [
        ("Status", res["status"]), ("Solver", res["engine"]),
        ("Solve time (s)", round(res["seconds"], 3)),
        ("Cross-check (exhaustive search)", check),
        ("Makespan (min)", res.get("makespan")), ("Weighted cost Σ C", round(res["cost"], 3) if "cost" in res else None),
        ("Worker exposure (min)", res.get("exposure")),
        ("Crew", f"{settings['nW']} worker(s), {settings['nR']} robot(s)"),
        ("Rules", f"λ = {settings['lam']}, zone rule {'on' if settings['zone'] else 'off'}, "
                  f"exposure cap {settings['cap'] if settings['cap'] is not None else 'off'}"),
    ]
    ws["A1"] = "Allocation result"; ws["A1"].font = Font(name=A, bold=True, size=14)
    for i, (k, v) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=k).font = b
        ws.cell(row=i, column=2, value=v).font = f
    if "items" not in res:
        ws.cell(row=13, column=1, value="No feasible plan: relax a safety rule, add workers or robots, or check feasibility.").font = b
        wb.save(out); return
    hdr = ["Task ID", "Task name", "Option", "Executor", "Start (min)", "End (min)",
           "C(i,k)", "Exposure (min)"] + [f"s {c}" for c in CRIT]
    r0 = 14
    for j, h in enumerate(hdr, start=1):
        c = ws.cell(row=r0, column=j, value=h)
        c.font, c.fill = Font(name=A, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="14213D")
    for n, (t, it) in enumerate(zip(tasks, res["items"]), start=r0 + 1):
        o = t["opts"][it["mode"]]
        ex = " + ".join(x for x in [f"Worker {it['w'] + 1}" if it["w"] >= 0 else "",
                                    f"Robot {it['r'] + 1}" if it["r"] >= 0 else ""] if x)
        vals = [t["id"], t["name"], it["mode"], ex, it["s"], it["e"],
                o["cost_int"] / SCALE, o["exp"]] + [round(o["scores"][c], 3) for c in CRIT]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=n, column=j, value=v).font = f
        ws.cell(row=n, column=3).fill = FILL[it["mode"]]
        ws.cell(row=n, column=3).font = Font(name=A, bold=True, color=INK[it["mode"]])
    for col, w in zip("ABCDEFGH", [30, 26, 8, 22, 11, 10, 9, 13]):
        ws.column_dimensions[col].width = w

    # Timeline: one column per 5 minutes
    tl = wb.create_sheet("Timeline", 1)
    tl["A1"] = "Timeline (each column = 5 minutes)"; tl["A1"].font = Font(name=A, bold=True, size=14)
    step = 5
    ncols = -(-res["makespan"] // step)
    tl.cell(row=3, column=1, value="Lane / minute").font = b
    for c in range(ncols):
        cell = tl.cell(row=3, column=2 + c, value=c * step)
        cell.font, cell.alignment = Font(name=A, size=8), Alignment(horizontal="center")
        tl.column_dimensions[openpyxl.utils.get_column_letter(2 + c)].width = 4.2
    tl.column_dimensions["A"].width = 14
    lanes = [("Worker", w) for w in range(settings["nW"])] + [("Robot", r) for r in range(settings["nR"])]
    for li, (kind, idx) in enumerate(lanes):
        row = 4 + li
        tl.cell(row=row, column=1, value=f"{kind} {idx + 1}").font = b
        tl.row_dimensions[row].height = 24
        for i, it in enumerate(res["items"]):
            on = (kind == "Worker" and it["w"] == idx) or (kind == "Robot" and it["r"] == idx)
            if not on:
                continue
            first = True
            for c in range(it["s"] // step, -(-it["e"] // step)):
                cell = tl.cell(row=row, column=2 + c)
                cell.fill = FILL[it["mode"]]
                if first:
                    cell.value = f"{tasks[i]['id']}·{it['mode']}"
                    cell.font = Font(name=A, size=8, bold=True, color=INK[it["mode"]])
                    first = False
    leg = 5 + len(lanes)
    tl.cell(row=leg, column=1, value="Legend").font = b
    for j, (k, lab) in enumerate([("W", "W worker alone"), ("R", "R robot alone"), ("HR", "HR human–robot team")]):
        col = 2 + j * 8
        tl.merge_cells(start_row=leg, start_column=col, end_row=leg, end_column=col + 6)
        c = tl.cell(row=leg, column=col, value=lab)
        c.fill, c.font = FILL[k], Font(name=A, size=9, bold=True, color=INK[k])
    wb.save(out)


# --------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Safety-aware human–robot task allocation for any project.")
    ap.add_argument("workbook", help="project workbook, e.g. project_template.xlsx")
    ap.add_argument("--out", help="output workbook (default: <input>_result.xlsx)")
    ap.add_argument("--lam", type=float, help="override Lambda")
    ap.add_argument("--zone", choices=["yes", "no"], help="override Zone rule")
    ap.add_argument("--cap", help="override Exposure cap (number, or 'off')")
    ap.add_argument("--workers", type=int, help="override Workers")
    ap.add_argument("--robots", type=int, help="override Robots")
    ap.add_argument("--panels", type=int, help="override Panels (number of times per-panel tasks repeat)")
    a = ap.parse_args()
    try:
        settings, weights, tasks = read_project(a.workbook, panels=a.panels)
    except InputError as e:
        sys.exit(f"Input problem: {e}")
    if a.lam is not None: settings["lam"] = a.lam
    if a.zone: settings["zone"] = a.zone == "yes"
    if a.cap: settings["cap"] = None if a.cap.lower() == "off" else float(a.cap)
    if a.workers is not None: settings["nW"] = a.workers
    if a.robots is not None: settings["nR"] = a.robots
    build_costs(settings, weights, tasks)

    small = len(tasks) <= 7
    if cp_model is not None:
        res = solve_cpsat(settings, tasks)
    elif small:
        print("ortools not found: using exhaustive search (fine for up to 7 tasks).")
        res = solve_exhaustive(settings, tasks)
    else:
        try:
            import scipy  # noqa: F401
        except ImportError:
            sys.exit("ortools is needed for more than 7 tasks: activate ortools-env, or pip install ortools.")
        print("ortools not found: using the HiGHS MILP fallback (slower; install ortools for CP-SAT).")
        res = solve_with_milp(settings, tasks)

    check = "not run (more than 7 tasks)"
    if small and res["engine"] == "CP-SAT":
        ex = solve_exhaustive(settings, tasks)
        if "obj" in ex and "obj" in res:
            check = "agrees (same objective)" if ex["obj"] == res["obj"] else f"DIFFERS: {ex['obj']} vs {res['obj']}"
        else:
            check = "agrees (no feasible plan)" if ("obj" not in ex and "obj" not in res) else "DIFFERS on feasibility"
    elif res["engine"] != "CP-SAT":
        check = f"not run ({res['engine']} was the solver)"

    out = a.out or os.path.splitext(a.workbook)[0] + "_result.xlsx"
    write_result(a.workbook, out, settings, weights, tasks, res, check)

    print(f"{res['engine']}: {res['status']} in {res['seconds']:.2f} s · cross-check: {check}")
    if "items" in res:
        for t, it in zip(tasks, res["items"]):
            ex = " + ".join(x for x in [f"W{it['w'] + 1}" if it["w"] >= 0 else "", f"R{it['r'] + 1}" if it["r"] >= 0 else ""] if x)
            print(f"  {t['id']:6s} {t['name'][:24]:24s} -> {it['mode']:2s} {ex:8s} {it['s']:4d}–{it['e']:<4d} min")
        print(f"  Makespan {res['makespan']} min · cost {res['cost']:.3f} · worker exposure {res['exposure']} min")
    print(f"Saved {out}")


if __name__ == "__main__":
    main()
