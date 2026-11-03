"""
Independent cross-check of allocate.py with a MILP solver (HiGHS, via SciPy)
============================================================================

Reads the same project workbook, builds a time-indexed MILP of the same problem
and solves it with HiGHS. If CP-SAT and HiGHS report the same optimal objective,
the CP-SAT model is confirmed by a second, differently formulated model.

    python verify_milp.py project_template.xlsx
    python verify_milp.py worked_example.xlsx --zone yes --cap 40

Needs: pip install scipy openpyxl   (SciPy 1.9 or later)

Formulation (time step g = greatest common divisor of all durations):
    y[i,k,t] = 1 if task i starts with option k in slot t
    sum_{k,t} y[i,k,t] = 1                                   one option and one start per task
    start(j) >= end(i)                                       precedence
    sum of worker-using tasks active in slot u <= workers    (HR needs a worker and a robot)
    sum of robot-using tasks active in slot u  <= robots
    zone rule: W-alone and R-alone never active in the same slot (binary a[u])
    sum exposure <= cap
    M >= end(i);  minimise sum C_ik y + lambda * g * M
With identical workers (and identical robots), a slot-by-slot capacity limit is
equivalent to assigning tasks to individual workers without overlap, so this
model has the same optimum as the CP-SAT model in allocate.py.
"""

import argparse
import math
import sys
from functools import reduce

import numpy as np
from scipy.optimize import Bounds, LinearConstraint, milp
from scipy.sparse import lil_matrix

import allocate as core


def solve_milp(settings, tasks, time_limit=120):
    durs = [o["dur"] for t in tasks for o in t["opts"].values()]
    g = reduce(math.gcd, durs)
    T = sum(max(o["dur"] for o in t["opts"].values()) for t in tasks) // g   # horizon in slots
    nW, nR, zone, cap = settings["nW"], settings["nR"], settings["zone"], settings["cap"]
    lam = round(core.SCALE * settings["lam"])

    var = []                          # (i, k, t, d) for y variables
    for i, t in enumerate(tasks):
        for k, o in t["opts"].items():
            d = o["dur"] // g
            for s in range(T - d + 1):
                var.append((i, k, s, d))
    ny = len(var)
    na = T if zone else 0
    iM = ny + na
    n = ny + na + 1

    c = np.zeros(n)
    for v, (i, k, s, d) in enumerate(var):
        c[v] = tasks[i]["opts"][k]["cost_int"]
    c[iM] = lam * g

    rows, lo, hi = [], [], []
    def add(coefs, l, h):
        rows.append(coefs); lo.append(l); hi.append(h)

    by_task = {}
    for v, (i, k, s, d) in enumerate(var):
        by_task.setdefault(i, []).append(v)
    for i in range(len(tasks)):                                  # one option, one start
        add({v: 1 for v in by_task[i]}, 1, 1)
    for i, t in enumerate(tasks):                                # precedence: start_i - end_p >= 0
        for p in t["preds"]:
            coef = {}
            for v in by_task[i]:
                coef[v] = coef.get(v, 0) + var[v][2]
            for v in by_task[p]:
                coef[v] = coef.get(v, 0) - (var[v][2] + var[v][3])
            add(coef, 0, np.inf)
    for i in range(len(tasks)):                                  # M >= end_i
        coef = {v: -(var[v][2] + var[v][3]) for v in by_task[i]}
        coef[iM] = 1
        add(coef, 0, np.inf)
    for u in range(T):                                           # capacity per slot
        w = {v: 1 for v, (i, k, s, d) in enumerate(var) if core.USES_WORKER[k] and s <= u < s + d}
        r = {v: 1 for v, (i, k, s, d) in enumerate(var) if core.USES_ROBOT[k] and s <= u < s + d}
        if w: add(w, -np.inf, nW)
        if r: add(r, -np.inf, nR)
        if zone:
            wa = {v: 1 for v, (i, k, s, d) in enumerate(var) if k == "W" and s <= u < s + d}
            ra = {v: 1 for v, (i, k, s, d) in enumerate(var) if k == "R" and s <= u < s + d}
            if wa:
                wa[ny + u] = -nW
                add(wa, -np.inf, 0)                                  # W-alone only if a[u] = 1
            if ra:
                ra[ny + u] = nR
                add(ra, -np.inf, nR)                                 # R-alone only if a[u] = 0
    if cap is not None:
        add({v: tasks[i]["opts"][k]["exp"] for v, (i, k, s, d) in enumerate(var)}, -np.inf, cap)

    A = lil_matrix((len(rows), n))
    for r_, coef in enumerate(rows):
        for j, val in coef.items():
            A[r_, j] = val
    integrality = np.ones(n); integrality[iM] = 0
    ub = np.ones(n); ub[iM] = np.inf
    res = milp(c, constraints=LinearConstraint(A.tocsr(), lo, hi), integrality=integrality,
               bounds=Bounds(np.zeros(n), ub), options={"time_limit": time_limit, "disp": False})
    if res.x is None:
        return {"status": "INFEASIBLE" if res.status == 2 else f"no solution ({res.message})"}
    x = res.x
    plan = []
    for i in range(len(tasks)):
        v = max(by_task[i], key=lambda v: x[v])
        _, k, s, d = var[v]
        plan.append((k, s * g, (s + d) * g))
    cost_int = sum(tasks[i]["opts"][k]["cost_int"] for i, (k, _, _) in enumerate(plan))
    makespan = max(e for _, _, e in plan)
    return {"status": "OPTIMAL" if res.status == 0 else f"stopped ({res.message})", "plan": plan,
            "makespan": makespan, "cost_int": cost_int, "obj": cost_int + lam * makespan,
            "exposure": sum(tasks[i]["opts"][k]["exp"] for i, (k, _, _) in enumerate(plan)),
            "variables": n, "constraints": len(rows), "step": g}


def main():
    ap = argparse.ArgumentParser(description="Cross-check allocate.py with a MILP (HiGHS).")
    ap.add_argument("workbook")
    ap.add_argument("--lam", type=float); ap.add_argument("--zone", choices=["yes", "no"])
    ap.add_argument("--cap"); ap.add_argument("--workers", type=int); ap.add_argument("--robots", type=int)
    ap.add_argument("--panels", type=int)
    a = ap.parse_args()
    try:
        settings, weights, tasks = core.read_project(a.workbook, panels=a.panels)
    except core.InputError as e:
        sys.exit(f"Input problem: {e}")
    if a.lam is not None: settings["lam"] = a.lam
    if a.zone: settings["zone"] = a.zone == "yes"
    if a.cap: settings["cap"] = None if a.cap.lower() == "off" else float(a.cap)
    if a.workers is not None: settings["nW"] = a.workers
    if a.robots is not None: settings["nR"] = a.robots
    core.build_costs(settings, weights, tasks)
    r = solve_milp(settings, tasks)
    print(f"HiGHS MILP: {r['status']}")
    if "plan" not in r:
        return
    print(f"  {r['variables']} variables, {r['constraints']} constraints, time step {r['step']} min")
    for t, (k, s, e) in zip(tasks, r["plan"]):
        print(f"  {t['id']:5s} {t['name'][:28]:28s} -> {k:2s} {s:4d}–{e:<4d} min")
    print(f"  Makespan {r['makespan']} min · cost {r['cost_int'] / core.SCALE:.3f} · worker exposure {r['exposure']} min"
          f" · objective {r['obj']}")
    print("Compare the objective with allocate.py (CP-SAT): equal objectives confirm the result.")


if __name__ == "__main__":
    main()
