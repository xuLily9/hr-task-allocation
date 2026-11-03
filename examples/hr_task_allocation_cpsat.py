"""
Human–robot task allocation: a 3-task CP-SAT example with the 7 BWM criteria
=============================================================================

Model (what the code implements):

    min  sum_{i,k} C_ik * x_ik  +  lambda * makespan

    s.t. (1) each task is done exactly once:        sum_k x_ik = 1
         (2) feasibility:                           x_ik does not exist if option k cannot do task i
         (3) precedence:                            start_j >= end_i   for (i -> j)
         (4) resources: one worker, one robot,      no two tasks using the same resource overlap
             a human–robot team (HR) uses both
         (5) safety rules (optional):
             (5a) zone rule: a worker working alone (W) and a robot working alone (R)
                  must not work at the same time in the shared zone
             (5b) cumulative worker exposure in the hazard zone <= cap

    where C_ik = sum_j w_j * s_ijk,  j = C1 ... C7   (the 7 criteria of the BWM questionnaire)
          w_j   = criterion weight (questionnaire Part C, BWM)
          s_ijk = score of option k on task i for criterion j, 0–1, higher = worse
                  (C1 and C2 from case data; C3–C7 from questionnaire Part D)

Execution options (k):
    W  = worker alone
    R  = robot alone
    HR = human–robot team (occupies the worker AND the robot)

Traceability: comments marked [factor] name the factors from the 94-factor
literature table (Factors_classified.xlsx) that each element represents.

Requires: pip install ortools   (version 9.8 or later)
Run:      python hr_task_allocation_cpsat.py
"""

from ortools.sat.python import cp_model

# ---------------------------------------------------------------------------
# 1. DATA  (all numbers are illustrative assumptions, not measured values)
# ---------------------------------------------------------------------------

TASKS = ["T1 lift panel", "T2 align panel", "T3 fix bolts"]
MODES = ["W", "R", "HR"]

# [factor] Team size and composition; worker / agent availability (category C)
USES_WORKER = {"W": True, "R": False, "HR": True}
USES_ROBOT = {"W": False, "R": True, "HR": True}

# Duration (minutes) of each feasible option -> C1 score and interval length.
# A missing entry = option not feasible -> constraint (2).
# [factor] Production and task duration (B -> C1)
# [factor] Precision and assembly tolerances (A): the robot alone cannot align T2
DURATION = {
    ("T1 lift panel", "W"): 60, ("T1 lift panel", "R"): 40, ("T1 lift panel", "HR"): 35,
    ("T2 align panel", "W"): 30,                             ("T2 align panel", "HR"): 25,
    ("T3 fix bolts", "W"): 50,  ("T3 fix bolts", "R"): 40,  ("T3 fix bolts", "HR"): 30,
}

# Cost rate per minute of each option -> C2 score (worker wage, robot rental, both).
# [factor] Production, labour and equipment cost (B -> C2)
COST_RATE = {"W": 0.6, "R": 1.0, "HR": 1.6}

# Expert ratings for C3–C7 (questionnaire Part D), normalised to 0–1, higher = worse.
# Order of each tuple: (W, R, HR); None = option not feasible.
RATINGS = {
    # C3 Safety risk exposure
    #   [factor] Task hazards and risk exposure; Human–robot proximity
    "C3": {"T1 lift panel": (0.8, 0.3, 0.4), "T2 align panel": (0.2, None, 0.3), "T3 fix bolts": (0.2, 0.2, 0.3)},
    # C4 Physical demand and fatigue
    #   [factor] Physical fatigue; Physical task demand and load; Physical strain; ...
    "C4": {"T1 lift panel": (0.9, 0.1, 0.3), "T2 align panel": (0.3, None, 0.2), "T3 fix bolts": (0.5, 0.1, 0.2)},
    # C5 Cognitive workload
    #   [factor] Cognitive strain and fatigue; Attention demand; Spatial-perception demand
    "C5": {"T1 lift panel": (0.4, 0.2, 0.5), "T2 align panel": (0.6, None, 0.5), "T3 fix bolts": (0.3, 0.2, 0.4)},
    # C6 Task complexity
    #   [factor] Task complexity; Measurement demand; Collaboration difficulty
    "C6": {"T1 lift panel": (0.5, 0.4, 0.6), "T2 align panel": (0.7, None, 0.6), "T3 fix bolts": (0.3, 0.3, 0.5)},
    # C7 Standardisation and supervisability (poor fit between task and option = high)
    #   [factor] Task standardisation and supervisability; Task repetitiveness; Human oversight
    "C7": {"T1 lift panel": (0.5, 0.2, 0.3), "T2 align panel": (0.2, None, 0.3), "T3 fix bolts": (0.6, 0.1, 0.3)},
}

# Minutes the worker spends inside the hazard zone under each option -> constraint (5b).
# Robot-alone is not zero: the worker still sets up and inspects.
# [factor] Task hazards and risk exposure (also part of C3: see the note on double use)
EXPOSURE = {
    ("T1 lift panel", "W"): 60, ("T1 lift panel", "R"): 5, ("T1 lift panel", "HR"): 20,
    ("T2 align panel", "W"): 10,                            ("T2 align panel", "HR"): 10,
    ("T3 fix bolts", "W"): 20,  ("T3 fix bolts", "R"): 5,  ("T3 fix bolts", "HR"): 15,
}

# Precedence: T1 must finish before T2 and T3 start -> constraint (3).
# [factor] Task precedence, dependencies and completion status (A)
PRECEDENCE = [("T1 lift panel", "T2 align panel"), ("T1 lift panel", "T3 fix bolts")]

# Criterion weights w_j (in the study: from the BWM questionnaire, Part C). Must sum to 1.
WEIGHTS = {"C1": 0.15, "C2": 0.10, "C3": 0.30, "C4": 0.15, "C5": 0.10, "C6": 0.10, "C7": 0.10}
assert abs(sum(WEIGHTS.values()) - 1) < 1e-9

CRITERIA = {
    "C1": "Task duration", "C2": "Cost", "C3": "Safety risk exposure",
    "C4": "Physical demand and fatigue", "C5": "Cognitive workload",
    "C6": "Task complexity", "C7": "Standardisation and supervisability",
}

FEASIBLE = list(DURATION.keys())
MAX_DURATION = max(DURATION.values())
MAX_COST = max(COST_RATE[k] * DURATION[t, k] for (t, k) in FEASIBLE)


# ---------------------------------------------------------------------------
# 2. SCORES AND COST  C_ik = sum_j w_j * s_ijk
# ---------------------------------------------------------------------------

def scores(task, mode):
    """Normalised scores s_ijk (0–1, higher = worse) for the 7 criteria."""
    s = {
        "C1": DURATION[task, mode] / MAX_DURATION,
        "C2": COST_RATE[mode] * DURATION[task, mode] / MAX_COST,
    }
    idx = MODES.index(mode)
    for c in ("C3", "C4", "C5", "C6", "C7"):
        s[c] = RATINGS[c][task][idx]
    return s


def cost(task, mode):
    s = scores(task, mode)
    return sum(WEIGHTS[c] * s[c] for c in WEIGHTS)


SCALE = 1000  # CP-SAT works with integers, so costs are multiplied by 1000


# ---------------------------------------------------------------------------
# 3. MODEL + SOLVE
# ---------------------------------------------------------------------------

def solve(name, lam, zone_rule=False, exposure_cap=None):
    """lam = weight on makespan, in cost units per minute."""
    m = cp_model.CpModel()
    horizon = sum(max(DURATION[t, k] for (t2, k) in FEASIBLE if t2 == t) for t in TASKS)

    start = {t: m.new_int_var(0, horizon, f"start_{t}") for t in TASKS}
    end = {t: m.new_int_var(0, horizon, f"end_{t}") for t in TASKS}

    x = {}         # x[i,k] = 1 if task i is done by option k
    interval = {}  # interval[i,k] exists only if x[i,k] = 1
    for (t, k) in FEASIBLE:                       # constraint (2): only feasible pairs get a variable
        x[t, k] = m.new_bool_var(f"x_{t}_{k}")
        interval[t, k] = m.new_optional_interval_var(
            start[t], DURATION[t, k], end[t], x[t, k], f"iv_{t}_{k}")

    # (1) each task done exactly once
    for t in TASKS:
        m.add_exactly_one(x[t, k] for k in MODES if (t, k) in x)

    # (3) precedence  [factor] Task precedence, dependencies and completion status
    for (a, b) in PRECEDENCE:
        m.add(start[b] >= end[a])

    # (4) resources: worker and robot can each do one thing at a time
    #     [factor] Task precedence and resource availability; Agent / worker availability
    m.add_no_overlap([interval[t, k] for (t, k) in FEASIBLE if USES_WORKER[k]])
    m.add_no_overlap([interval[t, k] for (t, k) in FEASIBLE if USES_ROBOT[k]])

    # (5a) safety: worker-alone and robot-alone tasks must not run at the same time
    #      [factor] Safety conditions and zone requirements; Workspace suitability and conflicts
    if zone_rule:
        m.add_no_overlap([interval[t, k] for (t, k) in FEASIBLE if k in ("W", "R")])

    # (5b) safety: cumulative worker exposure cap
    #      [factor] Task hazards and risk exposure
    if exposure_cap is not None:
        m.add(sum(EXPOSURE[t, k] * x[t, k] for (t, k) in FEASIBLE) <= exposure_cap)

    # makespan = latest end time
    makespan = m.new_int_var(0, horizon, "makespan")
    m.add_max_equality(makespan, [end[t] for t in TASKS])

    # objective: weighted cost + lambda * makespan
    m.minimize(sum(round(SCALE * cost(t, k)) * x[t, k] for (t, k) in FEASIBLE)
               + round(SCALE * lam) * makespan)

    solver = cp_model.CpSolver()
    status = solver.solve(m)

    print(f"\n=== {name} ===")
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("No feasible allocation.")
        return
    print("Status:", solver.status_name(status))
    total_cost = total_exposure = 0
    for t in TASKS:
        k = next(k for k in MODES if (t, k) in x and solver.value(x[t, k]))
        total_cost += cost(t, k)
        total_exposure += EXPOSURE[t, k]
        print(f"  {t:15s} -> {k:2s}  {solver.value(start[t]):3d}–{solver.value(end[t]):3d} min"
              f"   C = {cost(t, k):.3f}")
    print(f"  Total cost = {total_cost:.3f} | Makespan = {solver.value(makespan)} min"
          f" | Worker exposure = {total_exposure} min")


if __name__ == "__main__":
    print("Criterion weights:", ", ".join(f"{c} {w:.2f}" for c, w in WEIGHTS.items()))
    print("\nScores s_ijk and cost C(i,k):")
    print(f"  {'task':15s} {'opt':3s} " + " ".join(f"{c:>5s}" for c in WEIGHTS) + "      C")
    for (t, k) in FEASIBLE:
        s = scores(t, k)
        print(f"  {t:15s} {k:3s} " + " ".join(f"{s[c]:5.2f}" for c in WEIGHTS) + f"  {cost(t, k):.3f}")

    # A: almost pure weighted cost (tiny lambda only breaks ties in timing)
    solve("A  Weighted cost only", lam=0.001)
    # B: also value a short makespan
    solve("B  Cost + makespan", lam=0.03)
    # C: B + zone rule
    solve("C  Cost + makespan + zone rule", lam=0.03, zone_rule=True)
    # D: C + worker exposure cap of 40 min
    solve("D  C + exposure cap 40 min", lam=0.03, zone_rule=True, exposure_cap=40)
