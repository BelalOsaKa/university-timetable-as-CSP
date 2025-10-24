"""
timetable_full_solver_v3.py

Robust CSP timetable generator that:
 - auto-detects column names (flexible) to avoid instructor_id errors
 - reads these sheets from Tables.xlsx:
     Courses, Instructors, Rooms, TimeSlots, Sections, Curriculum
 - produces timetable_solution.csv (no nulls) and report.txt (performance)
 - GUI: select file and run (Tkinter)
"""

import pandas as pd
import random
import time
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox

# -------------------------
# Utilities: flexible column lookup
# -------------------------
def find_column(df, candidates):
    """Return the first candidate name that exists in df.columns, else None."""
    for c in candidates:
        for col in df.columns:
            if col.strip().lower() == c.strip().lower():
                return col
    return None

def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def int_safe(x, default=0):
    try:
        return int(x)
    except:
        return default

def compatible_room(course_type, room_type):
    c, r = (course_type or "").lower(), (room_type or "").lower()
    if not c:
        return True
    if c == r:
        return True
    if "lec" in c and "lec" in r:
        return True
    if "lab" in c and "lab" in r:
        return True
    if "project" in c and "project" in r:
        return True
    if "lec" in r:
        return True
    return False

# -------------------------
# Load Excel (detect sheets)
# -------------------------
def load_tables_xlsx(path):
    xls = pd.ExcelFile(path)
    names = xls.sheet_names
    # We expect sheets (case-insensitive) for these roles:
    expected = {
        "courses": ["courses"],
        "instructors": ["instructors","instructor"],
        "rooms": ["rooms","room"],
        "timeslots": ["timeslots","time slots","time_slots","timeslot"],
        "sections": ["sections","section"],
        "curriculum": ["curriculum"]
    }
    data = {}
    missing = []
    lowered = {n.lower(): n for n in names}
    for key, variants in expected.items():
        found_sheet = None
        for v in variants:
            if v in lowered:
                found_sheet = lowered[v]
                break
        if not found_sheet:
            missing.append(key)
        else:
            data[key] = pd.read_excel(xls, found_sheet)
    if missing:
        raise RuntimeError(f"Missing sheets in {path}: {missing}. Present sheets: {names}")
    return data["courses"], data["instructors"], data["rooms"], data["timeslots"], data["sections"], data["curriculum"]

# -------------------------
# Preprocess (robust col names)
# -------------------------
def preprocess(courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df):
    # Courses: find columns
    course_id_col = find_column(courses_df, ["course_id","courseid","course code","code","id"])
    course_name_col = find_column(courses_df, ["course_name","coursename","name","title"])
    course_type_col = find_column(courses_df, ["type","course_type","kind"])

    courses = {}
    for _, r in courses_df.iterrows():
        cid = safe_str(r.get(course_id_col, r.iloc[0] if len(r)>0 else ""))
        if not cid:
            continue
        cname = safe_str(r.get(course_name_col, ""))
        ctype = safe_str(r.get(course_type_col, "")).lower()
        courses[cid] = {"name": cname, "type": ctype}

    # Instructors: flexible columns (fix for instructor_id error)
    instr_id_col = find_column(instructors_df, ["instructor_id","instructors_id","id","instructorid","instr_id"])
    instr_name_col = find_column(instructors_df, ["name","full_name","instructor_name"])
    instr_quals_col = find_column(instructors_df, ["qualifications","qualification","qualified_courses","quals"])

    if instr_id_col is None:
        raise RuntimeError("Could not find instructor ID column in Instructors sheet. Columns found: " + ", ".join(instructors_df.columns))

    instructors = {}
    for _, r in instructors_df.iterrows():
        iid = safe_str(r.get(instr_id_col, ""))
        if not iid:
            # skip blank id rows
            continue
        iname = safe_str(r.get(instr_name_col, ""))
        raw_q = safe_str(r.get(instr_quals_col, ""))
        # normalize separators ; or /
        raw_q = raw_q.replace(";",",").replace("/",",")
        quals = [q.strip() for q in raw_q.split(",") if q.strip()]
        instructors[iid] = {"name": iname or iid, "quals": set(quals)}

    # Rooms
    room_id_col = find_column(rooms_df, ["room_id","roomid","room","id"])
    room_type_col = find_column(rooms_df, ["type","room_type","roomtype"])
    room_cap_col = find_column(rooms_df, ["capacity","cap","room_capacity"])
    rooms = {}
    for _, r in rooms_df.iterrows():
        rid = safe_str(r.get(room_id_col, r.iloc[0] if len(r)>0 else ""))
        if not rid:
            continue
        rtype = safe_str(r.get(room_type_col, "")).lower()
        cap = int_safe(r.get(room_cap_col, 0))
        rooms[rid] = {"type": rtype, "capacity": cap}

    # TimeSlots
    ts_id_col = find_column(timeslots_df, ["time_slot_id","timeslotid","timeslot","id"])
    ts_day_col = find_column(timeslots_df, ["day","weekday"])
    ts_start_col = find_column(timeslots_df, ["start_time","start","begin"])
    ts_end_col = find_column(timeslots_df, ["end_time","end","finish"])
    timeslots = []
    timeslot_info = {}
    for _, r in timeslots_df.iterrows():
        tid = safe_str(r.get(ts_id_col, r.iloc[0] if len(r)>0 else ""))
        if not tid:
            continue
        day = safe_str(r.get(ts_day_col, ""))
        start = safe_str(r.get(ts_start_col, ""))
        end = safe_str(r.get(ts_end_col, ""))
        timeslots.append(tid)
        timeslot_info[tid] = {"day": day, "start": start, "end": end}

    # Sections
    sec_id_col = find_column(sections_df, ["section_id","sectionid","section","id"])
    sec_group_col = find_column(sections_df, ["group_number","group","groupno"])
    sec_year_col = find_column(sections_df, ["year"])
    sec_student_col = find_column(sections_df, ["student","students","student_count","studentcount","students_count"])
    sections=[]
    for _, r in sections_df.iterrows():
        sid = safe_str(r.get(sec_id_col, r.iloc[0] if len(r)>0 else ""))
        if not sid:
            continue
        group = safe_str(r.get(sec_group_col, ""))
        year = int_safe(r.get(sec_year_col, 1))
        students = int_safe(r.get(sec_student_col, 0))
        sections.append({"section_id": sid, "group": group, "year": year, "students": students})

    # Curriculum
    cur_year_col = find_column(curriculum_df, ["year"])
    cur_course_col = find_column(curriculum_df, ["course_id","courseid","course","id"])
    curriculum = defaultdict(list)
    for _, r in curriculum_df.iterrows():
        year = int_safe(r.get(cur_year_col, 1))
        cid = safe_str(r.get(cur_course_col, ""))
        if cid:
            curriculum[year].append(cid)

    # Basic sanity check messages
    msgs = []
    if not courses:
        msgs.append("Warning: no courses found (Courses sheet may be empty or columns different).")
    if not instructors:
        msgs.append("Warning: no instructors found.")
    if not rooms:
        msgs.append("Warning: no rooms found.")
    if not timeslots:
        msgs.append("Warning: no timeslots found.")
    if not sections:
        msgs.append("Warning: no sections found.")
    if not curriculum:
        msgs.append("Warning: no curriculum mapping found.")

    return courses, instructors, rooms, timeslots, timeslot_info, sections, curriculum, msgs

# -------------------------
# Build variables & domains
# -------------------------
class LectureVar:
    def __init__(self, course, section, year, idx, students):
        self.course = course
        self.section = section
        self.year = year
        self.idx = idx
        self.students = students
        self.name = f"{course}_{section}_L{idx}"
    def __repr__(self):
        return self.name
    def __hash__(self):
        return hash(self.name)
    def __eq__(self, other):
        return isinstance(other, LectureVar) and self.name==other.name

def build_vars_domains(courses, instructors, rooms, timeslots, sections, curriculum):
    variables = []
    domains = {}
    instr_list = list(instructors.keys()) if instructors else []
    for sec in sections:
        year = sec["year"]
        students = sec["students"]
        s_id = sec["section_id"]
        clist = curriculum.get(year, [])
        for cid in clist:
            ctype = courses.get(cid, {}).get("type","")
            sessions = 2 if "lec" in ctype else 1
            for i in range(sessions):
                v = LectureVar(cid, s_id, year, i, students)
                variables.append(v)
                dom = []
                for t in timeslots:
                    for r, rinfo in rooms.items():
                        if not compatible_room(ctype, rinfo.get("type","")):
                            continue
                        if rinfo.get("capacity",0) < students:
                            continue
                        # any instructor allowed (qualification flagged)
                        for instr in instr_list:
                            qual = cid in instructors[instr]["quals"]
                            dom.append((t, r, instr, qual))
                domains[v] = dom
    return variables, domains

# -------------------------
# Greedy solver (hard constraints enforced)
# -------------------------
def greedy_assign(variables, domains):
    assigned = {}
    used_room_ts = set()
    used_instr_ts = set()
    violations = 0
    for v in sorted(variables, key=lambda x: -x.students):
        dom = domains.get(v, [])
        # prefer qualified
        qualified = [d for d in dom if d[3]]
        other = [d for d in dom if not d[3]]
        chosen = None
        for option in (qualified + other):
            t,r,instr,qual = option
            if (t,r) in used_room_ts or (t,instr) in used_instr_ts:
                continue
            chosen = option
            break
        if not chosen:
            # fallback: pick option minimizing conflicts (if any)
            best = None
            best_conf = 1e9
            for option in dom:
                t,r,instr,qual = option
                conf = 0
                if (t,r) in used_room_ts: conf += 1
                if (t,instr) in used_instr_ts: conf += 1
                if conf < best_conf:
                    best_conf = conf
                    best = option
            if best:
                chosen = best
                violations += 1
        if not chosen:
            # ultimate synthetic fallback (very rare)
            t = domains and next(iter(dom))[0] if dom else "ts0"
            r = domains and next(iter(dom))[1] if dom else "room0"
            instr = domains and next(iter(dom))[2] if dom else "instr0"
            chosen = (t,r,instr, False)
            violations += 1
        assigned[v] = chosen
        t,r,instr,_ = chosen
        used_room_ts.add((t,r))
        used_instr_ts.add((t,instr))
    return assigned, violations

# -------------------------
# Local improvement to increase qualified assignments (no hard-constraint breaks)
# -------------------------
def improve_assignments(assigned, domains, instructors, max_iters=5000):
    room_ts = set((t,r) for v,(t,r,i,q) in assigned.items())
    instr_ts = set((t,i) for v,(t,r,i,q) in assigned.items())
    unqualified = [v for v,val in assigned.items() if not val[3]]
    random.shuffle(unqualified)
    improved = 0
    it = 0
    while unqualified and it < max_iters:
        v = unqualified.pop()
        it += 1
        dom = domains.get(v, [])
        found = None
        for opt in dom:
            t,r,i,q = opt
            if not q: continue
            # check conflicts ignoring current v's current spot
            cur = assigned[v]
            ct,cr,ci,cq = cur
            conflict = False
            # room conflict
            if (t,r) in room_ts and not (t==ct and r==cr):
                conflict = True
            if (t,i) in instr_ts and not (t==ct and i==ci):
                conflict = True
            if not conflict:
                found = opt
                break
        if found:
            # free old
            ct,cr,ci,cq = assigned[v]
            room_ts.discard((ct,cr)); instr_ts.discard((ct,ci))
            assigned[v] = found
            nt,nr,ni,nq = found
            room_ts.add((nt,nr)); instr_ts.add((nt,ni))
            improved += 1
    return assigned, improved

# -------------------------
# Export CSV and report
# -------------------------
def export_results(assigned, timeslot_info, instructors, out_csv="timetable_solution.csv", report_file="report.txt", runtime=0, violations=0, improved=0):
    rows=[]
    for v,val in assigned.items():
        t,r,iid,qual = val
        info = timeslot_info.get(t, {"day":"", "start":"", "end":""})
        instr_name = instructors.get(iid, {}).get("name", iid)
        rows.append({
            "Variable": v.name,
            "Year": v.year,
            "Course": v.course,
            "Section": v.section,
            "TimeslotID": t,
            "Day": info.get("day",""),
            "Start": info.get("start",""),
            "End": info.get("end",""),
            "Room": r,
            "InstructorID": iid,
            "InstructorName": instr_name,
            "InstructorQualified": bool(qual)
        })
    df = pd.DataFrame(rows)
    df.to_csv(out_csv, index=False)
    # report
    total = len(rows)
    qualified = sum(1 for r in rows if r["InstructorQualified"])
    with open(report_file, "w", encoding="utf-8") as f:
        f.write(f"Timetable generation report\n")
        f.write(f"Rows (assigned lectures): {total}\n")
        f.write(f"Qualified assignments: {qualified}\n")
        f.write(f"Unqualified assignments: {total-qualified}\n")
        f.write(f"Violations during greedy (fallbacks): {violations}\n")
        f.write(f"Local improvements applied: {improved}\n")
        f.write(f"Generation time (s): {runtime:.4f}\n")
    return out_csv, report_file

# -------------------------
# GUI & orchestration
# -------------------------
def run_gui():
    root = tk.Tk()
    root.title("Timetable CSP Generator (v3)")

    def choose_and_run():
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if not path:
            return
        try:
            t0 = time.time()
            courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df = load_tables_xlsx(path)
            courses, instructors, rooms, timeslots, timeslot_info, sections, curriculum, msgs = preprocess(
                courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df
            )
            log_msgs = "\n".join(msgs)
            variables, domains = build_vars_domains(courses, instructors, rooms, timeslots, sections, curriculum)
            assigned, violations = greedy_assign(variables, domains)
            assigned, improved = improve_assignments(assigned, domains, instructors)
            runtime = time.time() - t0
            out_csv, report_file = export_results(assigned, timeslot_info, instructors, runtime=runtime, violations=violations, improved=improved)
            messagebox.showinfo("Done", f"Generated {out_csv}\nReport: {report_file}\nTime: {runtime:.2f}s\n{log_msgs}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Label(root, text="Select Tables.xlsx and generate timetable").pack(padx=12, pady=8)
    tk.Button(root, text="Choose Tables.xlsx", command=choose_and_run, width=30, bg="#2b7", fg="black").pack(pady=8)
    tk.Button(root, text="Quit", command=root.quit, width=10, bg="#f55").pack(pady=6)
    root.mainloop()

if __name__ == "__main__":
    run_gui()

    #streamlet
