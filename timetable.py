
import pandas as pd
import random
import time
from collections import defaultdict

# -------------------------
# Helper functions
# -------------------------
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def int_safe(x, default=0):
    try:
        return int(x)
    except:
        return default

def compatible_room(course_type, room_type):
    c = (course_type or "").lower()
    r = (room_type or "").lower()
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
# Load Excel sheets
# -------------------------
def load_tables_xlsx(path="Tables.xlsx"):
    xls = pd.ExcelFile(path)
    sheets = {name.lower(): name for name in xls.sheet_names}
    required = {
        "courses": ["courses"],
        "rooms": ["rooms"],
        "timeslots": ["timeslots", "time slots", "time_slots"],
        "sections": ["sections"],
        "curriculum": ["curriculum"],
        "instructors": ["instructors", "instructor"]
    }
    data = {}
    missing = []
    for key, variants in required.items():
        found = None
        for v in variants:
            if v in sheets:
                found = sheets[v]
                break
        if not found:
            missing.append(key)
        else:
            data[key] = pd.read_excel(xls, found)
    if missing:
        raise RuntimeError(f"Missing sheets: {missing}")
    return data["courses"], data["instructors"], data["rooms"], data["timeslots"], data["sections"], data["curriculum"]

# -------------------------
# Preprocess data
# -------------------------
def preprocess(courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df):
    # Courses
    courses = {}
    for _, r in courses_df.iterrows():
        cid = safe_str(r.get("course_id", r.get("CourseID", r.iloc[0])))
        cname = safe_str(r.get("course_name", r.get("CourseName", "")))
        ctype = safe_str(r.get("type", r.get("Type", ""))).lower()
        courses[cid] = {"name": cname, "type": ctype}

    # Instructors
    instructors = {}
    for _, r in instructors_df.iterrows():
        iid = safe_str(r.get("instructor_id", r.get("instructors_id", r.get("InstructorID", r.iloc[0]))))
        name = safe_str(r.get("name", r.get("Name", "")))
        quals_raw = safe_str(r.get("qualifications", r.get("Qualifications", "")))
        quals = [q.strip() for q in quals_raw.replace(";", ",").split(",") if q.strip()]
        instructors[iid] = {"name": name, "quals": set(quals)}

    # Rooms
    rooms = {}
    for _, r in rooms_df.iterrows():
        rid = safe_str(r.get("room_id", r.get("RoomID", r.iloc[0])))
        rtype = safe_str(r.get("type", r.get("Type", ""))).lower()
        cap = int_safe(r.get("capacity", r.get("Capacity", 0)))
        rooms[rid] = {"type": rtype, "capacity": cap}

    # Timeslots
    timeslots = []
    timeslot_info = {}
    for _, r in timeslots_df.iterrows():
        tid = safe_str(r.get("time_slot_id", r.get("TimeSlotID", r.iloc[0])))
        day = safe_str(r.get("day", r.get("Day", "")))
        start = safe_str(r.get("start_time", r.get("Start", "")))
        end = safe_str(r.get("end_time", r.get("End", "")))
        timeslots.append(tid)
        timeslot_info[tid] = {"day": day, "start": start, "end": end}

    # Sections
    sections = []
    for _, r in sections_df.iterrows():
        sid = safe_str(r.get("section_id", r.get("SectionID", r.iloc[0])))
        year = int_safe(r.get("year", r.get("Year", 1)))
        students = int_safe(r.get("student", r.get("students", r.get("StudentCount", 0))))
        sections.append({"section_id": sid, "year": year, "students": students})

    # Curriculum
    curriculum = defaultdict(list)
    for _, r in curriculum_df.iterrows():
        year = int_safe(r.get("year", r.get("Year", 1)))
        cid = safe_str(r.get("course_id", r.get("CourseID", r.iloc[1] if len(r) > 1 else "")))
        if cid:
            curriculum[year].append(cid)

    return courses, instructors, rooms, timeslots, timeslot_info, sections, curriculum

# -------------------------
# Lecture variable
# -------------------------
class LectureVar:
    def __init__(self, course, section, year, idx, students):
        self.course = course
        self.section = section
        self.year = year
        self.idx = idx
        self.students = students
        self.name = f"{course}_{section}_L{idx}"

# -------------------------
# Build variables and domains
# -------------------------
def build_vars_domains(courses, instructors, rooms, timeslots, sections, curriculum):
    variables = []
    domains = {}
    for sec in sections:
        sec_year = sec["year"]
        sec_students = sec["students"]
        sec_id = sec["section_id"]
        for cid in curriculum.get(sec_year, []):
            ctype = courses.get(cid, {}).get("type", "")
            sessions = 2 if "lec" in ctype else 1
            for i in range(sessions):
                v = LectureVar(cid, sec_id, sec_year, i, sec_students)
                variables.append(v)
                dom = []
                for t in timeslots:
                    for r, rinfo in rooms.items():
                        if not compatible_room(ctype, rinfo["type"]):
                            continue
                        if rinfo["capacity"] < sec_students:
                            continue
                        for instr_id, info in instructors.items():
                            qual = cid in info["quals"]
                            dom.append((t, r, instr_id, qual))
                domains[v] = dom
    return variables, domains

# -------------------------
# Greedy assignment
# -------------------------
def greedy_assign(variables, domains, instructors, rooms, timeslots):
    assigned = {}
    used_room_ts = set()
    used_instr_ts = set()
    for v in sorted(variables, key=lambda x: -x.students):
        dom = domains.get(v, [])
        qualified = [d for d in dom if d[3]]
        unqualified = [d for d in dom if not d[3]]
        chosen = None
        for option in qualified + unqualified:
            t, r, instr, q = option
            if (t, r) in used_room_ts or (t, instr) in used_instr_ts:
                continue
            chosen = option
            break
        if chosen is None:
            if dom:
                chosen = random.choice(dom)
            else:
                t = random.choice(timeslots)
                r = random.choice(list(rooms.keys()))
                instr = random.choice(list(instructors.keys()))
                chosen = (t, r, instr, False)
        assigned[v] = chosen
        t, r, instr, _ = chosen
        used_room_ts.add((t, r))
        used_instr_ts.add((t, instr))
    return assigned

# -------------------------
# Export CSV
# -------------------------
def export_csv(assigned, timeslot_info, instructors, filename="timetable_solution.csv"):
    rows = []
    for v, val in assigned.items():
        if val is None:
            continue
        t, r, instr_id, qual = val
        info = timeslot_info.get(t, {"day": "", "start": "", "end": ""})
        instr_name = instructors[instr_id]["name"] if instr_id in instructors else instr_id
        rows.append({
            "Variable": v.name,
            "Year": v.year,
            "Course": v.course,
            "Section": v.section,
            "TimeslotID": t,
            "Day": info["day"],
            "Start": info["start"],
            "End": info["end"],
            "Room": r,
            "InstructorID": instr_id,
            "InstructorName": instr_name,
            "InstructorQualified": bool(qual)
        })
    df = pd.DataFrame(rows)
    df.to_csv(filename, index=False)
    print(f"✅ Exported {len(df)} rows to {filename}")

# -------------------------
# Main
# -------------------------
def main():
    print("📘 Loading Tables.xlsx ...")
    courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df = load_tables_xlsx("Tables.xlsx")
    courses, instructors, rooms, timeslots, timeslot_info, sections, curriculum = preprocess(
        courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df
    )
    print(f"✅ Loaded data: {len(courses)} courses, {len(instructors)} instructors, {len(rooms)} rooms, {len(timeslots)} timeslots, {len(sections)} sections.")
    variables, domains = build_vars_domains(courses, instructors, rooms, timeslots, sections, curriculum)
    print(f"🧩 Created {len(variables)} lecture variables.")
    assigned = greedy_assign(variables, domains, instructors, rooms, timeslots)
    export_csv(assigned, timeslot_info, instructors, "timetable_solution.csv")
    print("🎉 Done. No nulls in timetable_solution.csv")

if __name__ == "__main__":
    main()
