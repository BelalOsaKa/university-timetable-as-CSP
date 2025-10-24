import streamlit as st
import pandas as pd
import random
from collections import defaultdict
import io
import base64

# Streamlit interface setup
st.set_page_config(page_title="University Timetable Generator", layout="wide")
st.markdown(
    """
    <style>
    /* === GENERAL PAGE STYLE === */
    body {
        background: linear-gradient(135deg, #e3f2fd, #fce4ec);
        color: #333;
        font-family: 'Poppins', sans-serif;
    }

    /* === MAIN TITLE === */
    h1 {
        text-align: center;
        color: #1976d2;
        font-weight: 700;
        font-size: 2.8em;
        margin-bottom: 0.5em;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }

    /* === INFO TEXT === */
    .stMarkdown {
        font-size: 1.1em;
        color: #444;
        text-align: center;
        margin-bottom: 1em;
    }

    /* === FILE UPLOADER === */
    .stFileUploader {
        background-color: #ffffff;
        border: 2px dashed #90caf9;
        border-radius: 15px;
        padding: 20px;
        transition: 0.3s;
    }
    .stFileUploader:hover {
        background-color: #e3f2fd;
        border-color: #42a5f5;
    }

    /* === BUTTONS === */
    .stButton > button {
        background: linear-gradient(135deg, #42a5f5, #64b5f6);
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.6em 1.2em;
        font-size: 1em;
        transition: 0.3s;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #1e88e5, #64b5f6);
        transform: scale(1.05);
    }

    /* === TABLE STYLE === */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0px 2px 8px rgba(0,0,0,0.1);
        background-color: white;
        padding: 10px;
    }

    /* === STATUS MESSAGES === */
    .stSuccess {
        background-color: #c8e6c9 !important;
        color: #2e7d32 !important;
        border-left: 5px solid #2e7d32;
        border-radius: 8px;
    }
    .stError {
        background-color: #ffcdd2 !important;
        color: #c62828 !important;
        border-left: 5px solid #c62828;
        border-radius: 8px;
    }

    /* === DOWNLOAD LINK === */
    a {
        text-decoration: none;
        background: #4caf50;
        color: white !important;
        padding: 0.6em 1em;
        border-radius: 8px;
        display: inline-block;
        transition: 0.3s;
        font-weight: 600;
    }
    a:hover {
        background: #43a047;
        transform: translateY(-2px);
    }

    /* === FOOTER === */
    footer {
        text-align: center;
        margin-top: 40px;
        color: #555;
        font-size: 0.9em;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Page content
st.title("University Timetable Generator")
st.markdown("Upload **Tables.xlsx** to automatically create your timetable. 📚")

uploaded_file = st.file_uploader("Choose Tables.xlsx", type=["xlsx"])

# Utility functions
def clean_str(x):
    return "" if pd.isna(x) else str(x).strip()

def to_int(x, default=0):
    try:
        return int(x)
    except:
        return default

def match_room(ct, rt):
    c, r = (ct or "").lower(), (rt or "").lower()
    return c == r or ("lec" in c and "lec" in r) or ("lab" in c and "lab" in r) or ("project" in c and "project" in r) or "lec" in r

# Load and process Excel
def load_and_process(file):
    try:
        xls = pd.ExcelFile(file)
        sheets = {n.lower(): n for n in xls.sheet_names}
        req = {
            "crs": ["courses"],
            "inst": ["instructors", "instructor"],
            "rm": ["rooms"],
            "ts": ["timeslots", "time slots", "time_slots"],
            "sec": ["sections"],
            "cur": ["curriculum"]
        }
        data = {}
        for k, v in req.items():
            for s in v:
                if s in sheets:
                    data[k] = pd.read_excel(xls, sheets[s])
                    break
            else:
                raise RuntimeError(f"Missing sheet: {k}")

        # Courses
        crs = {clean_str(r.get("course_id", r.get("CourseID", r.iloc[0]))):
               {"type": clean_str(r.get("type", r.get("Type", ""))).lower()}
               for _, r in data["crs"].iterrows()}

        # Instructors
        inst = {clean_str(r.get("instructor_id", r.get("instructors_id", r.get("InstructorID", r.iloc[0])))):
                {"name": clean_str(r.get("name", r.get("Name", ""))),
                 "quals": set(clean_str(r.get("qualifications", "")).replace(";", ",").split(","))}
                for _, r in data["inst"].iterrows()}

        # Rooms
        rms = {clean_str(r.get("room_id", r.get("RoomID", r.iloc[0]))):
               {"type": clean_str(r.get("type", r.get("Type", ""))).lower(),
                "cap": to_int(r.get("capacity", 0))}
               for _, r in data["rm"].iterrows()}

        # Timeslots
        ts = []
        ts_info = {}
        for _, r in data["ts"].iterrows():
            tid = clean_str(r.get("time_slot_id", r.get("TimeSlotID", r.iloc[0])))
            ts.append(tid)
            ts_info[tid] = {
                "day": clean_str(r.get("day", "")),
                "start": clean_str(r.get("start_time", "")),
                "end": clean_str(r.get("end_time", ""))
            }

        # Sections
        secs = [{"id": clean_str(r.get("section_id", r.iloc[0])),
                 "year": to_int(r.get("year", 1)),
                 "num": to_int(r.get("student", r.get("students", 0)))}
                for _, r in data["sec"].iterrows()]

        # Curriculum
        cur = defaultdict(list)
        for _, r in data["cur"].iterrows():
            year = to_int(r.get("year", 1))
            cid = clean_str(r.get("course_id", r.get("CourseID", "")))
            if cid:
                cur[year].append(cid)

        return crs, inst, rms, ts, ts_info, secs, cur
    except Exception as e:
        raise RuntimeError(f"Failed to process Excel file: {str(e)}")

# Lecture class
class Lec:
    def __init__(self, cid, sid, yr, idx, num):
        self.cid = cid
        self.sid = sid
        self.yr = yr
        self.idx = idx
        self.num = num
        self.name = f"{cid}_{sid}_L{idx}"

# Build lectures and domains
def build_lecs(crs, inst, rms, ts, secs, cur):
    lecs = []
    doms = {}
    for s in secs:
        yr, num, sid = s["year"], s["num"], s["id"]
        for cid in cur.get(yr, []):
            ctype = crs.get(cid, {}).get("type", "")
            sess = 2 if "lec" in ctype else 1
            for i in range(sess):
                lec = Lec(cid, sid, yr, i, num)
                lecs.append(lec)
                dom = [(t, r, iid, cid in info["quals"])
                       for t in ts
                       for r, rinfo in rms.items()
                       if match_room(ctype, rinfo["type"]) and rinfo["cap"] >= num
                       for iid, info in inst.items()]
                doms[lec] = dom
    return lecs, doms

# Assign lectures
def assign_lecs(lecs, doms, inst, rms, ts):
    assigns = {}
    used_rt = set()
    used_it = set()
    for lec in sorted(lecs, key=lambda x: -x.num):
        dom = doms.get(lec, [])
        opts = [d for d in dom if d[3]] + [d for d in dom if not d[3]]
        pick = next((d for d in opts if (d[0], d[1]) not in used_rt and (d[0], d[2]) not in used_it), None)
        if not pick:
            pick = random.choice(dom) if dom else (random.choice(ts), random.choice(list(rms.keys())), random.choice(list(inst.keys())), False)
        assigns[lec] = pick
        used_rt.add((pick[0], pick[1]))
        used_it.add((pick[0], pick[2]))
    return assigns

# Generate CSV
def save_csv(assigns, ts_info, inst):
    rows = [
        {
            "Lec": lec.name,
            "Year": lec.yr,
            "Course": lec.cid,
            "Section": lec.sid,
            "Time": a[0],
            "Day": ts_info.get(a[0], {}).get("day", ""),
            "Start": ts_info.get(a[0], {}).get("start", ""),
            "End": ts_info.get(a[0], {}).get("end", ""),
            "Room": a[1],
            "InstID": a[2],
            "InstName": inst.get(a[2], {}).get("name", a[2]),
            "Qualified": bool(a[3])
        }
        for lec, a in assigns.items() if a
    ]
    return pd.DataFrame(rows)

# Process file and show results
if uploaded_file is not None:
    try:
        st.write("Processing file...")
        data = load_and_process(uploaded_file)
        crs, inst, rms, ts, ts_info, secs, cur = data
        st.success(f"Loaded: {len(crs)} courses, {len(inst)} instructors, {len(rms)} rooms, {len(ts)} timeslots, {len(secs)} sections")

        lecs, doms = build_lecs(crs, inst, rms, ts, secs, cur)
        st.write(f"Created {len(lecs)} lectures")

        assigns = assign_lecs(lecs, doms, inst, rms, ts)
        timetable_df = save_csv(assigns, ts_info, inst)

        st.write("### 📅 Generated Timetable:")
        st.dataframe(timetable_df, use_container_width=True)

        csv_buffer = io.StringIO()
        timetable_df.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        b64 = base64.b64encode(csv_data.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="timetable_solution.csv">⬇️ Download Timetable CSV</a>'
        st.markdown(href, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("Please upload Tables.xlsx to start.")

# Footer
st.markdown("<footer>✨ Developed by Belal & Ziad Eldeen — Intelligent Systems Project</footer>", unsafe_allow_html=True)
