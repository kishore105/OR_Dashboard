# IIM Ranchi MBA OR Timetable – Streamlit Dashboard

**Live OR Dashboard** for the IIM Ranchi MBA Term III Timetable.

## Features

| Tab | Description |
|-----|-------------|
| 📅 Timetable Grid | Interactive weekly grid — filter by week, day, department |
| 📊 Analytics | Charts: load by dept/day/slot, heatmap, section sizes, classroom utilisation |
| 👨‍🏫 Faculty View | Per-faculty timeline + session list; DTI premid/postmid split shown |
| 🎓 Student View | Per-student personal timetable + conflict check |
| ✅ Validation | Full 9-point constraint checklist + OR model explanation |
| ⬇️ Export | Download Excel (.xlsx) or filtered CSV |

## Repo Structure

```
OR_Dashboard/
├── app.py                  ← Streamlit entry point
├── solver.py               ← OR solver (DSatur + rebalancing)
├── WAI_Data.xlsx           ← Course enrollment data (bundled default)
├── requirements.txt
└── .streamlit/
    └── config.toml         ← Theme & server config
```

## Deploy to Streamlit Community Cloud (Free)

1. **Push to GitHub** (already done at `github.com/kishore105/OR_Dashboard`)

2. **Go to** [share.streamlit.io](https://share.streamlit.io)

3. Click **"New app"** → Connect your GitHub account → Select:
   - Repository: `kishore105/OR_Dashboard`
   - Branch: `main`
   - Main file: `app.py`

4. Click **Deploy** — live in ~2 minutes ✅

## Run Locally

```bash
# Clone
git clone https://github.com/kishore105/OR_Dashboard.git
cd OR_Dashboard

# Install dependencies
pip install -r requirements.txt

# Run
streamlit run app.py
```

The app will open at **http://localhost:8501**

## OR Model

**Algorithm:** Two-Pass DSatur Graph Coloring + Student Rebalancing

- Build conflict graph G (nodes = sections, edges = shared students/faculty)
- Pass 1 DSatur → assign recurring slot p1 to each section
- Pass 2 DSatur → assign slot p2, forbidden = own p1 + all neighbors' p1
- Student rebalancing → for residual conflicts, move students between S1/S2 of the same course
- Result: **940/940 sessions, 0 conflicts**

**Special DTI Constraint:**
- Weeks 1–5 → Prof. Rohit Kumar (Pre-Mid)
- Weeks 6–10 → Prof. Rojers Puthur Joseph (Post-Mid)
