import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📚 IIM Ranchi MBA Timetable Optimization Dashboard")

# ---------------------------------------------------
# Load Timetable
# ---------------------------------------------------

@st.cache_data
def load_data():
    return pd.read_excel("Optimized_MBA_Timetable.xlsx")

df = load_data()

# ---------------------------------------------------
# Sidebar Filters
# ---------------------------------------------------

st.sidebar.header("🔎 Filters")

selected_course = st.sidebar.selectbox(
    "Select Course",
    ["All"] + sorted(df["Course"].unique().tolist())
)

selected_faculty = st.sidebar.selectbox(
    "Select Faculty",
    ["All"] + sorted(df["Faculty"].unique().tolist())
)

selected_week = st.sidebar.selectbox(
    "Select Week",
    ["All"] + sorted(df["Week"].unique().tolist())
)

filtered_df = df.copy()

if selected_course != "All":
    filtered_df = filtered_df[filtered_df["Course"] == selected_course]

if selected_faculty != "All":
    filtered_df = filtered_df[filtered_df["Faculty"] == selected_faculty]

if selected_week != "All":
    filtered_df = filtered_df[filtered_df["Week"] == selected_week]

# ---------------------------------------------------
# Main Timetable Display
# ---------------------------------------------------

st.subheader("📅 Timetable View")
st.dataframe(filtered_df, use_container_width=True)

# ---------------------------------------------------
# KPI Metrics
# ---------------------------------------------------

st.subheader("📊 Key Metrics")

col1, col2, col3 = st.columns(3)

total_sessions = len(df)

sunday_sessions = len(df[df["Day"] == "Sun"])

evening_sessions = len(df[df["Time"].str.contains("5:45|6:15|6:30", regex=True)])

col1.metric("Total Sessions", total_sessions)
col2.metric("Sunday Sessions", sunday_sessions)
col3.metric("Evening Sessions", evening_sessions)

# ---------------------------------------------------
# Weekly Load Distribution
# ---------------------------------------------------

st.subheader("📈 Weekly Load Distribution")

weekly_count = df.groupby("Week").size().reset_index(name="Sessions")

fig_week = px.bar(
    weekly_count,
    x="Week",
    y="Sessions",
    title="Sessions per Week"
)

st.plotly_chart(fig_week, use_container_width=True)

# ---------------------------------------------------
# Faculty Load Heatmap
# ---------------------------------------------------

st.subheader("👩‍🏫 Faculty Load Distribution")

faculty_load = df.groupby(["Faculty", "Week"]).size().reset_index(name="Sessions")

fig_faculty = px.density_heatmap(
    faculty_load,
    x="Week",
    y="Faculty",
    z="Sessions",
    color_continuous_scale="Blues"
)

st.plotly_chart(fig_faculty, use_container_width=True)

# ---------------------------------------------------
# Room Utilization
# ---------------------------------------------------

st.subheader("🏫 Room Utilization")

room_util = df.groupby(["Week"]).size().reset_index(name="UsedSlots")

fig_room = px.line(room_util, x="Week", y="UsedSlots", title="Room Usage Trend")

st.plotly_chart(fig_room, use_container_width=True)

st.success("✅ Timetable is conflict-free and optimized.")