import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

st.set_page_config(page_title="期末考试监考安排系统", layout="wide")
st.title("📘 期末考试监考安排系统")

# 读取数据
@st.cache_data

def load_data():
    file_path = "exam_invigilation_data.xlsx"
    teachers_df = pd.read_excel(file_path, sheet_name="教师名单")
    rooms_df = pd.read_excel(file_path, sheet_name="教室列表")
    sessions_df = pd.read_excel(file_path, sheet_name="考试时间段")
    schedule_df = pd.read_excel(file_path, sheet_name="监考安排表")
    return teachers_df, rooms_df, sessions_df, schedule_df

teachers_df, rooms_df, sessions_df, schedule_df = load_data()

# 用户输入模块
st.sidebar.header("🧩 安排考试")
exam_time = st.sidebar.selectbox("选择考试时间段：", sessions_df["考试时间段"].unique())
room = st.sidebar.selectbox("选择教室：", rooms_df["教室编号"].unique())
teacher1 = st.sidebar.selectbox("选择监考老师 1：", [""] + teachers_df["教师编号"].tolist())
teacher2 = st.sidebar.selectbox("选择监考老师 2（可选）：", [""] + teachers_df["教师编号"].tolist())

if st.sidebar.button("➕ 添加安排"):
    new_row = {
        "考试时间段": exam_time,
        "教室编号": room,
        "监考老师1": teacher1,
        "监考老师2": teacher2,
        "冲突检查结果": "待检测"
    }
    schedule_df = pd.concat([schedule_df, pd.DataFrame([new_row])], ignore_index=True)

# 冲突检测
conflict_msgs = []
for i, row in schedule_df.iterrows():
    time = row["考试时间段"]
    room_id = row["教室编号"]
    t1 = row["监考老师1"]
    t2 = row["监考老师2"]
    msg = "✅ 正常"
    dup_room = schedule_df[(schedule_df["考试时间段"] == time) & (schedule_df["教室编号"] == room_id)]
    if len(dup_room) > 1:
        msg = "❌ 教室冲突"
    elif t1 and len(schedule_df[(schedule_df["考试时间段"] == time) & (schedule_df["监考老师1"] == t1)]) > 1:
        msg = "❌ 教师1冲突"
    elif t2 and len(schedule_df[(schedule_df["考试时间段"] == time) & (schedule_df["监考老师2"] == t2)]) > 1:
        msg = "❌ 教师2冲突"
    schedule_df.at[i, "冲突检查结果"] = msg

st.subheader("📋 当前监考安排")
st.dataframe(schedule_df, use_container_width=True)

# 教师监考次数统计
def get_teacher_stats(df):
    all_teachers = df["监考老师1"].tolist() + df["监考老师2"].dropna().tolist()
    stats = pd.Series(all_teachers).value_counts().reset_index()
    stats.columns = ["教师编号", "监考次数"]
    return stats.sort_values("教师编号")

st.subheader("📊 教师监考次数统计")
teacher_stats = get_teacher_stats(schedule_df)
st.dataframe(teacher_stats, use_container_width=True)

# 导出安排
def convert_df(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="监考安排")
    return output.getvalue()

st.download_button(
    label="📥 导出监考安排为 Excel",
    data=convert_df(schedule_df),
    file_name='监考安排_导出.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

st.caption("Made for Lifeng ✨ | ChatGPT Custom Streamlit App")
