from datetime import *
import pandas as pd
import os
import re
import matplotlib.pyplot as plt
from io import BytesIO
import streamlit as st
import tempfile

# ─── Helpers (unchanged logic) ───────────────────────────────────────────────

def round_to_quarter(dt):
    minutes = dt.minute
    remainder = minutes % 15
    if remainder < 7.5:
        delta = -remainder
    else:
        delta = 15 - remainder
    return dt + timedelta(minutes=delta, seconds=-dt.second)


def readFile(filePath):
    with open(filePath, mode="rb+") as f:
        contents = f.readlines()
    contents = [i.decode("utf-8").strip() for i in contents]
    return contents


def createGraph(df, attendanceDf):
    fig, ax = plt.subplots()
    ax.plot(df["Time"], df["Attendance"], label="Line")
    ax.set_title("Time-Attendance Plot")

    manual_ticks_x = df["Time"].to_list()
    ax.set_xticks(manual_ticks_x)
    plt.xticks(rotation=45)
    plt.tight_layout()

    y_max = attendanceDf["Attendance"].max()
    x_max = attendanceDf[attendanceDf["Attendance"] == attendanceDf["Attendance"].max()].loc[:, "Time"].to_list()[0]

    ax.scatter(x_max, y_max, color="red", s=80, zorder=3, label="Peak")
    ax.annotate(
        f"(Peak {x_max}, {y_max})",
        xy=(x_max, y_max),
        xytext=(x_max, y_max),
        arrowprops=dict(arrowstyle="<-", color="red")
    )

    img_data = BytesIO()
    fig.savefig(img_data, format='png')
    img_data.seek(0)
    plt.close(fig)
    return img_data


def save_upload(uploaded_file):
    """Save a Streamlit UploadedFile to a temp file and return the path."""
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.flush()
    tmp.close()
    return tmp.name


def process(attendee_path, chat_path=None, Interval=15):
    # ── Read attendee CSV (skip bad header rows) ──────────────────────────────
    cnt = 0
    attendance = None
    while cnt is not None:
        try:
            attendance = pd.read_csv(attendee_path, sep=",", index_col=False, skiprows=cnt)
            cnt = None
        except Exception:
            cnt += 1

    attendance = attendance[(attendance["Join Time"] != "--")].dropna(subset=["Join Time"])
    attendance = attendance.where(lambda x: x["Attended"] == "Yes", other=pd.NA).dropna(subset=["Attended"])
    attendance[["Join Time", "Leave Time"]] = attendance.loc[:, ["Join Time", "Leave Time"]].apply(
        lambda x: pd.to_datetime(x, format="%m/%d/%Y %I:%M:%S %p")
    )

    attendanceDf = pd.DataFrame(columns=["DateTime", "Attendance"])

    minTime = attendance["Join Time"].min()
    maxTime = attendance["Join Time"].max()
    totalDuration = maxTime - minTime
    totalDurationInMinutes = round(totalDuration.seconds / 60)
    startTime = round_to_quarter(minTime)

    progress = st.progress(0, text="Processing attendance data…")
    total_steps = totalDurationInMinutes * 12  # 60/5 = 12 steps per minute
    step = 0

    for min_ in range(0, totalDurationInMinutes, 1):
        for sec in range(0, 60, 5):
            aTime = startTime + timedelta(hours=0, minutes=min_, seconds=sec)
            attendanceDf.loc[len(attendanceDf), ["DateTime", "Attendance"]] = [
                aTime,
                len(attendance[
                    (attendance["Join Time"].dt.time <= aTime.time()) &
                    (attendance["Leave Time"].dt.time >= aTime.time())
                ])
            ]
            step += 1
            progress.progress(min(step / max(total_steps, 1), 1.0), text="Processing attendance data…")

    progress.empty()

    attendanceDf.insert(loc=1, column="Time", value=attendanceDf["DateTime"].apply(lambda x: datetime.strftime(x, "%H:%M")))
    attendanceDf["Date"] = attendanceDf["DateTime"].astype("M8[ns]").dt.date
    attendanceDf = attendanceDf.loc[:, ["Date", "Time", "Attendance"]].groupby(by=["Date", "Time"], as_index=False).max()

    Summary = attendanceDf[
        (pd.to_datetime(attendanceDf.Time, format="%H:%M").dt.minute.isin([_ for _ in range(0, 60, Interval)])) |
        (attendanceDf["Time"] == attendanceDf["Time"].min())
    ]
    Graph = Summary

    attendanceDf.drop_duplicates(subset=attendanceDf.columns, inplace=True)
    Summary = Summary.drop_duplicates(subset=Summary.columns)
    Summary = Summary.reindex(columns=["Date", "Time", "Attendance"])
    attendanceDf = attendanceDf.reindex(columns=["Date", "Time", "Attendance"])

    img_data = createGraph(Graph, attendanceDf)

    # ── Extract topic & panelist from raw file ────────────────────────────────
    contents = readFile(attendee_path)
    topic = None
    panelists = None
    for line in contents[:50]:
        if line.startswith("Topic"):
            topic = contents.index(line) + 1
        if line.startswith("Panelist Details"):
            panelists = contents.index(line) + 2

    topicName = contents[topic].split(",")[0] if topic is not None else "Summary"

    try:
        mentorName = contents[panelists].split(",")[1]
    except Exception:
        mentorName = "Simulive"

    # ── Chat links (optional) ─────────────────────────────────────────────────
    chatDf = None
    if chat_path:
        chatContents = readFile(chat_path)
        lines = []
        for chatLine in chatContents:
            if (chatLine.find(f"From {mentorName}") != -1 or
                    chatLine.find("From Team ") != -1 or
                    chatLine.find("From Anushka ") != -1):
                lines.append(chatContents.index(chatLine) + 1)

        chatLinks = set(
            chatContents[line].replace('"', '').strip()
            for line in lines
            if chatContents[line].find("://") != -1
        )
        if chatLinks:
            chatDf = pd.DataFrame(columns=["Links"], data=chatLinks)["Links"].str.split(r" \r\t\r\t", expand=True)
            if len(chatDf.columns) > 1:
                chatDf = pd.concat(
                    [chatDf.iloc[:, [0]]] + [chatDf.iloc[:, i].dropna() for i in chatDf.columns[1:]],
                    axis=0
                ).drop_duplicates().dropna()
                chatDf.rename(columns={0: "Links"}, inplace=True)

    # ── Write output Excel to BytesIO ─────────────────────────────────────────
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine="xlsxwriter", mode="w") as file:
        workbook = file.book
        worksheet = workbook.add_worksheet("Plot")
        worksheet.insert_image("B2", "plot.png", {"image_data": img_data})

        attendanceDf.to_excel(file, sheet_name="Data", index=False)
        Summary.to_excel(file, sheet_name=topicName[:30], index=False)
        attendanceDf.drop_duplicates(subset=attendanceDf.columns)\
                    .sort_values(by="Attendance", ascending=False)\
                    .head(10)\
                    .to_excel(file, sheet_name="Top 10 Peak Times", index=False)

        if chatDf is not None and len(chatDf) > 0:
            chatDf.to_excel(file, sheet_name="Important Links", index=False)

    output_buffer.seek(0)
    return output_buffer, topicName


# ─── Streamlit UI ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="Attendance Insights", page_icon="📊", layout="centered")
st.title("📊 Attendance Insights Generator")
st.caption("Upload your Zoom attendee sheet to generate an insights report.")

attendee_file = st.file_uploader("Upload Attendee Sheet (.csv)", type=["csv"])

use_10min = st.checkbox("Check for 10 minutes interval, ideal for onboarding/intro session.")
Interval = 10 if use_10min else 15

include_chat = st.checkbox("Include chat file for link extraction.")

chat_file = None
if include_chat:
    chat_file = st.file_uploader("Upload Chat File (.csv / .txt)", type=["csv", "txt"])

if st.button("Generate Report", type="primary"):
    if attendee_file is None:
        st.error("Please upload an attendee sheet to continue.")
    elif include_chat and chat_file is None:
        st.error("Please upload a chat file or uncheck the option.")
    else:
        with st.spinner("Generating insights…"):
            attendee_path = save_upload(attendee_file)
            chat_path = save_upload(chat_file) if chat_file else None

            try:
                output_buffer, topicName = process(attendee_path, chat_path, Interval)

                st.success("✅ Report generated successfully!")

                summary_df = pd.read_excel(output_buffer, sheet_name=topicName[:30])
                output_buffer.seek(0)

                st.subheader("📋 Summary")
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

                st.download_button(
                    label="⬇️ Download Excel Report",
                    data=output_buffer,
                    file_name=f"Insights_{topicName[:30]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred while processing: {e}")
            finally:
                # Clean up temp files
                if os.path.exists(attendee_path):
                    os.remove(attendee_path)
                if chat_path and os.path.exists(chat_path):
                    os.remove(chat_path)
