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


def process(attendee_path, chat_path=[], Interval=15):
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

    if Summary.iloc[0,2] == 0:
        Summary.loc[Summary.iloc[0, 2], "Attendance"] = attendanceDf[attendanceDf["Attendance"] != 0].iloc[0, 2]

    Summary["Date"] = Summary["Date"].astype(str)
    Summary["Time"] = pd.to_datetime(Summary["Time"]).dt.strftime("%I:%M %p").astype(str)
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

    if len(chat_path) != 0:
        chats = []
        for filename in chat_path:
            with open(filename, "rb+") as f:
                chats.extend(f.readlines())

        data = pd.DataFrame(columns=["TimeStamp", "Comments"])
        for comments in chats:
            ValidComment = comments.decode("utf-8")
            if ValidComment.find("panelists:") > -1 or ValidComment.find(" Everyone:") > -1 or ValidComment.find("(direct message)") > -1:
                data.loc[len(data), "TimeStamp"] = ValidComment
            else:
                data.loc[len(data)-1, "Comments"] = ValidComment

        data["Time"] = data.TimeStamp.str.split(" ", n=1, expand=True)[0]
        data["Info"] = data.TimeStamp.str.split(" ", n=1, expand=True)[1]
        data["From"] = data["Info"].str.split(" to ", n=1, expand=True)[0]
        data["To"]  = data["Info"].str.split(" to ", n=1, expand=True)[1]
        data["From"] = data["From"].str.replace("From", "").str.strip()
        data["To"] = data["To"].str.replace(":", "").str.strip()
        data["Comments"] = data["Comments"].str.strip()
        data["To"] = data["To"].apply(lambda x: x.replace(", Hosts and panelists", "").replace(", host and panelists", "") if x.find(",") > -1 else x )
        data = data.loc[:, ["Time", "From", "To", "Comments"]]

        chatDf = data[(data["From"].str.lower().str.contains("team be10x") | data["From"].str.lower().str.contains("anushka")) \
                    & data["Comments"].str.contains("://")]

        chatDf = chatDf[chatDf["Comments"].str.contains(r'^(https://)', regex=True)]
        chatDf = chatDf.drop_duplicates(subset= "Comments")
        chatDf = chatDf[["Time","Comments" ]]
        chatDf["Time"] = pd.to_datetime(chatDf["Time"]).dt.strftime("%I:%M:%S %p").astype(str)
        
        chatDf.sort_values(by="Time", ascending=True, inplace=True)

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
    chat_file = st.file_uploader("Upload Chat File (.txt)", type=["txt"], accept_multiple_files=True)

if st.button("Generate Report", type="primary"):
    if attendee_file is None:
        st.error("Please upload an attendee sheet to continue.")
    elif include_chat and chat_file is None:
        st.error("Please upload a chat file or uncheck the option.")
    else:
        attendee_path = None
        chat_paths = []
        
        with st.spinner("Generating insights…"):
            try:
                attendee_path = save_upload(attendee_file)
                if len(chat_file)>0:
                    chat_paths = [save_upload(cf) for cf in chat_file]
            
                output_buffer, topicName = process(attendee_path, chat_paths if len(chat_paths)>0 else [], Interval)

                st.success("✅ Report generated successfully!")

                summary_df = pd.read_excel(output_buffer, sheet_name=topicName[:30])
                output_buffer.seek(0)

                chat_df = pd.read_excel(output_buffer, sheet_name="Important Links")
                output_buffer.seek(0)
                
                #st.subheader(f"📋 {topicName}")
                st.code(topicName, language=None)
                
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

                st.subheader("📋 Chat Links")
                st.dataframe(chat_df, use_container_width=True, hide_index=True)
                
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
                if len(chat_paths) != 0:
                    for cp in chat_paths:
                        if os.path.exists(cp):
                            os.remove(cp)
