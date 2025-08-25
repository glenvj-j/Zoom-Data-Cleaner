import streamlit as st
import pandas as pd
import re

st.title("Zoom Data Cleaner")
st.set_page_config(page_title="Zoom Data Cleaner", layout="wide")


# excluded_name = 'admin|iblooming|interpreter|host'


# -----------------------------
# FUNCTIONS
# -----------------------------
def count_webinar_participant(file):
    # Reset file pointer
    file.seek(0)
    lines = file.readlines()
    file.seek(0)

    # Find header row (the one containing "User Name (Original Name)")
    header_row = None
    for i, line in enumerate(lines):
        if b"User Name (Original Name)" in line:
            header_row = i
            break

    if header_row is None:
        st.error(f"❌ Could not detect Webinar header row in {file.name}")
        return pd.DataFrame()

    # Topic is usually around row 2
    file.seek(0)
    topic_df = pd.read_csv(file, skiprows=2, nrows=1)
    Topic = topic_df['Topic'].iloc[0].replace('iBlooming: ', "")

    # Read actual table
    file.seek(0)
    df_webinar = pd.read_csv(file, skiprows=header_row)

    # Take the latest time to get the date
    Date = df_webinar['Join Time'].iloc[-1][0:10]

    # Total Panelist : before "Attendee Details"
    attendee_idx = df_webinar[df_webinar['Attended']=="Attendee Details"].index
    if len(attendee_idx) == 0:
        st.error(f"❌ Could not detect Attendee section in {file.name}")
        return pd.DataFrame()

    df_panelist = df_webinar.iloc[2:int(attendee_idx[0])]
    total_panelist = df_panelist['Email'].drop_duplicates().shape[0]

    # Total Attendee : after "Attendee Details"
    df_attendee = df_webinar.iloc[int(attendee_idx[0])+2:]
    duplicated_data = df_attendee.drop(
        columns=['User Name (Original Name)','Attended','Join Time','Leave Time',
                 'Time in Session (minutes)','Is Guest','Country/Region Name']
    ).duplicated().sum()

    total_attendee = df_attendee.drop(
        columns=['User Name (Original Name)','Attended','Join Time','Leave Time',
                 'Time in Session (minutes)','Is Guest','Country/Region Name']
    ).drop_duplicates().shape[0]

    new_data = pd.DataFrame([{
        "Date": Date,
        "Topic": Topic,
        "Total_Attendee": total_attendee,
        "Total_Panelist": total_panelist,
        "Total_All": total_attendee + total_panelist,
        "Row_Deleted": duplicated_data,
        "Type": "Webinar"
    }])   

    new_data['Date'] = pd.to_datetime(new_data['Date'], errors='coerce')
    new_data['Date'] = new_data['Date'].dt.strftime("%d/%m/%Y")

    return new_data


def count_meeting_participant(file):
    # Reset file pointer
    file.seek(0)
    lines = file.readlines()
    file.seek(0)

    # Find header row (the one that contains "Name (original name)")
    header_row = None
    for i, line in enumerate(lines):
        if b"Name (original name)" in line:
            header_row = i
            break

    if header_row is None:
        st.error(f"❌ Could not detect Meeting header row in {file.name}")
        return pd.DataFrame()

    # Read dataframe starting from the header row
    file.seek(0)
    df_meeting = pd.read_csv(file, skiprows=header_row)

    # Topic is usually row 1 col 1
    file.seek(0)
    Topic = pd.read_csv(file, nrows=1).iloc[0,0]

    # Date comes from "Start time"
    file.seek(0)
    df_meta = pd.read_csv(file)
    Date = pd.to_datetime(df_meta['Start time'].astype('datetime64[ns]')[0]).strftime("%d/%m/%Y")

    duplicated_data = df_meeting[['Name (original name)','Total duration (minutes)']].duplicated().sum()
    df_meeting_clean = df_meeting[['Name (original name)','Total duration (minutes)']].drop_duplicates()

    # Panelist vs Attendee
    # df_meeting_clean['Panelist'] = df_meeting_clean['Name (original name)'].str.contains(excluded_name, flags=re.IGNORECASE)
    if excluded_name.strip():  # only run if not empty
        df_meeting_clean['Panelist'] = df_meeting_clean['Name (original name)'].str.contains(
            excluded_name, flags=re.IGNORECASE, na=False
        )
    else:
        # default: no one is excluded
        df_meeting_clean['Panelist'] = False

    total_panelist = (df_meeting_clean['Panelist']==1).sum()
    total_attendee = (df_meeting_clean['Panelist']==0).sum()

    new_data = pd.DataFrame([{
        "Date": Date,
        "Topic": Topic,
        "Total_Attendee": total_attendee,
        "Total_Panelist": total_panelist,
        "Total_All": total_attendee + total_panelist,
        "Row_Deleted": duplicated_data,
        "Type": "Meeting"
    }])   

    return new_data


# -----------------------------
# STREAMLIT APP
# -----------------------------
uploaded_files = st.sidebar.file_uploader(
    "Upload Zoom CSV files", 
    type=["csv"], 
    accept_multiple_files=True
)

excluded_name = st.sidebar.text_area(r"User Name to Exclude (Meeting Only)",value='admin, iblooming, interpreter, host')
excluded_name = [x.strip() for x in excluded_name.split(",")]
excluded_name = "|".join(excluded_name)

if uploaded_files:
    data_processed_result = pd.DataFrame()

    for file in uploaded_files:
        filename = file.name
        if "attendee" in filename.lower():
            result = count_webinar_participant(file)
        elif "participants" in filename.lower():
            result = count_meeting_participant(file)
        else:
            st.warning(f"⚠️ Skipped: {filename} (unknown type)")
            continue

        if not result.empty:
            data_processed_result = pd.concat([data_processed_result, result], ignore_index=True)

    if not data_processed_result.empty:
        st.success("✅ Processing complete!")
        st.dataframe(data_processed_result)

        # Download button
        csv = data_processed_result.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Processed Data as CSV",
            data=csv,
            file_name="zoom_processed_result.csv",
            mime="text/csv"
        )
else:
    st.info("Upload one or more Zoom CSV files to begin.")
