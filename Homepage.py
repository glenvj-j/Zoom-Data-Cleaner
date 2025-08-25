import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Zoom Data Cleaner", layout="wide")
st.title("Zoom Data Cleaner")


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
        return pd.DataFrame(), pd.DataFrame()

    # Topic is usually around row 2
    file.seek(0)
    topic_df = pd.read_csv(file, skiprows=2, nrows=1)
    Topic = topic_df['Topic'].iloc[0].replace('iBlooming: ', "")

    # Read actual table
    file.seek(0)
    df_webinar = pd.read_csv(file, skiprows=header_row)

    # Take the latest time to get the date
    Date = df_webinar['Join Time'].iloc[-1][0:10]

    # Find attendee section
    attendee_idx = df_webinar[df_webinar['Attended']=="Attendee Details"].index
    if len(attendee_idx) == 0:
        st.error(f"❌ Could not detect Attendee section in {file.name}")
        return pd.DataFrame(), pd.DataFrame()

    # Panelists section
    df_panelist = df_webinar.iloc[2:int(attendee_idx[0])]
    df_panelist = df_panelist.drop_duplicates(subset=["Email"])
    df_panelist["Role"] = "Panelist"

    # Attendees section
    df_attendee = df_webinar.iloc[int(attendee_idx[0])+2:]
    # Step 2: Find duplicated rows (ignoring some columns)
    duplicated_mask = df_attendee.drop(
        columns=[
            'User Name (Original Name)','Attended','Join Time','Leave Time',
            'Time in Session (minutes)','Is Guest','Country/Region Name'
        ],
        errors="ignore"
    ).duplicated()

    # Step 3: Get indexes of duplicates
    df_attendee_duplicated_index = df_attendee[duplicated_mask].index

    # Step 4: Drop those duplicates from df_attendee
    df_attendee_clean = df_attendee.drop(df_attendee_duplicated_index)


    # df_attendee_clean = df_attendee.drop(
    #     columns=['User Name (Original Name)','Attended','Join Time','Leave Time',
    #              'Time in Session (minutes)','Is Guest','Country/Region Name'],
    #     errors="ignore"
    # ).drop_duplicates()
    df_attendee_clean["Role"] = "Attendee"

    # Merge both
    df_clean = pd.concat([df_panelist, df_attendee_clean], ignore_index=True)

    # Counts
    total_panelist = (df_clean["Role"]=="Panelist").sum()
    total_attendee = (df_clean["Role"]=="Attendee").sum()
    duplicated_data = df_attendee.duplicated().sum()

    # Summary
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

    return new_data, df_clean


def count_meeting_participant(file, excluded_name):
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
        return pd.DataFrame(), pd.DataFrame()

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

    # Clean data
    df_meeting_clean = df_meeting[['Name (original name)','Total duration (minutes)']].drop_duplicates()

    # Panelist vs Attendee
    if excluded_name.strip():
        df_meeting_clean['Role'] = df_meeting_clean['Name (original name)'].apply(
            lambda x: "Panelist" if re.search(excluded_name, str(x), flags=re.IGNORECASE) else "Attendee"
        )
    else:
        df_meeting_clean['Role'] = "Attendee"

    # Counts
    total_panelist = (df_meeting_clean['Role']=="Panelist").sum()
    total_attendee = (df_meeting_clean['Role']=="Attendee").sum()
    duplicated_data = df_meeting[['Name (original name)','Total duration (minutes)']].duplicated().sum()

    # Summary
    new_data = pd.DataFrame([{
        "Date": Date,
        "Topic": Topic,
        "Total_Attendee": total_attendee,
        "Total_Panelist": total_panelist,
        "Total_All": total_attendee + total_panelist,
        "Row_Deleted": duplicated_data,
        "Type": "Meeting"
    }])   

    return new_data, df_meeting_clean


# -----------------------------
# STREAMLIT APP
# -----------------------------
uploaded_files = st.sidebar.file_uploader(
    "Upload Zoom CSV files", 
    type=["csv"], 
    accept_multiple_files=True
)

excluded_name = st.sidebar.text_area(
    "User Name to Exclude (Meeting Only)\n(separate with commas)",
    value='admin, iblooming, interpreter, host'
)
excluded_name = [x.strip() for x in excluded_name.split(",")]
excluded_name = "|".join(excluded_name)

if uploaded_files:
    data_processed_result = pd.DataFrame()
    cleaned_files = {}  # store cleaned DataFrames per file

    for file in uploaded_files:
        filename = file.name
        if "attendee" in filename.lower():
            result, cleaned = count_webinar_participant(file)
        elif "participants" in filename.lower():
            result, cleaned = count_meeting_participant(file, excluded_name)
        else:
            st.warning(f"⚠️ Skipped: {filename} (unknown type)")
            continue

        if not result.empty:
            data_processed_result = pd.concat([data_processed_result, result], ignore_index=True)
            cleaned_files[filename] = cleaned  # save cleaned data per file

    if not data_processed_result.empty:
        st.success("✅ Processing complete!")
        st.dataframe(data_processed_result)

        # Download aggregated summary
        csv = data_processed_result.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Processed Summary CSV",
            data=csv,
            file_name="zoom_processed_result.csv",
            mime="text/csv"
        )

        # Download each cleaned file
        st.subheader("📂 Download Combined Cleaned Data per File")
        for fname, df_clean in cleaned_files.items():
            if not df_clean.empty:
                clean_csv = df_clean.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download cleaned {fname}",
                    data=clean_csv,
                    file_name=f"cleaned_{fname}",
                    mime="text/csv"
                )
else:
    st.info("Upload one or more Zoom CSV files to begin.")
