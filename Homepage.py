import streamlit as st
import pandas as pd
import re
import io
import zipfile
from datetime import datetime

# Get the current date and time
now_utc = datetime.now(timezone.utc)         # current UTC time
now = now_utc.astimezone(timezone(timedelta(hours=7))) #WIB

# Format the date and time to display only until minutes
formatted_datetime = now.strftime("%Y-%m-%d %H:%M")


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

    # Find header row
    header_row = None
    for i, line in enumerate(lines):
        if b"User Name (Original Name)" in line:
            header_row = i
            break

    if header_row is None:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Topic
    file.seek(0)
    topic_df = pd.read_csv(file, skiprows=2, nrows=1)
    Topic = topic_df['Topic'].iloc[0].replace('iBlooming: ', "")

    # Read actual table
    file.seek(0)
    df_webinar = pd.read_csv(file, skiprows=header_row)

    # Take latest time as date
    Date = df_webinar['Join Time'].iloc[-1][0:10]

    # Find attendee section
    attendee_idx = df_webinar[df_webinar['Attended']=="Attendee Details"].index
    if len(attendee_idx) == 0:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Panelists section
    df_panelist = df_webinar.iloc[2:int(attendee_idx[0])]
    df_panelist = df_panelist.drop_duplicates(subset=["Email"])
    df_panelist["Role"] = "Panelist"

    # Attendees section
    df_attendee = df_webinar.iloc[int(attendee_idx[0])+2:]
    duplicated_mask = df_attendee.drop(
        columns=['User Name (Original Name)','Attended','Join Time','Leave Time',
                 'Time in Session (minutes)','Is Guest','Country/Region Name'],
        errors="ignore"
    ).duplicated()
    df_attendee_clean = df_attendee[~duplicated_mask].copy()
    df_attendee_clean["Role"] = "Attendee"

    # Merge both
    df_clean = pd.concat([df_panelist, df_attendee_clean], ignore_index=True)

    # Counts
    total_panelist = (df_clean["Role"]=="Panelist").sum()
    total_attendee = (df_clean["Role"]=="Attendee").sum()
    duplicated_data = df_attendee.duplicated().sum()

    # Country pivot
    df_country = df_clean[['Email','Country/Region Name']].dropna()
    df_t = df_country.pivot_table(
        index=[],
        columns="Country/Region Name",
        values="Email",
        aggfunc="count",
        fill_value=0
    )
    df_t.columns.name = None
    df_t['Date'] = Date
    df_t['Topic'] = Topic
    df_t = df_t[['Date','Topic'] + [c for c in df_t.columns if c not in ['Date','Topic']]]

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

    return new_data, df_clean, df_t


def count_meeting_participant(file, excluded_name):
    file.seek(0)
    lines = file.readlines()
    file.seek(0)

    # Find header row
    header_row = None
    for i, line in enumerate(lines):
        if b"Name (original name)" in line:
            header_row = i
            break

    if header_row is None:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Read dataframe starting from header row
    file.seek(0)
    df_meeting = pd.read_csv(file, skiprows=header_row)

    # Topic
    file.seek(0)
    Topic = pd.read_csv(file, nrows=1).iloc[0,0]

    # Date from "Start time"
    file.seek(0)
    df_meta = pd.read_csv(file)
    Date = pd.to_datetime(df_meta['Start time'].astype('datetime64[ns]')[0]).strftime("%Y-%m-%d")

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

    # Fake country pivot (all 0)
    df_t = pd.DataFrame([{"Date": Date, "Topic": Topic}])

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

    return new_data, df_meeting_clean, df_t


def clean_email_level(file):
    """Build attendee-level CSV (unique per Email–Topic–Date) from webinar files."""
    # Find header row
    file.seek(0)
    lines = file.readlines()
    file.seek(0)
    header_row = None
    for i, line in enumerate(lines):
        if b"User Name (Original Name)" in line:
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()

    # Read full table
    file.seek(0)
    df = pd.read_csv(file, skiprows=header_row)

    # Topic & attendee section
    file.seek(0)
    topic_df = pd.read_csv(file, skiprows=2, nrows=1)
    Topic = topic_df['Topic'].iloc[0].replace('iBlooming: ', "")

    attendee_anchor = df[df['Attended'] == "Attendee Details"]
    if attendee_anchor.empty:
        return pd.DataFrame()

    df_attendee = df.iloc[int(attendee_anchor.index[0]) + 2 :].copy()

    # Date (from last join time)
    Date = str(df_attendee['Join Time'].iloc[-1])[:10]

    # ---- DEDUP ----
    dedup_mask = df_attendee.drop(
        columns=[
            'User Name (Original Name)','Attended','Join Time','Leave Time',
            'Time in Session (minutes)','Is Guest','Country/Region Name'
        ],
        errors="ignore"
    ).duplicated(keep='first')
    df_attendee_clean = df_attendee.loc[~dedup_mask].copy()

    # Keep only relevant columns
    keep_cols = ['User Name (Original Name)', 'Email']
    if 'Country/Region Name' in df_attendee_clean.columns:
        keep_cols.append('Country/Region Name')
    keep_cols = [c for c in keep_cols if c in df_attendee_clean.columns]

    if 'Email' not in keep_cols:
        return pd.DataFrame()

    out = df_attendee_clean[keep_cols].copy()
    out['Email'] = out['Email'].astype(str).str.strip()
    out = out[out['Email'].ne('') & out['Email'].ne('nan')]

    out['Topic'] = Topic
    out['Date'] = Date

    out = out.drop_duplicates(subset=['Email', 'Topic', 'Date'], keep='first')

    return out


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
    data_summary = pd.DataFrame()
    data_email = pd.DataFrame()
    data_country = pd.DataFrame()
    processed_files = set()   # prevent duplicates

    for file in uploaded_files:
        filename = file.name
        if filename in processed_files:
            st.warning(f"⚠️ Skipped duplicate file: {filename}")
            continue
        processed_files.add(filename)

        if "attendee" in filename.lower():
            result, cleaned, df_t = count_webinar_participant(file)
            df_email = clean_email_level(file)
        elif "participants" in filename.lower():
            result, cleaned, df_t = count_meeting_participant(file, excluded_name)
            df_email = pd.DataFrame()
        else:
            st.warning(f"⚠️ Skipped: {filename} (unknown type)")
            continue

        if not result.empty:
            data_summary = pd.concat([data_summary, result], ignore_index=True)
        if not df_email.empty:
            data_email = pd.concat([data_email, df_email], ignore_index=True)
        if not df_t.empty:
            data_country = pd.concat([data_country, df_t], ignore_index=True)

    if not data_summary.empty:
        # Merge aggregated country counts
        if not data_country.empty:
            data_country = data_country.groupby(["Date","Topic"]).sum().reset_index()
            data_summary = pd.merge(
                data_summary,
                data_country,
                on=["Date","Topic"],
                how="left"
            ).fillna(0)

        st.success("✅ Processing complete!")
        st.dataframe(data_summary[['Date','Topic','Total_Attendee','Total_Panelist','Total_All','Row_Deleted','Type']])
        st.text(f"Total Email {data_email.shape[0]}. Exclude Zoom Meeting (Region Not Available)")
        st.dataframe(data_email)
        
        # -----------------------------
        # Prepare CSVs
        # -----------------------------
        csv_summary = data_summary.to_csv(index=False).encode('utf-8')
        csv_email = data_email.to_csv(index=False).encode('utf-8')

        # Create ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            zip_file.writestr(f"{formatted_datetime}_zoom_summary.csv", csv_summary)
            zip_file.writestr(f"{formatted_datetime}_zoom_email_level.csv", csv_email)
        zip_buffer.seek(0)

        st.download_button(
            label="📥 Download Results (ZIP)",
            data=zip_buffer,
            file_name=f"{formatted_datetime}_zoom_reports.zip",
            mime="application/zip"
        )
else:
    st.info("Upload one or more Zoom CSV files to begin.")
