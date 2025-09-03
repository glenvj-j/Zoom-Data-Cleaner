import streamlit as st
import pandas as pd

def clean_data(file):
    # Reset file pointer
    file.seek(0)
    lines = file.readlines()
    file.seek(0)

    # Find header row (the one that contains "Name (original name)")
    header_row = None
    for i, line in enumerate(lines):
        if b"User Name (Original Name)" in line:
            header_row = i
            break

    if header_row is None:
        st.error(f"❌ Could not detect Attendee header row in {file.name}")
        return pd.DataFrame()

    # Read CSV starting from the detected header row
    df = pd.read_csv(file, skiprows=header_row)

    # Get Topic name
    file.seek(0)
    topic_df = pd.read_csv(file, skiprows=2, nrows=1)
    Topic = topic_df['Topic'].iloc[0].replace('iBlooming: ', "")

    # Get the Attendee Data
    df_attendee = df.iloc[int(df[df['Attended']=="Attendee Details"].index[0])+2:]
    
    # Take the latest time to get the date
    Date = df_attendee['Join Time'].iloc[-1][0:10]

    # Drop the duplicated email
    duplicated_mask = df_attendee.drop(
        columns=[
            'User Name (Original Name)','Attended','Join Time','Leave Time',
            'Time in Session (minutes)','Is Guest','Country/Region Name'
        ],
        errors="ignore"
    ).duplicated()

    df_attendee_clean = df_attendee[~duplicated_mask]
    
    df_temp = df_attendee_clean[['User Name (Original Name)','Email','Country/Region Name']]
    df_temp['topic'] = Topic
    df_temp['date'] = Date
    return df_temp

# -----------------------------
# STREAMLIT APP
# -----------------------------
uploaded_files = st.sidebar.file_uploader(
    "Upload Zoom Attendee CSV files", 
    type=["csv"], 
    accept_multiple_files=True
)

if uploaded_files:
    df_all = pd.DataFrame()
    cleaned_files = {}

    for file in uploaded_files:
        filename = file.name
        if "attendee" in filename.lower():
            result = clean_data(file)
            df_all = pd.concat([df_all, result], ignore_index=True)
            cleaned_files[filename] = result
        else:
            st.warning(f"⚠️ Skipped: {filename} (not recognized as attendee file)")
            continue

    if not df_all.empty:
        st.sidebar.success("✅ Processing complete!")
        st.text(f'Total Data : {df_all.shape[0]}')
        st.dataframe(df_all)
        

        # Download combined summary
        csv_all = df_all.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Combined Cleaned CSV",
            data=csv_all,
            file_name="zoom_attendees_all.csv",
            mime="text/csv"
        )
else:
    st.info("Upload one or more Zoom attendee CSV files to begin.")