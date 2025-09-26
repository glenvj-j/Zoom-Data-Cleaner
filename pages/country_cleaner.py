import streamlit as st
import pandas as pd
import io
import zipfile

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

    # -----------------------------
    # Panelists section
    # -----------------------------
    attendee_idx = df.index[df['Attended'] == "Attendee Details"].tolist()
    df_panelist = pd.DataFrame()
    if attendee_idx:  
        df_panelist = df.iloc[2:attendee_idx[0]].copy()  # rows above "Attendee Details"
        if not df_panelist.empty and "Email" in df_panelist.columns:
            df_panelist = df_panelist.drop_duplicates(subset=["Email"])
            keep_cols = ["User Name (Original Name)", "Email"]
            if "Country/Region Name" in df_panelist.columns:
                keep_cols.append("Country/Region Name")
            df_panelist = df_panelist[keep_cols].copy()
            df_panelist["Role"] = "Panelist"
        else:
            df_panelist = pd.DataFrame()

    # -----------------------------
    # Attendee section
    # -----------------------------
    df_attendee = df.iloc[int(attendee_idx[0]) + 2 :] if attendee_idx else pd.DataFrame()

    if not df_attendee.empty:
        Date = df_attendee['Join Time'].iloc[-1][0:10]

        duplicated_mask = df_attendee.drop(
            columns=[
                'User Name (Original Name)','Attended','Join Time','Leave Time',
                'Time in Session (minutes)','Is Guest','Country/Region Name'
            ],
            errors="ignore"
        ).duplicated()

        df_attendee_clean = df_attendee[~duplicated_mask]

        keep_cols = ["User Name (Original Name)", "Email"]
        if "Country/Region Name" in df_attendee_clean.columns:
            keep_cols.append("Country/Region Name")

        df_attendee_clean = df_attendee_clean[keep_cols].copy()
        df_attendee_clean["Role"] = "Attendee"
    else:
        df_attendee_clean = pd.DataFrame()
        Date = None

    # -----------------------------
    # Merge Panelists + Attendees
    # -----------------------------
    df_combined = pd.concat([df_panelist, df_attendee_clean], ignore_index=True)

    if not df_combined.empty:
        df_combined["Topic"] = Topic
        df_combined["Date"] = Date

    return df_combined



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
            if not result.empty:
                df_all = pd.concat([df_all, result], ignore_index=True)
                cleaned_files[filename] = result
        else:
            st.warning(f"⚠️ Skipped: {filename} (not recognized as attendee file)")

    if not df_all.empty:
        st.sidebar.success("✅ Processing complete!")
        st.text(f'Total Data : {df_all.shape[0]}')
        st.dataframe(df_all)

        # Pivot for horizontal view
        df_t = df_all.pivot_table(
            index=["Date","Topic"],
            columns="Country/Region Name",
            values="Email",
            aggfunc="count",
            fill_value=0
        )
        df_t.columns.name = None   # remove "Country/Region Name" label
        df_t = df_t.reset_index()

        st.text(f'Horizontal Data')
        st.dataframe(df_t)

        # -----------------------------
        # Prepare both CSVs
        # -----------------------------
        csv_all = df_all.to_csv(index=False).encode('utf-8')
        csv_horizontal = df_t.to_csv(index=False).encode('utf-8')
        
        # Create an in-memory ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            zip_file.writestr("zoom_attendees_all.csv", csv_all)
            zip_file.writestr("horizontal.csv", csv_horizontal)
        
        # Move cursor to the beginning of the buffer
        zip_buffer.seek(0)
        
        # Single download button
        st.download_button(
            label="📥 Download All CSVs (ZIP)",
            data=zip_buffer,
            file_name="zoom_reports.zip",
            mime="application/zip"
        )
else:
    st.info("Upload one or more Zoom attendee CSV files to begin.")
