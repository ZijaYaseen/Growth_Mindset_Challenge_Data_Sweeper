# Imports :

import streamlit as st
import pandas as pd
import os 
from io import BytesIO

# Set up ur App:
st.set_page_config(page_title="📀 Data Sweeper", layout="wide")
st.title("Data Sweeper!")
st.write("Transform your files between CSV and Excel formats with built-in data cleaning and visualization!")

uploaded_files = st.file_uploader("Upload your files (CSV or Excel):", type=["csv", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower() 

        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file)
        else:st.error(f"Unsupported file type: {file_ext}")
        continue

    # Display info about the file:
    st.write(f"**File Name:** {file.name}")
    st.write(f"**File Size:** {file.size / 1024:.2f} KB")

    # Showa % rows of our df
    st.write("Preview the Head of the Dataframe")
    st.dataframe(df.head())

    # Options for data cleaning
    st.subheader("Data Cleaning Options:")
    st.write("Check the box below to clean the data for this file. "
         "If you tick it, you'll be able to remove duplicate rows and fill missing values using the available options.")
    if st.checkbox(f"Clean Data For {file.name}"):
        col1, col2 = st.columns(2)

        with col1:
            if st.button(f"Remove duplicates from {file.name}"):
                df.drop_duplicates(inplace=True)
                st.write("Duplicates Removed!")

        with col2:
            if st.button(f"Fill Missig Values for {file.name}"):
                numeric_cols = df.select_dtypes(include=['number']).columns
                df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                st.write("Missing Values has been filled")

    # Choose specific Columns to keep or Convert
    st.subheader("Select Columns to Convert:") 
    columns = st.multiselect(f"Choose Colimns for  {file.name}", df.columns, default=df.columns)
    df = df[columns]

    # Create some Visualization
    st.subheader("📊 Data Visualization:")
    if st.checkbox(f"Show Visualization for {file.name}"):
        st.bar_chart(df.select_dtypes(include="number"))

    # Convert the file -> CSV to excel ;
    st.subheader("🔄 Conversion Options:")
    conversion_type = st.radio(f"Convert {file.name} to: ", ["CSV", "Excel"], key=file.name)
    if st.button(f"Convert {file.name}"):
        buffer = BytesIO()
        if conversion_type == "CSV":
            df.to_csv(buffer, index=False)
            file_name = file.name.replace(file_ext, ".csv")
            mime_type = "text/csv"

        elif conversion_type == "Excel":
            df.to_excel(buffer, index=False)
            file_name = file.name.replace(file_ext, ".xlsx")
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # it means k buffer start se read or download hoga    
        buffer.seek(0)  

    # Download Button
        st.download_button(
            label=f"Download {file.name} as {conversion_type}",
            data=buffer,
            file_name=file.name,
            mime=mime_type
    )

        st.success("🎉 All Files processeed!")
           