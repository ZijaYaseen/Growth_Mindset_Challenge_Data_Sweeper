import streamlit as st
import pandas as pd
import os
from io import BytesIO
from fpdf import FPDF

# MUST be the first Streamlit command
st.set_page_config(page_title="📀 Data Sweeper Advanced", layout="wide")

# --- Navigation via URL Query Params ---
params = st.query_params
if "page" in params:
    if isinstance(params["page"], list):
        current_page = params["page"][0].lower()
    else:
        current_page = params["page"].lower()
else:
    current_page = "converter"

# --- Sidebar Navigation (Link Style) ---
st.sidebar.markdown(
    """
    <style>
    .nav-header {
        text-align: center;
        font-size: 20px;
        color: #1f77b4;
        font-weight: bold;
        margin-bottom: 15px;
    }
    .nav-link {
        font-size: 18px;
        margin: 10px 0;
        text-decoration: none;
        color: #1f77b4;
        font-weight: bold;
    }
    .nav-link:hover {
        text-decoration: underline;
    }
    </style>
    """, unsafe_allow_html=True
)
st.sidebar.markdown("<div class='nav-header'>NAV BAR</div>", unsafe_allow_html=True)
st.sidebar.markdown("[Files Converter](?page=converter)", unsafe_allow_html=True)
st.sidebar.markdown("[Documentation](?page=documentation)", unsafe_allow_html=True)

# --- Main App Header ---
st.title("Data Sweeper Advanced")
st.write("""
Welcome to **Data Sweeper Advanced** – your one-stop tool for converting files between various formats.
You can convert files in the following ways:
- **CSV ⇄ Excel**
- **CSV/Excel → PDF**
- **Word → PDF**
- **PDF → Word**
- **PDF → CSV**

Use the sidebar links to switch between the conversion tool and detailed user documentation.
""")

# --- Display Content Based on Query Parameter ---
if current_page == "documentation":
    st.header("Data Sweeper Advanced - Detailed User Documentation")
    st.markdown("""
    ## Introduction
    **Data Sweeper Advanced** is a comprehensive file conversion tool designed to simplify the process of converting your documents between multiple formats. Whether you need to transform CSV files to Excel, convert PDFs to editable Word documents, or extract tables from PDFs into CSV, this tool provides a seamless solution.

    ## Purpose and Benefits
    - **Ease of Use:** The tool provides an intuitive interface that lets you perform conversions with just a few clicks.
    - **Versatility:** Supports multiple conversions such as CSV ⇄ Excel, CSV/Excel to PDF, Word to PDF, PDF to Word, and PDF to CSV.
    - **Time-Saving:** Automates the conversion process, eliminating the need for manual reformatting.
    - **Accessibility:** Built on Streamlit, it runs directly in your browser, making it accessible without complex installations.
    - **Open-Source:** Leverages popular Python libraries, ensuring reliability and community support.

    ## How to Use Data Sweeper Advanced
    Follow these simple steps to convert your files:

    ### 1. Navigation
    - **Sidebar Links:**  
      Use the **NAV BAR** on the left to switch between the **Converter** and **Documentation** pages.  
      - **Converter:** Perform your file conversions.
      - **Documentation:** Read this guide for detailed instructions.

    ### 2. Select a Conversion Option
    - On the **Converter** page, choose your desired conversion type from the dropdown menu. For example:
      - **CSV to Excel:** Converts CSV files into Excel spreadsheets.
      - **Excel to CSV:** Converts Excel files into CSV format.
      - **CSV/Excel to PDF:** Converts tabular data into a formatted PDF document.
      - **Word to PDF:** Converts Word documents (.docx) into PDFs.
      - **PDF to Word:** Converts PDF files into editable Word documents.
      - **PDF to CSV:** Extracts tables from PDF files and saves them as CSV.

    ### 3. Upload Your File
    - After selecting a conversion option, an upload box will appear.
    - **Note:** Ensure that you upload a file with the correct format for your chosen conversion (e.g., CSV for "CSV to Excel").

    ### 4. Preview and Convert
    - Once a file is uploaded, a preview (such as the first few rows of data) is displayed.
    - Click the conversion button to process the file. The tool will perform the conversion using reliable libraries and methods.

    ### 5. Download Your Converted File
    - After a successful conversion, a **Download** button will appear.
    - Click the button to save your newly converted file to your device.

    ## Why Was Data Sweeper Advanced Built?
    This tool was created to address common file conversion challenges:
    - **Simplifying Workflow:** Many users face difficulties when converting files between formats manually. Data Sweeper Advanced streamlines this process.
    - **Integration:** By combining multiple conversion functionalities in one application, it reduces the need for multiple tools.
    - **Reliability:** Built using well-established Python libraries, it offers dependable performance and accurate conversions.

    ## Additional Notes
    - **File Compatibility:** Always verify that your file is in the correct format before uploading.
    - **Error Handling:** If an error occurs during conversion, the app will display an error message with suggestions for troubleshooting.
    - **Future Updates:** The tool is designed to be extendable. Future updates may include additional conversion options and enhanced features.

    Enjoy hassle-free file conversions with Data Sweeper Advanced!
    """)
else:
    st.header("Data Sweeper Advanced - Converter")
    st.markdown("""
    **Select a Conversion Option Below:**

    * **CSV to Excel:** Convert CSV files into Excel format.
    * **Excel to CSV:** Convert Excel files into CSV format.
    * **CSV to PDF:** Generate a PDF from CSV data.
    * **Excel to PDF:** Generate a PDF from Excel data.
    * **Word to PDF:** Convert Word (.docx) files into PDF.
    * **PDF to Word:** Convert PDF files into editable Word documents.
    * **PDF to CSV:** Extract tables from PDFs and convert them into CSV.
    """)
    
    conversion_option = st.selectbox("Conversion Options", [
        "CSV to Excel",
        "Excel to CSV",
        "CSV to PDF",
        "Excel to PDF",
        "Word to PDF",
        "PDF to Word",
        "PDF to CSV"
    ])

    if conversion_option in ["CSV to Excel", "CSV to PDF"]:
        uploaded_file = st.file_uploader("Upload your CSV file:", type=["csv"])
    elif conversion_option in ["Excel to CSV", "Excel to PDF"]:
        uploaded_file = st.file_uploader("Upload your Excel file:", type=["xlsx"])
    elif conversion_option == "Word to PDF":
        uploaded_file = st.file_uploader("Upload your Word file (.docx):", type=["docx"])
    elif conversion_option in ["PDF to Word", "PDF to CSV"]:
        uploaded_file = st.file_uploader("Upload your PDF file:", type=["pdf"])
    else:
        uploaded_file = None

    if uploaded_file:
        st.write(f"**Uploaded File:** {uploaded_file.name} | Size: {uploaded_file.size / 1024:.2f} KB")
        
        if conversion_option == "CSV to Excel":
            with st.spinner("Converting CSV to Excel... Please wait."):
                df = pd.read_csv(uploaded_file)
                st.subheader("Preview of CSV Data:")
                st.dataframe(df.head())
                buffer = BytesIO()
                df.to_excel(buffer, index=False)
                buffer.seek(0)
                st.download_button("Download as Excel", data=buffer, file_name="converted.xlsx", 
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("CSV to Excel Conversion Successful!")
            
        elif conversion_option == "Excel to CSV":
            with st.spinner("Converting Excel to CSV... Please wait."):
                df = pd.read_excel(uploaded_file)
                st.subheader("Preview of Excel Data:")
                st.dataframe(df.head())
                buffer = BytesIO()
                df.to_csv(buffer, index=False)
                buffer.seek(0)
                st.download_button("Download as CSV", data=buffer, file_name="converted.csv", mime="text/csv")
            st.success("Excel to CSV Conversion Successful!")
            
        elif conversion_option == "CSV to PDF":
            with st.spinner("Converting CSV to PDF... Please wait."):
                df = pd.read_csv(uploaded_file)
                st.subheader("Preview of CSV Data:")
                st.dataframe(df.head())
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=10)
                col_width = pdf.w / (len(df.columns) + 1)
                for col in df.columns:
                    pdf.cell(col_width, 10, str(col), border=1)
                pdf.ln()
                for _, row in df.iterrows():
                    for item in row:
                        pdf.cell(col_width, 10, str(item), border=1)
                    pdf.ln()
                pdf_buffer = BytesIO()
                pdf.output(pdf_buffer)
                pdf_buffer.seek(0)
                st.download_button("Download as PDF", data=pdf_buffer, file_name="converted.pdf", mime="application/pdf")
            st.success("CSV to PDF Conversion Successful!")
            
        elif conversion_option == "Excel to PDF":
            with st.spinner("Converting Excel to PDF... Please wait."):
                df = pd.read_excel(uploaded_file)
                st.subheader("Preview of Excel Data:")
                st.dataframe(df.head())
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=10)
                col_width = pdf.w / (len(df.columns) + 1)
                for col in df.columns:
                    pdf.cell(col_width, 10, str(col), border=1)
                pdf.ln()
                for _, row in df.iterrows():
                    for item in row:
                        pdf.cell(col_width, 10, str(item), border=1)
                    pdf.ln()
                pdf_buffer = BytesIO()
                pdf.output(pdf_buffer)
                pdf_buffer.seek(0)
                st.download_button("Download as PDF", data=pdf_buffer, file_name="converted.pdf", mime="application/pdf")
            st.success("Excel to PDF Conversion Successful!")
            
        elif conversion_option == "Word to PDF":
            from docx2pdf import convert
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception as e:
                st.warning("COM initialization failed: " + str(e))
            with st.spinner("Converting Word to PDF... Please wait."):
                with open("temp.docx", "wb") as f:
                    f.write(uploaded_file.getbuffer())
                try:
                    convert("temp.docx", "converted.pdf")
                    with open("converted.pdf", "rb") as pdf_file:
                        pdf_data = pdf_file.read()
                    st.download_button("Download PDF", data=pdf_data, file_name="converted.pdf", mime="application/pdf")
                    st.success("Word to PDF Conversion Successful!")
                except Exception as e:
                    st.error(f"Error during conversion: {e}\nMake sure Microsoft Word is installed.")
                finally:
                    if os.path.exists("temp.docx"):
                        os.remove("temp.docx")
                    if os.path.exists("converted.pdf"):
                        os.remove("converted.pdf")
            
        elif conversion_option == "PDF to Word":
            from pdf2docx import Converter
            with st.spinner("Converting PDF to Word... Please wait."):
                with open("temp.pdf", "wb") as f:
                    f.write(uploaded_file.getbuffer())
                try:
                    cv = Converter("temp.pdf")
                    cv.convert("converted.docx", start=0, end=None)
                    cv.close()
                    with open("converted.docx", "rb") as docx_file:
                        docx_data = docx_file.read()
                    st.download_button("Download Word File", data=docx_data, file_name="converted.docx", 
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    st.success("PDF to Word Conversion Successful!")
                except Exception as e:
                    st.error(f"Error during conversion: {e}")
                finally:
                    if os.path.exists("temp.pdf"):
                        os.remove("temp.pdf")
                    if os.path.exists("converted.docx"):
                        os.remove("converted.docx")
            
        elif conversion_option == "PDF to CSV":
            with st.spinner("Converting PDF to CSV... Please wait."):
                try:
                    import tabula
                    with open("temp.pdf", "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    dfs = tabula.read_pdf("temp.pdf", pages='all', multiple_tables=True)
                    if dfs:
                        df = dfs[0]
                        st.subheader("Extracted Table Preview:")
                        st.dataframe(df.head())
                        buffer = BytesIO()
                        df.to_csv(buffer, index=False)
                        buffer.seek(0)
                        st.download_button("Download as CSV", data=buffer, file_name="converted.csv", mime="text/csv")
                        st.success("PDF to CSV Conversion Successful!")
                    else:
                        st.error("No tables found in the PDF.")
                except Exception as e:
                    st.error(f"Error during conversion: {e}")
                finally:
                    if os.path.exists("temp.pdf"):
                        os.remove("temp.pdf")
