import streamlit as st
import pandas as pd
import os
from io import BytesIO
from fpdf import FPDF

############################
# Uniform Row Table Helpers
############################

def draw_uniform_row(pdf, row_data, col_widths, line_height=5):
    """
    1) Pre-split each cell's text into wrapped lines (split_only=True).
    2) Find the max number of lines (max_lines) among all columns for this row.
    3) For each column, draw a single rectangle that spans row_height = max_lines * line_height.
    4) Print each line of text inside that rectangle, so columns remain aligned.
    """

    # Pre-split each cell's text
    splitted_lines = []
    for i, cell_text in enumerate(row_data):
        lines = pdf.multi_cell(
            col_widths[i],
            line_height,
            str(cell_text),
            border=0,
            align='L',
            split_only=True
        )
        splitted_lines.append(lines)

    # Determine max lines for this row
    max_lines = max(len(lines) for lines in splitted_lines)
    row_height = max_lines * line_height

    # Store the starting X/Y
    x_start = pdf.get_x()
    y_start = pdf.get_y()

    # Draw rectangles for each column
    # so the entire row has a top/bottom border, and vertical borders for columns.
    current_x = x_start
    for i in range(len(col_widths)):
        pdf.rect(current_x, y_start, col_widths[i], row_height)
        current_x += col_widths[i]

    # Now, print the text lines inside each rectangle
    for line_index in range(max_lines):
        # Start each line at the left margin of this row
        pdf.set_xy(x_start, y_start + line_index * line_height)
        for i, col_lines in enumerate(splitted_lines):
            text_line = col_lines[line_index] if line_index < len(col_lines) else ""
            # Print a single line in this column
            pdf.cell(col_widths[i], line_height, text_line, border=0, ln=0)
        pdf.ln(line_height)

    # Move the cursor down to the next row
    pdf.set_xy(x_start, y_start + row_height)


def draw_uniform_table(pdf, df, line_height=5):
    """
    Creates a new page (landscape, A4), sets small font,
    calculates column widths, and draws the table from a DataFrame with uniform row heights.
    """
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", size=7)

    col_count = len(df.columns)
    page_width = pdf.w - 20  # account for left/right margins
    col_width = page_width / col_count
    col_widths = [col_width] * col_count

    # 1) Draw header row
    draw_uniform_row(pdf, list(df.columns), col_widths, line_height=line_height)

    # 2) Draw each data row
    for _, row_data in df.iterrows():
        draw_uniform_row(pdf, list(row_data), col_widths, line_height=line_height)


############################
# Streamlit App Starts Here
############################

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
    """,
    unsafe_allow_html=True
)
st.sidebar.markdown("<div class='nav-header'>NAV BAR</div>", unsafe_allow_html=True)
st.sidebar.markdown("[Files Converter](?page=converter)", unsafe_allow_html=True)
st.sidebar.markdown("[Documentation](?page=documentation)", unsafe_allow_html=True)

# --- Main App Header ---
st.title("Data Sweeper Advanced")
st.write("""
Welcome to **Data Sweeper Advanced** – your one-stop tool for converting files between various formats:
- **CSV ⇄ Excel**
- **CSV/Excel → PDF** (uniform row heights, multi-line wrapping, no extra lines!)
- **Word → PDF**
- **PDF → Word**
- **PDF → CSV**

Use the sidebar links to switch between the conversion tool and detailed user documentation.
""")

if current_page == "documentation":
    st.header("Data Sweeper Advanced - Detailed User Documentation")
    st.markdown("""
    ## Introduction
    **Data Sweeper Advanced** is a comprehensive file conversion tool designed to simplify
    the process of converting your documents between multiple formats...
    (Your documentation content here)
    """)
else:
    st.header("Data Sweeper Advanced - Converter")
    st.markdown("""
    **Select a Conversion Option Below:**
    
    * **CSV to Excel**
    * **Excel to CSV**
    * **CSV to PDF** (uniform row heights, multi-line wrap)
    * **Excel to PDF** (uniform row heights, multi-line wrap)
    * **Word to PDF**
    * **PDF to Word**
    * **PDF to CSV**
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

    # File Uploader
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

        # CSV to Excel
        if conversion_option == "CSV to Excel":
            with st.spinner("Converting CSV to Excel... Please wait."):
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, engine="python", sep=None, on_bad_lines='skip', encoding="latin1")
                except Exception as e:
                    st.error(f"Error reading CSV file: {e}")
                else:
                    if df.empty:
                        st.error("CSV file is empty or invalid.")
                    else:
                        st.subheader("Preview of CSV Data:")
                        st.dataframe(df.head())
                        buffer = BytesIO()
                        df.to_excel(buffer, index=False)
                        buffer.seek(0)
                        st.download_button("Download as Excel",
                                           data=buffer,
                                           file_name="converted.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("CSV to Excel Conversion Successful!")

        # Excel to CSV
        elif conversion_option == "Excel to CSV":
            with st.spinner("Converting Excel to CSV... Please wait."):
                df = pd.read_excel(uploaded_file)
                st.subheader("Preview of Excel Data:")
                st.dataframe(df.head())
                buffer = BytesIO()
                df.to_csv(buffer, index=False)
                buffer.seek(0)
                st.download_button("Download as CSV",
                                   data=buffer,
                                   file_name="converted.csv",
                                   mime="text/csv")
            st.success("Excel to CSV Conversion Successful!")

        # CSV to PDF (Uniform Row)
        elif conversion_option == "CSV to PDF":
            with st.spinner("Converting CSV to PDF... Please wait."):
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, engine="python", sep=None, on_bad_lines='skip', encoding="latin1")
                except Exception as e:
                    st.error(f"Error reading CSV file: {e}")
                else:
                    if df.empty:
                        st.error("CSV file is empty or invalid.")
                    else:
                        st.subheader("Preview of CSV Data:")
                        st.dataframe(df.head())
                        pdf = FPDF(orientation='L', unit='mm', format='A4')
                        draw_uniform_table(pdf, df, line_height=5)
                        pdf_bytes = pdf.output(dest="S").encode("latin1")
                        pdf_buffer = BytesIO(pdf_bytes)
                        st.download_button("Download as PDF",
                                           data=pdf_buffer,
                                           file_name="converted.pdf",
                                           mime="application/pdf")
            st.success("CSV to PDF Conversion Successful!")

        # Excel to PDF (Uniform Row)
        elif conversion_option == "Excel to PDF":
            with st.spinner("Converting Excel to PDF... Please wait."):
                df = pd.read_excel(uploaded_file)
                if df.empty:
                    st.error("The Excel file appears to be empty or invalid.")
                else:
                    st.subheader("Preview of Excel Data:")
                    st.dataframe(df.head())
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    draw_uniform_table(pdf, df, line_height=5)
                    pdf_bytes = pdf.output(dest="S").encode("latin1")
                    pdf_buffer = BytesIO(pdf_bytes)
                    st.download_button("Download as PDF",
                                       data=pdf_buffer,
                                       file_name="converted.pdf",
                                       mime="application/pdf")
            st.success("Excel to PDF Conversion Successful!")

        # Word to PDF
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
                    st.download_button("Download PDF",
                                       data=pdf_data,
                                       file_name="converted.pdf",
                                       mime="application/pdf")
                    st.success("Word to PDF Conversion Successful!")
                except Exception as e:
                    st.error(f"Error during conversion: {e}\nMake sure Microsoft Word is installed.")
                finally:
                    if os.path.exists("temp.docx"):
                        os.remove("temp.docx")
                    if os.path.exists("converted.pdf"):
                        os.remove("converted.pdf")

        # PDF to Word
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
                    st.download_button("Download Word File",
                                       data=docx_data,
                                       file_name="converted.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    st.success("PDF to Word Conversion Successful!")
                except Exception as e:
                    st.error(f"Error during conversion: {e}")
                finally:
                    if os.path.exists("temp.pdf"):
                        os.remove("temp.pdf")
                    if os.path.exists("converted.docx"):
                        os.remove("converted.docx")

        # PDF to CSV
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
                        st.download_button("Download as CSV",
                                           data=buffer,
                                           file_name="converted.csv",
                                           mime="text/csv")
                        st.success("PDF to CSV Conversion Successful!")
                    else:
                        st.error("No tables found in the PDF.")
                except Exception as e:
                    st.error(f"Error during conversion: {e}")
                finally:
                    if os.path.exists("temp.pdf"):
                        os.remove("temp.pdf")
