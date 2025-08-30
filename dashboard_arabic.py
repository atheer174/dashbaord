"""
Arabic Dashboard for Workforce Insights
--------------------------------------

This Streamlit application allows users to upload two Excel spreadsheets and
an optional PDF report to generate key workforce metrics for their
organisation. It supports Arabic labels throughout the interface and
calculates the following indicators based on the uploaded data:

* عدد العقود محددة المدة وعدد العقود غير محددة المدة
* عدد الموظفين الأجانب وعدد الموظفين الأجانب الذين لديهم أكثر من أربعة تابعين
* معدل توثيق العقود ومعدل التوطين – يتم استخلاصهما تلقائياً من تقرير PDF إن وجد

The app expects the following structure in the Excel files:

1. **Employee file (قوى.xlsx)** – must contain at least the columns:
   - ``Id number`` – a unique identifier for each employee (may be the national ID or iqama number).
   - ``Contract Status`` – values such as ``محدد`` (fixed term) or ``غير محدد`` (indefinite).
   - ``Nationality`` – the employee’s nationality written in English (e.g., ``Saudi Arabia``).

2. **Dependents file (المقيمين النشطين و تابعيهم.xlsx)** – must contain at least:
   - ``رقم إقامة رب الأسرة`` – the iqama number of the head of household (employee).
   - one row per dependent.

3. **Monthly report PDF** – optional. If provided, the app uses ``pdfplumber`` to extract the
   contract documentation rate (معدل توثيق العقود) and the Saudisation rate (معدل التوطين).
   The PDF must contain these Arabic phrases followed by the percentage values. If the text
   extraction fails, the app will ask the user to enter the values manually.

To run the app in Google Colab you can execute the following in a cell:

```python
!pip install streamlit pdfplumber
!streamlit run dashboard_arabic.py --server.port 8501 --server.address 0.0.0.0
```

Streamlit will provide a public URL when run inside Colab. Open that URL to view the dashboard.

"""

import io
import re
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

# Optional imports for PDF parsing. We try to use docling first and fall back to
# pdfplumber if docling is unavailable or fails. Both imports are wrapped in
# try/except blocks so the app can still run if the libraries are not
# installed. See the documentation for docling for installation instructions【725798722026482†L344-L352】.
try:
    # Prefer docling for robust PDF parsing
    from docling.document_converter import DocumentConverter  # type: ignore
    _docling_available = True
except Exception:
    DocumentConverter = None  # type: ignore
    _docling_available = False

try:
    import pdfplumber  # type: ignore
except Exception:
    pdfplumber = None  # type: ignore


def extract_pdf_metrics(file: io.BytesIO) -> Tuple[Optional[str], Optional[str]]:
    """Extract Saudisation and contract documentation rates from a PDF.

    This function attempts to locate the phrases ``معدل التوطين`` and
    ``معدل توثيق العقود`` in the text of the PDF and then extract the first
    percentage value that appears nearby. It prefers using the Docling
    ``DocumentConverter`` for robust PDF parsing when available and falls back
    to pdfplumber otherwise. After converting the document to plain text,
    whitespace is normalised and simple regex patterns are used to capture
    Arabic or Western numerals followed by a percent sign. Eastern Arabic
    numerals are converted to Western numerals before returning.

    Parameters
    ----------
    file : io.BytesIO
        A PDF file-like object.

    Returns
    -------
    tuple
        ``(saudisation_rate, contract_doc_rate)`` where each value is a string
        including the percent sign (e.g., ``"75%"``) or ``None`` if not found.
    """
    # Helper to parse text and extract the rates
    def parse_rates(text: str) -> Tuple[Optional[str], Optional[str]]:
        saud_rate: Optional[str] = None
        doc_rate: Optional[str] = None
        if not text:
            return saud_rate, doc_rate
        # Normalize whitespace to single spaces
        text = re.sub(r"\s+", " ", text)
        # Search for Saudisation phrase and extract first percentage afterwards
        idx_saud = text.find("معدل التوطين")
        if idx_saud != -1:
            window = text[idx_saud:idx_saud + 80]  # look ahead 80 characters
            m = re.search(r"([0-9٠-٩]+)\s*%", window)
            if m:
                num = m.group(1)
                arabic_to_western = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
                num = num.translate(arabic_to_western)
                saud_rate = f"{num}%"
        # Search for contract documentation phrase
        idx_doc = text.find("معدل توثيق العقود")
        if idx_doc != -1:
            window = text[idx_doc:idx_doc + 80]
            m = re.search(r"([0-9٠-٩]+)\s*%", window)
            if m:
                num = m.group(1)
                arabic_to_western = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
                num = num.translate(arabic_to_western)
                doc_rate = f"{num}%"
        return saud_rate, doc_rate

    extracted_text: Optional[str] = None

    # Try using docling first if available
    if _docling_available and DocumentConverter is not None:
        try:
            # Save the PDF to a temporary file path because Docling expects a path or URL
            import tempfile, os
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_fp:
                tmp_fp.write(file.getvalue())
                tmp_path = tmp_fp.name
            converter = DocumentConverter()
            result = converter.convert(tmp_path)
            # Some conversions may return None if the pipeline fails
            if result and result.document:
                # Try export_to_text (preferred) then fallback to export_to_markdown
                try:
                    extracted_text = result.document.export_to_text()
                except Exception:
                    try:
                        extracted_text = result.document.export_to_markdown()
                    except Exception:
                        extracted_text = None
        except Exception:
            extracted_text = None
        finally:
            # Clean up temporary file
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

    # If docling was not successful, try pdfplumber
    if (not extracted_text) and pdfplumber is not None:
        try:
            text_parts: list[str] = []
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
            if text_parts:
                extracted_text = "\n".join(text_parts)
        except Exception:
            extracted_text = None

    if not extracted_text:
        return None, None
    return parse_rates(extracted_text)


def main() -> None:
    """Run the Streamlit dashboard application."""
    st.set_page_config(page_title="لوحة مؤشرات الموارد البشرية", layout="wide")
    st.title("لوحة مؤشرات الموارد البشرية")
    st.markdown("""
    **مرحبًا!** يمكنكم هنا رفع ملفات البيانات لتحليل معلومات الموظفين بشكل تلقائي.
    الرجاء رفع ملف الموظفين، ملف التابعين، ويمكن رفع تقرير PDF لاستخراج نسب التوطين وتوثيق العقود.
    """)

    # Uploaders
    emp_file = st.file_uploader("❶ – تحميل ملف الموظفين (Excel)", type=["xlsx"], key="emp")
    dep_file = st.file_uploader("❷ – تحميل ملف التابعين (Excel)", type=["xlsx"], key="dep")
    pdf_file = st.file_uploader("❸ – تحميل التقرير الشهري (PDF) – اختياري", type=["pdf"], key="pdf")

    # Initialize placeholders for metrics
    saudisation_rate: Optional[str] = None
    contract_doc_rate: Optional[str] = None

    if pdf_file is not None:
        saudisation_rate, contract_doc_rate = extract_pdf_metrics(pdf_file)

    # If PDF extraction failed, allow manual input
    with st.expander("تعديل نسب التوطين وتوثيق العقود يدويًا", expanded=(saudisation_rate is None or contract_doc_rate is None)):
        saudisation_rate = st.text_input(
            "نسبة التوطين (%)",
            value=saudisation_rate if saudisation_rate is not None else "",
            help="أدخل الرقم أو النسبة المئوية إذا لم يتم استخلاصها من التقرير"
        )
        contract_doc_rate = st.text_input(
            "نسبة توثيق العقود (%)",
            value=contract_doc_rate if contract_doc_rate is not None else "",
            help="أدخل الرقم أو النسبة المئوية إذا لم يتم استخلاصها من التقرير"
        )

    # Always display compliance metrics if any value is provided, even without Excel files
    if saudisation_rate or contract_doc_rate:
        st.subheader("مؤشرات الامتثال")
        comp_cols = st.columns(2)
        if saudisation_rate:
            comp_cols[0].metric("نسبة التوطين", saudisation_rate)
        else:
            comp_cols[0].write("نسبة التوطين غير مدخلة")
        if contract_doc_rate:
            comp_cols[1].metric("نسبة توثيق العقود", contract_doc_rate)
        else:
            comp_cols[1].write("نسبة توثيق العقود غير مدخلة")

    if emp_file and dep_file:
        try:
            # Read dependents data from the first sheet
            dep_xls = pd.ExcelFile(dep_file)
            dep_df = pd.read_excel(dep_xls, sheet_name=dep_xls.sheet_names[0])

            # Read employee file and automatically detect relevant sheets
            emp_xls = pd.ExcelFile(emp_file)
            contract_sheet = None
            master_sheet = None
            # Identify sheet containing 'Contract Status' and 'Nationality'
            for sheet in emp_xls.sheet_names:
                try:
                    tmp_df = pd.read_excel(emp_xls, sheet_name=sheet, nrows=1)
                except Exception:
                    continue
                cols = tmp_df.columns
                if 'Contract Status' in cols and contract_sheet is None:
                    contract_sheet = sheet
                if 'Nationality' in cols and 'Id number' in cols and master_sheet is None:
                    master_sheet = sheet
            if contract_sheet is None:
                st.error("لم يتم العثور على ورقة تحتوي على عمود 'Contract Status'. يرجى التحقق من ملف الموظفين.")
                return
            if master_sheet is None:
                st.error("لم يتم العثور على ورقة تحتوي على أعمدة 'Nationality' و 'Id number'. يرجى التحقق من ملف الموظفين.")
                return

            # Load sheets fully
            contract_df = pd.read_excel(emp_xls, sheet_name=contract_sheet)
            master_df = pd.read_excel(emp_xls, sheet_name=master_sheet)

            # Ensure the contract status column exists
            if 'Contract Status' not in contract_df.columns:
                st.error("العمود 'Contract Status' غير موجود في ورقة العقد.")
                return
            # Normalize contract status values: remove surrounding spaces and drop missing
            contract_df['Contract Status'] = contract_df['Contract Status'].astype(str).str.strip()
            contract_counts = contract_df['Contract Status'][
                ~contract_df['Contract Status'].isin(['', 'nan', 'NaN'])
            ].value_counts()

            # Prepare for merging
            # Ensure 'Id number' exists in both
            if 'Id number' not in master_df.columns or 'Id number' not in contract_df.columns:
                st.error("العمود 'Id number' غير موجود في إحدى أوراق الموظفين.")
                return
            master_df['Id number'] = master_df['Id number'].astype(str)
            contract_df['Id number'] = contract_df['Id number'].astype(str)

            # Merge nationality and contract status
            emp_df = pd.merge(master_df[['Id number', 'Nationality']], contract_df[['Id number', 'Contract Status']], on='Id number', how='inner')

            # Count foreign employees
            foreign_emps = emp_df[emp_df['Nationality'].str.lower() != 'saudi arabia']
            num_foreign = len(foreign_emps)

            # Process dependents: convert to numeric for merge
            dep_df['رقم إقامة رب الأسرة'] = pd.to_numeric(dep_df['رقم إقامة رب الأسرة'], errors='coerce').astype('Int64')
            emp_df['Id number_int'] = pd.to_numeric(emp_df['Id number'], errors='coerce').astype('Int64')

            # Group dependents count by head id
            dep_counts = dep_df.groupby('رقم إقامة رب الأسرة').size().reset_index(name='dependents')
            emp_merged = pd.merge(emp_df, dep_counts, left_on='Id number_int', right_on='رقم إقامة رب الأسرة', how='left')
            emp_merged['dependents'] = emp_merged['dependents'].fillna(0)

            foreign_with_many_dep = emp_merged[(emp_merged['Nationality'].str.lower() != 'saudi arabia') & (emp_merged['dependents'] > 4)]

            # Display metrics
            st.subheader("النتائج الرئيسية")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("العقود محددة المدة", int(contract_counts.get('محدد', 0)))
            col2.metric("العقود غير محددة المدة", int(contract_counts.get('غير محدد', 0)))
            col3.metric("عدد الموظفين الأجانب", num_foreign)
            col4.metric("عدد الأجانب مع أكثر من أربعة تابعين", len(foreign_with_many_dep))

            # Display compliance metrics
            with st.container():
                st.subheader("مؤشرات الامتثال")
                metric_cols = st.columns(2)
                if saudisation_rate:
                    metric_cols[0].metric("نسبة التوطين", f"{saudisation_rate}")
                else:
                    metric_cols[0].write("لم يتم توفير نسبة التوطين")
                if contract_doc_rate:
                    metric_cols[1].metric("نسبة توثيق العقود", f"{contract_doc_rate}")
                else:
                    metric_cols[1].write("لم يتم توفير نسبة توثيق العقود")

            # Additional statistics
            with st.expander("إحصائيات إضافية"):
                st.subheader("توزيع حالة العقد")
                st.bar_chart(contract_counts)
                st.subheader("توزيع الجنسيات")
                nat_counts = emp_df['Nationality'].value_counts()
                st.bar_chart(nat_counts)

        except Exception as exc:
            st.error(f"حدث خطأ أثناء معالجة البيانات: {exc}")


if __name__ == "__main__":
    main()