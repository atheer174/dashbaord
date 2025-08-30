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

try:
    import pdfplumber  # type: ignore
except ImportError:
    pdfplumber = None  # pdfplumber will be None if not installed


def extract_pdf_metrics(file: io.BytesIO) -> Tuple[Optional[str], Optional[str]]:
    """Extract contract documentation and Saudisation rates from a PDF.

    The function searches for Arabic phrases ``معدل التوطين`` and ``معدل توثيق العقود``
    followed by a percentage. It returns a tuple ``(saudisation_rate, contract_doc_rate)``.
    If a value cannot be found, the corresponding entry in the tuple will be ``None``.

    Parameters
    ----------
    file : io.BytesIO
        An in-memory file-like object containing PDF data.

    Returns
    -------
    Tuple[Optional[str], Optional[str]]
        A tuple of the extracted Saudisation rate and contract documentation rate.
    """
    if pdfplumber is None:
        return None, None
    try:
        saudisation_rate = None
        contract_doc_rate = None
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                # Normalize spaces
                text = re.sub(r"\s+", " ", text)
                # Search for pattern: معدل التوطين 75%
                match_saud = re.search(r"معدل التوطين\s*([0-9]{1,3}%?)", text)
                if match_saud and saudisation_rate is None:
                    saudisation_rate = match_saud.group(1)
                match_doc = re.search(r"معدل توثيق العقود\s*([0-9]{1,3}%?)", text)
                if match_doc and contract_doc_rate is None:
                    contract_doc_rate = match_doc.group(1)
        return saudisation_rate, contract_doc_rate
    except Exception:
        return None, None


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
            help="أدخل الرقم فقط إذا لم يتم استخلاصه من التقرير"
        )
        contract_doc_rate = st.text_input(
            "نسبة توثيق العقود (%)",
            value=contract_doc_rate if contract_doc_rate is not None else "",
            help="أدخل الرقم فقط إذا لم يتم استخلاصه من التقرير"
        )

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

            # Clean contract status values
            if 'Contract Status' not in contract_df.columns:
                st.error("العمود 'Contract Status' غير موجود في ورقة العقد.")
                return
            contract_df['Contract Status'] = contract_df['Contract Status'].astype(str).str.strip()

            # Contract type counts
            contract_counts = contract_df['Contract Status'].value_counts()

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
