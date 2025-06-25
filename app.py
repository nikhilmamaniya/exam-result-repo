import streamlit as st
import pandas as pd
from io import BytesIO

def process_files(result_df, nad_df, subject_count=12, semester_number=2):
    merged_df = pd.merge(result_df, nad_df, left_on='PRN', right_on='REGN_NO')
    output_rows = []

    for i in range(1, subject_count + 1):
        prefix = f'Sub{i:02d}'
        if f'{prefix}_TOT' in merged_df.columns:
            for _, row in merged_df.iterrows():
                output_rows.append({
                    'StudentId': row.get('Stud ID', ''),
                    'ExamRollNo': row.get('RROLL', ''),
                    'StudentName': row.get('CNAME', ''),
                    'CourseName': row.get('COURSE_NAME', ''),
                    'Semester_No': semester_number,
                    'IsBackLog': 'NO',
                    'SubjectCode': row.get(f'{prefix}_CODE', ''),
                    'SubjectName': row.get(f'{prefix}_NAME', ''),
                    'Subject_Credit': row.get(f'{prefix}_CREDIT', ''),
                    'Internal_Marks': row.get(f'{prefix}_IA_MRKS', ''),
                    'External_Marks': row.get(f'{prefix}_UE_MRKS', ''),
                    'Total_Marks': row.get(f'{prefix}_TOT', ''),
                    'Grade': row.get(f'{prefix}_GRADE', ''),
                    'Grade_Point': row.get(f'{prefix}_GRADE_POINTS', ''),
                    'Credit_Earned': row.get(f'{prefix}_CREDIT', ''),
                    'Credit_Points': row.get(f'{prefix}_CREDIT_POINTS', ''),
                    'Result': row.get(f'{prefix}_Remark', ''),
                    'Semester_RegistredCredit': row.get('SEM_CREDIT_REGISTERED', ''),
                    'Semester_EarnCredit': row.get('SEM_CREDIT_EARNED', ''),
                    'Semester_EarnedGradePoint': row.get('SEM_EARNED_GRADE_POINTS', ''),
                    'SGPA': row.get('SGPA', ''),
                    'Semester_OverallGrade': row.get('SEM_GRADE', ''),
                    'CumullativeRegistredCredit': row.get('TOTAL_CREDIT_REGISTERED', ''),
                    'CumullativeEarnCredit': row.get('TOTAL_CREDIT_EARNED', ''),
                    'CumullativeEarnedGradePoint': row.get('TOTAL_EARNED_GRADE_POINTS', ''),
                    'CGPA': row.get('CGPA', ''),
                    'OverallGrade': row.get('GRADE', '')
                })

    return pd.DataFrame(output_rows)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- Streamlit UI ---
st.title("Student Exam Result Generator (Excel)")

st.markdown("""
Upload two Excel files:
1. 📄 **Result Format File** (e.g., `FRESH-RESULT-FORMAT.xlsx`)  
2. 📄 **NAD Data File** (e.g., `NAD-DATA.xlsx`)
""")

result_file = st.file_uploader("Upload Result Format File", type=["xlsx"])
nad_file = st.file_uploader("Upload NAD Data File", type=["xlsx"])

if result_file and nad_file:
    try:
        result_df = pd.read_excel(result_file)
        nad_df = pd.read_excel(nad_file)

        st.success("✅ Files uploaded and read successfully!")
        st.write("📋 Preview of Result File", result_df.head())
        st.write("📋 Preview of NAD File", nad_df.head())

        processed_df = process_files(result_df, nad_df)

        st.write("✅ Preview of Processed Data", processed_df.head())

        excel_data = to_excel(processed_df)
        st.download_button(
            label="📥 Download Output Excel File",
            data=excel_data,
            file_name="Generated_ExamMigrationData.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing files: {e}")
