import os
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="نظام مطابقة وتصفية صفحة سداد", layout="wide")
st.title("🔍 نظام مطابقة المدفوعات (التركيز على صفحة سداد فقط)")

uploaded_files = st.file_uploader(
    "ارفع ملفات الإكسل للمطابقة", 
    type=["xlsx"], 
    accept_multiple_files=True
)

def smart_read_sadad_sheet(file):
    """دالة تقرأ صفحة 'سداد' فقط وتبحث فيها عن صف العناوين المناسب"""
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names
    
    # البحث عن الورقة التي تحتوي على كلمة "سداد"
    target_sheet = None
    for sheet in sheet_names:
        if "سداد" in str(sheet).strip():
            target_sheet = sheet
            break
            
    # إذا لم يجد صفحة باسم "سداد"، يأخذ الورقة الأولى كخيار احتياطي
    if target_sheet is None:
        target_sheet = sheet_names[0]
        st.warning(f"⚠️ الملف ({file.name}) لا يحتوي على صفحة باسم 'سداد'، تم قراءة الصفحة الأولى ({target_sheet}) بدلاً عنها.")

    # قراءة أول 20 صف لمعاينة العناوين داخل صفحة سداد
    preview_df = pd.read_excel(file, sheet_name=target_sheet, nrows=20, header=None)
    
    header_row_index = 0
    target_keywords = ["رقم الحساب", "account no", "مبلغ المديونية", "payment amount"]
    
    # البحث عن صف العناوين الحقيقي
    for idx, row in preview_df.iterrows():
        row_str = row.astype(str).str.lower().str.strip().tolist()
        if any(any(kw in cell for kw in target_keywords) for cell in row_str):
            header_row_index = idx
            break
            
    # إعادة قراءة صفحة "سداد" فقط من الصف الصحيح
    df = pd.read_excel(file, sheet_name=target_sheet, header=header_row_index)
    return df

if uploaded_files and len(uploaded_files) > 1:
    match_cols = ["account no.", "payment amount", "payment date"]
    
    rename_dict = {
        "رقم الحساب": "account no.",
        "account no": "account no.",
        "account no.": "account no.",
        "مبلغ المديونية الحالي": "payment amount",
        "مبلغ المديونية": "payment amount",
        "payment amount": "payment amount",
        "تاريخ السداد": "payment date",
        "تاريخ سداد": "payment date",
        "payment date": "payment date"
    }

    all_dfs = []

    for file in uploaded_files:
        try:
            # قراءة صفحة سداد حصراً
            df = smart_read_sadad_sheet(file)
            
            # تنظيف وتنسيق أسماء الأعمدة
            df.columns = df.columns.astype(str).str.strip().str.lower()
            
            # إعادة تسمية الأعمدة المطابقة
            new_columns = {}
            for col in df.columns:
                for key, val in rename_dict.items():
                    if key in col:
                        new_columns[col] = val
                        break
            
            df = df.rename(columns=new_columns)

            # التحقق من الأعمدة
            if all(col in df.columns for col in match_cols):
                df["account no."] = df["account no."].astype(str).str.strip()
                df["source_file"] = file.name
                all_dfs.append(df)
            else:
                st.warning(f"⚠️ الملف ({file.name}) - صفحة 'سداد' لم نجد فيها الأعمدة المطلوبة.")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف {file.name}: {e}")

    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)
        total_files = len(all_dfs)

        # 1. عملية المطابقة
        record_counts = (
            combined_df.groupby(match_cols)["source_file"]
            .nunique()
            .reset_index(name="file_count")
        )
        final_df = pd.merge(combined_df, record_counts, on=match_cols, how="left")

        existing_in_all = final_df[final_df["file_count"] == total_files].drop_duplicates(subset=match_cols)
        not_in_all = final_df[final_df["file_count"] < total_files].drop_duplicates(subset=match_cols)

        acc_file_count = combined_df.groupby("account no.")["source_file"].nunique()
        acc_in_multiple_files = acc_file_count[acc_file_count > 1].index
        
        different_data = combined_df[
            (combined_df["account no."].isin(acc_in_multiple_files)) & 
            (~combined_df["account no."].isin(existing_in_all["account no."]))
        ]

        st.success("✅ تم توجيه المطابقة لصفحة 'سداد' فقط وتمت المعالجة بنجاح!")

        col1, col2, col3 = st.columns(3)
        col1.metric("موجودين في كل الملفات", len(existing_in_all))
        col2.metric("موجودين في بعض الملفات", len(not_in_all))
        col3.metric("حسابات بها اختلافات", len(different_data))

        export_rename = {
            "account no.": "Account No.",
            "payment amount": "Payment Amount",
            "payment date": "Payment Date"
        }

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            existing_in_all.drop(columns=["source_file", "file_count"], errors="ignore").rename(columns=export_rename).to_excel(writer, sheet_name="موجود بكل الملفات", index=False)
            not_in_all.drop(columns=["source_file", "file_count"], errors="ignore").rename(columns=export_rename).to_excel(writer, sheet_name="موجود ببعض الملفات", index=False)
            different_data.drop(columns=["source_file"], errors="ignore").rename(columns=export_rename).to_excel(writer, sheet_name="بيانات مختلفة", index=False)
        output.seek(0)

        st.download_button(
            label="📥 تحميل تقرير مطابقة صفحات سداد (Excel)",
            data=output,
            file_name="تقرير_مطابقة_سداد.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif uploaded_files and len(uploaded_files) == 1:
    st.info("💡 يرجى رفع ملفين إكسل أو أكثر لإجراء المقارنة والمطابقة.")
