import os
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="نظام مطابقة وتسوية الإكسل", layout="wide")
st.title("🔍 نظام مطابقة وتسوية بيانات المدفوعات عبر الملفات")

# رفع ملفات متعددة
uploaded_files = st.file_uploader(
    "ارفع ملفات الإكسل للمطابقة", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files and len(uploaded_files) > 1:
    match_cols = ["account no", "payment amount", "payment date"]
    all_dfs = []

    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        # التأكد من وجود الأعمدة
        if all(col in df.columns for col in match_cols):
            # تنظيف البيانات
            df["account no"] = df["account no"].astype(str).str.strip()
            df["source_file"] = file.name
            all_dfs.append(df)
        else:
            st.warning(f"الملف {file.name} يفتقد لبعض الأعمدة المطلوبة وتم تجاهله.")

    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)
        total_files = len(all_dfs)

        # 1. مطابقة السجلات بالكامل
        record_counts = (
            combined_df.groupby(match_cols)["source_file"]
            .nunique()
            .reset_index(name="file_count")
        )
        final_df = pd.merge(combined_df, record_counts, on=match_cols, how="left")

        # الفئة 1: موجودين في كل الملفات
        existing_in_all = final_df[final_df["file_count"] == total_files].drop_duplicates(subset=match_cols)

        # الفئة 2: غير موجودين في كل الملفات (موجود في بعضها فقط)
        not_in_all = final_df[final_df["file_count"] < total_files].drop_duplicates(subset=match_cols)

        # الفئة 3: حسابات مشتركة لكن مبالغها/تواريخها مختلفة بين الملفات
        acc_file_count = combined_df.groupby("account no")["source_file"].nunique()
        acc_in_multiple_files = acc_file_count[acc_file_count > 1].index
        
        # نأخذ الحسابات المكررة في أكثر من ملف ولكن لا تملك تطابق تام في الفئة الأولى
        different_data = combined_df[
            (combined_df["account no"].isin(acc_in_multiple_files)) & 
            (~combined_df["account no"].isin(existing_in_all["account no"]))
        ]

        st.success("✅ تمت عملية المطابقة بنجاح!")

        # عرض النتائج
        col1, col2, col3 = st.columns(3)
        col1.metric("موجودين في كل الملفات", len(existing_in_all))
        col2.metric("موجودين في بعض الملفات", len(not_in_all))
        col3.metric("حسابات بها اختلافات", len(different_data))

        # تنزيل النتائج في ملف إكسل واحد بشيتات متعددة
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            existing_in_all.drop(columns=["source_file", "file_count"], errors="ignore").to_excel(writer, sheet_name="موجود بكل الملفات", index=False)
            not_in_all.drop(columns=["source_file", "file_count"], errors="ignore").to_excel(writer, sheet_name="موجود ببعض الملفات", index=False)
            different_data.drop(columns=["source_file"], errors="ignore").to_excel(writer, sheet_name="بيانات مختلفة", index=False)
        output.seek(0)

        st.download_button(
            label="📥 تحميل تقرير المطابقة الشامل (Excel)",
            data=output,
            file_name="تقرير_المطابقة_النهائي.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
elif uploaded_files and len(uploaded_files) == 1:
    st.info("💡 يرجى رفع ملفين إكسل أو أكثر لإجراء المقارنة والمطابقة.")
