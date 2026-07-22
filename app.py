import os
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="نظام مطابقة وتسوية الإكسل المرن", layout="wide")
st.title("🔍 نظام مطابقة وتسوية بيانات المدفوعات عبر الملفات (بحث ذكي)")

uploaded_files = st.file_uploader(
    "ارفع ملفات الإكسل للمطابقة", 
    type=["xlsx"], 
    accept_multiple_files=True
)

def smart_read_excel(file):
    """دالة قراءة ذكية تبحث عن صف العناوين المناسب بغض النظر عن الصفوف الأولى الفارغة أو الترتيب"""
    # قراءة أول 20 صف لمعاينة العناوين
    preview_df = pd.read_excel(file, nrows=20, header=None)
    
    header_row_index = 0
    target_keywords = ["رقم الحساب", "account no", "مبلغ المديونية", "payment amount"]
    
    # البحث عن الصف الذي يحتوي على أحد الكلمات المفتاحية للأعمدة
    for idx, row in preview_df.iterrows():
        row_str = row.astype(str).str.lower().str.strip().tolist()
        if any(any(kw in cell for kw in target_keywords) for cell in row_str):
            header_row_index = idx
            break
            
    # إعادة قراءة الملف بدءاً من الصف الصحيح الذي تم العثور فيه على العناوين
    df = pd.read_excel(file, header=header_row_index)
    return df

if uploaded_files and len(uploaded_files) > 1:
    match_cols = ["account no.", "payment amount", "payment date"]
    
    # قاموس لتطابق المسميات العربية والإنجليزية بجميع الأشكال المحتملة
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
            df = smart_read_excel(file)
            
            # تنظيف وتنسيق أسماء الأعمدة الموجودة
            df.columns = df.columns.astype(str).str.strip().str.lower()
            
            # إعادة تسمية الأعمدة بناءً على القاموس
            new_columns = {}
            for col in df.columns:
                for key, val in rename_dict.items():
                    if key in col: # مطابقة حتى لو كان هناك كلام إضافي بالعمود
                        new_columns[col] = val
                        break
            
            df = df.rename(columns=new_columns)

            # التحقق من وجود الأعمدة المطلوبة بغض النظر عن مكانها بين الأعمدة الأخرى
            if all(col in df.columns for col in match_cols):
                # تنظيف البيانات
                df["account no."] = df["account no."].astype(str).str.strip()
                df["source_file"] = file.name
                all_dfs.append(df)
            else:
                st.warning(f"⚠️ الملف ({file.name}) لم نجد فيه الأعمدة المطلوبة حتى بعد البحث الذكي.")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف {file.name}: {e}")

    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)
        total_files = len(all_dfs)

        # 1. المطابقة
        record_counts = (
            combined_df.groupby(match_cols)["source_file"]
            .nunique()
            .reset_index(name="file_count")
        )
        final_df = pd.merge(combined_df, record_counts, on=match_cols, how="left")

        # التقسيم
        existing_in_all = final_df[final_df["file_count"] == total_files].drop_duplicates(subset=match_cols)
        not_in_all = final_df[final_df["file_count"] < total_files].drop_duplicates(subset=match_cols)

        acc_file_count = combined_df.groupby("account no.")["source_file"].nunique()
        acc_in_multiple_files = acc_file_count[acc_file_count > 1].index
        
        different_data = combined_df[
            (combined_df["account no."].isin(acc_in_multiple_files)) & 
            (~combined_df["account no."].isin(existing_in_all["account no."]))
        ]

        st.success("✅ تم العثور على الأعمدة والمطابقة بنجاح!")

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
            label="📥 تحميل تقرير المطابقة النهائي (Excel)",
            data=output,
            file_name="تقرير_المطابقة_النهائي.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif uploaded_files and len(uploaded_files) == 1:
    st.info("💡 يرجى رفع ملفين إكسل أو أكثر لإجراء المقارنة والمطابقة.")
