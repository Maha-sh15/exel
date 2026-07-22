import os
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="استخراج السدادات الإضافية", layout="wide")
st.title("🎯 نظام استخراج السدادات الإضافية (مع الحفاظ على أسماء الأعمدة الأصلية)")

st.markdown("""
الرجاء رفع الملفين لتحديد السدادات الإضافية (السدادات غير المسجلة في ملف النظام، أو المسجلة بتاريخ/مبلغ جديد):
""")

col1, col2 = st.columns(2)

with col1:
    system_file = st.file_uploader("1️⃣ ارفع ملف سدادات النظام (الملف الأساسي)", type=["xlsx"])

with col2:
    extra_file = st.file_uploader("2️⃣ ارفع الملف الذي يحتوي على سدادات إضافية", type=["xlsx"])

def smart_read_sadad_sheet(file):
    """دالة تقرأ صفحة 'سداد' فقط وتتجاوز أي صفوف تعريفية في البداية"""
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names
    
    target_sheet = None
    for sheet in sheet_names:
        if "سداد" in str(sheet).strip():
            target_sheet = sheet
            break
            
    if target_sheet is None:
        target_sheet = sheet_names[0]

    preview_df = pd.read_excel(file, sheet_name=target_sheet, nrows=20, header=None)
    
    header_row_index = 0
    target_keywords = ["رقم الحساب", "account no", "مبلغ المديونية", "payment amount"]
    
    for idx, row in preview_df.iterrows():
        row_str = row.astype(str).str.lower().str.strip().tolist()
        if any(any(kw in cell for kw in target_keywords) for cell in row_str):
            header_row_index = idx
            break
            
    df = pd.read_excel(file, sheet_name=target_sheet, header=header_row_index)
    return df

if system_file and extra_file:
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

    try:
        # قراءة الملفين كما هما
        df_system_raw = smart_read_sadad_sheet(system_file)
        df_extra_raw = smart_read_sadad_sheet(extra_file)

        # عمل نسخة مؤقتة توحد أسماء الأعمدة للقيام بذكاء المطابقة فقط
        df_system_proc = df_system_raw.copy()
        df_extra_proc = df_extra_raw.copy()

        for df in [df_system_proc, df_extra_proc]:
            df.columns = df.columns.astype(str).str.strip().str.lower()
            new_cols = {}
            for col in df.columns:
                for key, val in rename_dict.items():
                    if key in col:
                        new_cols[col] = val
                        break
            df.rename(columns=new_cols, inplace=True)

        match_cols = ["account no.", "payment amount", "payment date"]

        # التأكد من نجاح العثور على الأعمدة في النسخ المؤقتة
        if all(col in df_system_proc.columns for col in match_cols) and all(col in df_extra_proc.columns for col in match_cols):
            
            # تنظيف البيانات المؤقتة للمطابقة
            df_system_proc["account no."] = df_system_proc["account no."].astype(str).str.strip()
            df_extra_proc["account no."] = df_extra_proc["account no."].astype(str).str.strip()

            df_system_proc["payment date"] = pd.to_datetime(df_system_proc["payment date"], errors='coerce').dt.strftime('%Y-%m-%d')
            df_extra_proc["payment date"] = pd.to_datetime(df_extra_proc["payment date"], errors='coerce').dt.strftime('%Y-%m-%d')

            # الربط والمطابقة بالخلفية
            merged = pd.merge(
                df_extra_proc, 
                df_system_proc[match_cols].drop_duplicates(), 
                on=match_cols, 
                how='left', 
                indicator=True
            )
            
            # تحديد مؤشرات الصفوف الإضافية فقط
            extra_indexes = merged[merged['_merge'] == 'left_only'].index

            # استخراج الصفوف الإضافية من الملف الأصلي للحفاظ على الأعمدة ومسمياتها دون تغيير
            extra_payments_original = df_extra_raw.iloc[extra_indexes].copy()

            st.success("✅ تم استخراج السدادات الإضافية بنجاح مع الحفاظ على مسميات الأعمدة الأصلية!")

            c1, c2 = st.columns(2)
            c1.metric("إجمالي عمليات ملف الإضافات", len(df_extra_raw))
            c2.metric("السدادات الإضافية الجديدة المكتشفة", len(extra_payments_original))

            # تصدير الملف بشيت "سدادات اضافية"
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                extra_payments_original.to_excel(
                    writer, 
                    sheet_name="سدادات اضافية", 
                    index=False
                )

            output.seek(0)

            st.download_button(
                label="📥 تحميل شيت (سدادات اضافية)",
                data=output,
                file_name="تقرير_السدادات_الإضافية.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        else:
            st.error("⚠️ لم نتمكن من التعرف على أعمدة المطابقة الرئيسية (الحساب، المبلغ، التاريخ). تأكد من وجودها في ورقة سداد.")

    except Exception as e:
        st.error(f"حدث خطأ أثناء المعالجة: {e}")
