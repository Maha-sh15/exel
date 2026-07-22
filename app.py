import os
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="استخراج السدادات الإضافية", layout="wide")
st.title("🎯 نظام استخراج السدادات الإضافية فقط")

st.markdown("""
الرجاء رفع الملفين المحددين لتحديد السدادات الإضافية (السدادات غير المسجلة في ملف النظام، أو المسجلة بتاريخ/مبلغ جديد):
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

    try:
        # قراءة الملفين
        df_system = smart_read_sadad_sheet(system_file)
        df_extra = smart_read_sadad_sheet(extra_file)

        # توحيد أسماء الأعمدة للملفين
        for df in [df_system, df_extra]:
            df.columns = df.columns.astype(str).str.strip().str.lower()
            new_columns = {}
            for col in df.columns:
                for key, val in rename_dict.items():
                    if key in col:
                        new_columns[col] = val
                        break
            df.rename(columns=new_columns, inplace=True)

        # التأكد من وجود الأعمدة المطلوبة
        if all(col in df_system.columns for col in match_cols) and all(col in df_extra.columns for col in match_cols):
            
            # تنظيف البيانات وتقليم المسافات
            for df in [df_system, df_extra]:
                df["account no."] = df["account no."].astype(str).str.strip()
                # توحيد تنسيق التواريخ لتجنب أخطاء الفوارق الزمنية
                df["payment date"] = pd.to_datetime(df["payment date"], errors='coerce').dt.strftime('%Y-%m-%d')

            # -------------------------------------------------------------
            #  منطق الفلترة: استخراج السدادات الإضافية فقط
            # -------------------------------------------------------------
            # البحث عن العمليات في ملف الإضافات التي لا تملك تطابقاً كاملاً (حساب + مبلغ + تاريخ) في ملف النظام
            merged = pd.merge(
                df_extra, 
                df_system[match_cols].drop_duplicates(), 
                on=match_cols, 
                how='left', 
                indicator=True
            )
            
            # العمليات التي تظهر في left_only هي السدادات الإضافية فقط
            extra_payments_only = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])

            st.success("✅ تم تحديد واستخراج السدادات الإضافية بنجاح!")

            # إحصائيات سريعة
            c1, c2 = st.columns(2)
            c1.metric("إجمالي عمليات ملف الإضافات", len(df_extra))
            c2.metric("السدادات الإضافية الجديدة المكتشفة", len(extra_payments_only))

            # تجهيز مسميات التصدير
            export_rename = {
                "account no.": "رقم الحساب",
                "payment amount": "مبلغ المديونية الحالي",
                "payment date": "تاريخ السداد"
            }

            # إنشاء ملف الإكسل
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # الشيت الرئيسي المخصص للسدادات الإضافية فقط
                extra_payments_only.rename(columns=export_rename).to_excel(
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
            st.error("⚠️ إحدى الأعمدة المطلوبة غير موجودة في أحد الملفين. تأكد من وجود (رقم الحساب، مبلغ المديونية، تاريخ السداد).")

    except Exception as e:
        st.error(f"حدث خطأ أثناء المعالجة: {e}")
