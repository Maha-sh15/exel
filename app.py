import os
import pandas as pd


def process_excel_files(file_paths, output_folder="output_results"):
    """نظام تصنيف العملاء بناءً على التكرار وتطابق البيانات عبر ملفات إكسل متعددة."""
    # إنشاء مجلد للمخرجات إذا لم يكن موجوداً
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # الأعمدة الأساسية للمطابقة والمقارنة
    match_cols = ["account no", "payment amount", "payment date"]

    all_dfs = []

    # 1. قراءة جميع الملفات وتحديد مصدر كل سجل
    for path in file_paths:
        if not os.path.exists(path):
            print(f"الملف غير موجود: {path}")
            continue

        df = pd.read_excel(path)

        # توحيد أسماء الأعمدة (إزالة الفراغات وتحويلها لأحرف صغيرة لتفادي الأخطاء)
        df.columns = df.columns.str.strip().str.lower()

        # التأكد من وجود الأعمدة المطلوبة
        missing_cols = [col for col in match_cols if col not in df.columns]
        if missing_cols:
            print(
                f"تنبيه: الملف {path} لا يحتوي على الأعمدة التالية: {missing_cols}"
            )
            continue

        # إضافة عمود لمعرفة اسم الملف الأصلي لكل سجل
        df["source_file"] = os.path.basename(path)
        all_dfs.append(df)

    if not all_dfs:
        print("لم يتم العثور على ملفات صالحة للمعالجة.")
        return

    # دمج كل البيانات في جدول واحد كبير
    combined_df = pd.concat(all_dfs, ignore_index=True)

    # إجمالي عدد الملفات الفريدة الصالحة
    total_files_count = len(file_paths)

    # 2. حساب تكرار السجل المطابق تماماً عبر الملفات
    # هنا نقوم بجمع السجلات المتطابقة تماماً في (الحساب، المبلغ، التاريخ) ونرى في كم ملف ظهرت
    record_counts = (
        combined_df.groupby(match_cols)["source_file"]
        .nunique()
        .reset_index(name="file_count")
    )

    # دمج عدد الملفات مع الجدول المشترك
    final_df = pd.merge(combined_df, record_counts, on=match_cols, how="left")

    # -------------------------------------------------------------------------
    # تصنيف الفئات الثلاث المطلوبة:
    # -------------------------------------------------------------------------

    # الفئة 1: العملاء الموجودين في كل الملفات بنفس البيانات (تكرار كامل)
    # تظهر هذه السجلات في عدد ملفات يساوي إجمالي عدد الملفات المرفوعة
    existing_in_all = final_df[final_df["file_count"] == total_files_count]

    # الفئة 2: العملاء غير الموجودين في كامل الملفات (موجود في بعضها وغائب عن بعضها)
    # تظهر في ملف واحد أو أكثر ولكن أقل من العدد الإجمالي للملفات
    not_in_all = final_df[final_df["file_count"] < total_files_count]

    # الفئة 3: العملاء الموجودين بالملفات ولكن بياناتهم (المبلغ أو التاريخ) مختلفة لنفس الحساب
    # لمعرفة ذلك، نتحقق من الحسابات (account no) التي تظهر ببيانات مختلفة
    account_variance = (
        combined_df.groupby("account no")[["payment amount", "payment date"]]
        .nunique()
        .max(axis=1)
    )
    different_data_accounts = account_variance[account_variance > 1].index

    different_data = combined_df[
        combined_df["account no"].isin(different_data_accounts)
    ]

    # -------------------------------------------------------------------------
    # حفظ الملفات الناتجة
    # -------------------------------------------------------------------------

    # تنظيف الأعمدة الإضافية قبل الحفظ للحفاظ على شكل الملف الأصلي
    cols_to_drop = ["source_file", "file_count"]

    existing_in_all.drop(
        columns=cols_to_drop, errors="ignore"
    ).drop_duplicates(subset=match_cols).to_excel(
        os.path.join(output_folder, "1_العملاء_الموجودين_في_كل_الملفات.xlsx"),
        index=False,
    )

    not_in_all.drop(columns=cols_to_drop, errors="ignore").drop_duplicates(
        subset=match_cols
    ).to_excel(
        os.path.join(output_folder, "2_العملاء_الغير_موجودين_في_كامل_الملفات.xlsx"),
        index=False,
    )

    different_data.drop(columns=["source_file"], errors="ignore").to_excel(
        os.path.join(output_folder, "3_عملاء_بياناتهم_مختلفة.xlsx"), index=False
    )

    print(f"تمت العملية بنجاح! تم حفظ الملفات الثلاثة في المجلد: '{output_folder}'")


# --- مثال على طريقة التشغيل ---
if __name__ == "__main__":
    # ضع هنا مسارات ملفات الإكسل الخاصة بك (يمكنك إضافة أي عدد من الملفات)
    my_files = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]

    # تشغيل النظام
    process_excel_files(my_files)
