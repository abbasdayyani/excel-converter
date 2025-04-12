from flask import Flask, request, send_file, render_template_string
import pandas as pd
import tempfile

app = Flask(__name__)

HTML_PAGE = '''
<!doctype html>
<title>Excel Converter</title>
<h1>آپلود فایل اکسل و دریافت فایل خروجی</h1>
<form action="/convert" method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value="آپلود و دریافت خروجی">
</form>
'''

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file:
        return "فایلی انتخاب نشده است."

    df = pd.read_excel(file)

    # تبدیل ستون‌های عددی از متن به عدد
    for col in ['فی', 'تعداد', 'مالیات', 'تخفیف']:
        df[col] = df[col].astype(str).str.replace(',', '').astype(float)

    # ساختن دیتا فریم جدید
    new_df = pd.DataFrame({
	'نوع قلم': df['کالا'],
        'فاكتور شماره': [''] * len(df),
        'فاكتور تاريخ': df['تاریخ'],
        'فاكتور كد مشتري': [''] * len(df),
        'فاكتور كد نوع فروش': [''] * len(df),
        'اقلام فاكتور كد': [''] * len(df),
        'اقلام فاكتور كد انبار': [''] * len(df),
        'اقلام فاكتور عنوان رديابي': [''] * len(df),
        'اقلام فاكتور واحد اصلي': df['تعداد'],
        'اقلام فاكتور واحد فرعي': [''] * len(df),
        'اقلام فاكتور في': df['فی'],
        'اقلام فاكتور كل': df['تعداد'],
        'اقلام فاكتور ماليات': df['مالیات'],
        'اقلام فاكتور عوارض': [0] * len(df),
        'اقلام فاكتور تخفيف مبلغي اعلاميه قيمت': [0] * len(df),
        'اقلام فاكتور تخفيف': df['تخفیف'],
        'اقلام فاكتور اضافات': [0] * len(df),
        'اقلام فاكتور توضيحات': df['کاربر'],
        'فاكتور نام مشتري': [''] * len(df),
        'فاكتور محل تحويل1': [''] * len(df),
        'اقلام فاكتور تخفيف مشتري': [0] * len(df),
        'اقلام فاكتور تخفيف درصدي اعلاميه قيمت': [0] * len(df),
        'فاكتور ارز1': ['ریال'] * len(df),
        'فاكتور نرخ ارز': [1] * len(df),
        'فاكتور نوع تسويه': [3] * len(df),
    })

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    new_df.to_excel(temp_file.name, index=False)

    return send_file(temp_file.name, as_attachment=True, download_name="output.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
