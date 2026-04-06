from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# مسار حفظ البيانات (متوافق مع السيرفرات السحابية)
EXCEL_FILE = 'maintenance_data.xlsx'

def save_to_excel(data):
    if not os.path.isfile(EXCEL_FILE):
        df = pd.DataFrame([data])
        df.to_excel(EXCEL_FILE, index=False)
    else:
        df = pd.read_excel(EXCEL_FILE)
        new_row = pd.DataFrame([data])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    form_data = request.form.to_dict()
    form_data['تاريخ_التسجيل'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_to_excel(form_data)
    return "تم تسجيل البيانات بنجاح! شكراً يا هندسة."

@app.route('/download')
def download():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)
    return "لا توجد بيانات مسجلة حتى الآن."

if __name__ == '__main__':
    # تشغيل متوافق مع البيئات المحلية والسحابية
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)