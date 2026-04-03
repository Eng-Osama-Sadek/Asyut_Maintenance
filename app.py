from flask import Flask, render_template, request, jsonify
import pandas as pd
from datetime import datetime
import os
import webbrowser # المكتبة المسؤولة عن فتح المتصفح تلقائياً

app = Flask(__name__)

# --- قائمة الهندسات الـ 12 لقطاع أسيوط جنوب ---
DEPTS = [
    "هندسة كهرباء مركز جنوب", "هندسة كهرباء مركز شمال", "هندسة كهرباء شرق",
    "هندسة كهرباء الخزان", "هندسة كهرباء غرب", "هندسة كهرباء مبارك",
    "هندسة كهرباء أبوتيج قرى", "هندسة كهرباء أبوتيج المدينة", "هندسة كهرباء الغنايم",
    "هندسة كهرباء صدفا", "هندسة كهرباء الساحل", "هندسة كهرباء البداري"
]

# --- قائمة بنود الأعمال الشاملة (33 بند) ---
ITEMS = [
    {"id": "b1", "name": "صيانة محول معلق", "unit": "عدد"},
    {"id": "b2", "name": "صيانة محول حجرة", "unit": "عدد"},
    {"id": "b3", "name": "صيانة محول كشك", "unit": "عدد"},
    {"id": "b4", "name": "صيانة خط هوائي جهد متوسط", "unit": "كم"},
    {"id": "b5", "name": "هوائي جهد منخفض", "unit": "كم"},
    {"id": "b6", "name": "صندوق توزيع جهد منخفض", "unit": "عدد"},
    {"id": "b7", "name": "جهاز ريكلوزر", "unit": "عدد"},
    {"id": "b8", "name": "وحدة ربط حلقي (RMU)", "unit": "عدد"},
    {"id": "b9", "name": "موزع جهد متوسط", "unit": "عدد"},
    {"id": "b10", "name": "منظم جهد متوسط", "unit": "عدد"},
    {"id": "b11", "name": "زرع عامود جهد متوسط", "unit": "عدد"},
    {"id": "b12", "name": "زرع عامود جهد منخفض", "unit": "عدد"},
    {"id": "b13", "name": "خلع وتركيب عامود جهد متوسط", "unit": "عدد"},
    {"id": "b14", "name": "خلع وتركيب عامود جهد منخفض", "unit": "عدد"},
    {"id": "b15", "name": "دهان عامود جهد متوسط", "unit": "عدد"},
    {"id": "b16", "name": "دهان عامود جهد منخفض", "unit": "عدد"},
    {"id": "b17", "name": "شد وتحريب موصلات جهد متوسط", "unit": "كم"},
    {"id": "b18", "name": "شد وتحريب موصلات جهد منخفض", "unit": "كم"},
    {"id": "b19", "name": "تركيب عازل قرص", "unit": "عدد"},
    {"id": "b20", "name": "تركيب عازل مسمار", "unit": "عدد"},
    {"id": "b21", "name": "تركيب عازل صيني 14سم منخفض", "unit": "عدد"},
    {"id": "b22", "name": "عمل فورمة خرسانية لعامود متوسط", "unit": "عدد"},
    {"id": "b23", "name": "عمل فورمة خرسانية لعامود منخفض", "unit": "عدد"},
    {"id": "b24", "name": "صيانة عامود جهد متوسط", "unit": "عدد"},
    {"id": "b25", "name": "صيانة عامود جهد منخفض", "unit": "عدد"},
    {"id": "b26", "name": "تركيب كابل جهد متوسط", "unit": "كم"},
    {"id": "b27", "name": "تركيب كابل جهد منخفض", "unit": "كم"},
    {"id": "b28", "name": "تركيب قاطع ثلاثي جهد منخفض", "unit": "عدد"},
    {"id": "b29", "name": "تركيب سكينة هوائية متوسط", "unit": "عدد"},
    {"id": "b30", "name": "تركيب سكينة داخلية متوسط", "unit": "عدد"},
    {"id": "b31", "name": "دهان كشك محول", "unit": "عدد"},
    {"id": "b32", "name": "عمل طوب فرعوني لقاعدة كشك", "unit": "م²"},
    {"id": "b33", "name": "صب قاعدة خرسانية لكشك محول", "unit": "عدد"}
]

@app.route('/')
def index():
    return render_template('login.html', depts=DEPTS, items=ITEMS)

@app.route('/api/save', methods=['POST'])
def save_data():
    try:
        data = request.json
        file_path = "Maintenance_Database_Final.xlsx"
        
        records = []
        for item_id, val in data['entries'].items():
            item_info = next(i for i in ITEMS if i['id'] == item_id)
            target = float(val['target']) if val['target'] else 0
            done = float(val['done']) if val['done'] else 0
            percent = (done / target * 100) if target > 0 else 0
            
            records.append({
                "التوقيت": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "التاريخ": data['date'],
                "الهندسة": data['dept'],
                "المسؤول": data['user_name'],
                "بند العمل": item_info['name'],
                "الوحدة": item_info['unit'],
                "المستهدف": target,
                "المنفذ": done,
                "نسبة التنفيذ %": f"{percent:.1f}%",
                "اسم المحول": data['trans_name'],
                "قدرة المحول": data['trans_power'],
                "عنوان المحول": data['trans_addr'],
                "اسم الخط": data['line_name'],
                "مصدر الخط": data['line_src'],
                "طول الخط": data['line_len'],
                "ملاحظات": data['notes']
            })
        
        df_new = pd.DataFrame(records)
        if os.path.exists(file_path):
            df_old = pd.read_excel(file_path)
            df_final = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df_final = df_new
            
        df_final.to_excel(file_path, index=False)
        return jsonify({"msg": "تم حفظ التقرير بنجاح في ملف الإكسيل!"})
    except Exception as e:
        return jsonify({"msg": f"حدث خطأ أثناء الحفظ: {str(e)}"}), 500

if __name__ == '__main__':
    # فتح المتصفح أوتوماتيكياً على رابط البرنامج
    webbrowser.open("http://127.0.0.1:5000")
    # تشغيل السيرفر
    app.run(debug=False)