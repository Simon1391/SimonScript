from flask import Flask, render_template
from flask_socketio import SocketIO
import json, os
from datetime import datetime

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key_here'
# 允許所有來源連線
socketio = SocketIO(app, cors_allowed_origins="*")

SHARED_FOLDER = "/Volumes/助理美工區/#國軒"

def load_all_stats():
    agg = {"employees": {}}
    if os.path.exists(SHARED_FOLDER):
        for fn in os.listdir(SHARED_FOLDER):
            if fn.endswith(".json"):
                path = os.path.join(SHARED_FOLDER, fn)
                try:
                    with open(path, encoding="utf8") as f:
                        data = json.load(f)
                    emp = os.path.splitext(fn)[0]
                    agg["employees"][emp] = data
                except Exception as e:
                    print(f"讀取 {fn} 錯誤：{e}")
    else:
        print("共享資料夾不存在！")
    return agg

def convert_stats_for_template(emp_stats):
    today = datetime.today().strftime("%Y-%m-%d")
    month = datetime.today().strftime("%Y-%m")
    daily = emp_stats.get("daily", {}).get(today, {})
    monthly = emp_stats.get("monthly", {}).get(month, {})
    # 如果沒有平面 file_count 就合併 regular + overtime
    if "file_count" not in daily:
        r = daily.get("regular", {}); o = daily.get("overtime", {})
        daily["file_count"] = r.get("file_count",0) + o.get("file_count",0)
        daily["material"]   = r.get("material",0.0)   + o.get("material",0.0)
    if "file_count" not in monthly:
        r = monthly.get("regular", {}); o = monthly.get("overtime", {})
        monthly["file_count"] = r.get("file_count",0) + o.get("file_count",0)
        monthly["material"]   = r.get("material",0.0)   + o.get("material",0.0)
    return {
        "employee": emp_stats.get("display_name",""),
        "today_regular_count":   daily.get("regular",{}).get("file_count",0),
        "today_regular_material":daily.get("regular",{}).get("material",0.0),
        "today_overtime_count":  daily.get("overtime",{}).get("file_count",0),
        "today_overtime_material":daily.get("overtime",{}).get("material",0.0),
        "monthly": monthly
    }

@app.route("/")
def dashboard():
    agg = load_all_stats()
    stats = []
    for emp, data in agg["employees"].items():
        c = convert_stats_for_template(data)
        c["employee"] = emp
        stats.append(c)
    return render_template("dashboard.html", employee_stats=stats)

# 當任何 client emit 'stats_update'，重新讀檔並推給所有連線
@socketio.on('stats_update')
def on_stats_update(msg):
    print("收到 stats_update：", msg)
    agg = load_all_stats()
    stats = []
    for emp, data in agg["employees"].items():
        c = convert_stats_for_template(data)
        c["employee"] = emp
        stats.append(c)
    socketio.emit('stats_update', {'employee_stats': stats})

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=5001, debug=True, allow_unsafe_werkzeug=True)