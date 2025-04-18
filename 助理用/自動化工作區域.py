import pandas as pd
import json
import os
import subprocess
from tkinter import Tk, filedialog

root = Tk()
root.withdraw()

excel_path = filedialog.askopenfilename(
    title="請選擇 Excel 檔案",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not excel_path:
    print("❌ 你沒選 Excel，掰掰。")
    exit()

import tkinter as tk
from tkinter import simpledialog

class StartCodeDialog(simpledialog.Dialog):
    def body(self, master):
        self.title("請輸入起始代號")

        label = tk.Label(
            master,
            text="Illustrator 檔案的第一個設計代號（例如：001）",
            font=("Arial", 14),
            fg="white"   # ⬅ 黑字
        )
        label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.entry = tk.Entry(
            master,
            font=("Arial", 18),
            bg="#2C2C2C",    # ⬅ 白底
            fg="white",    # ⬅ 黑字
            justify='center',
            insertbackground="black"  # ⬅ 讓游標也是黑的（不然可能看不到）
        )
        self.entry.grid(row=1, column=0, padx=20, pady=(0, 20))

        return self.entry  # 初始 focus 在這

    def apply(self):
        self.result = self.entry.get().strip()

# 呼叫美化後的輸入視窗
start_code_dialog = StartCodeDialog(root)
start_code = start_code_dialog.result

if not start_code:
    print("❌ 沒輸入起始代號，掰掰。")
    exit()

jsx_template_path = '/Users/onlycolor/Desktop/自動化 工作區域.jsx'
generated_jsx_path = '/Users/onlycolor/Desktop/run_auto_layout.jsx'

df = pd.read_excel(excel_path, header=None)

start_index = None
for i, row in df.iterrows():
    if str(row[1]).strip() == start_code:
        start_index = i
        break

if start_index is None:
    print(f"❌ Excel 中找不到設計代號「{start_code}」")
    exit()

excel_rows = df.iloc[start_index:]
design_data = []

for _, row in excel_rows.iterrows():
    try:
        name = str(row[1]).strip().zfill(3)
        width = float(row[18])
        height = float(row[19])
        design_data.append({
            "name": name,
            "width": width,
            "height": height
        })
    except Exception as e:
        print(f"⚠️ 資料錯誤：{e}")
        continue

design_map = {
    d["name"]: {"width": d["width"], "height": d["height"]}
    for d in design_data
}

design_json = 'var designMap = ' + json.dumps(design_map, indent=2) + ';\n'

with open(jsx_template_path, 'r', encoding='utf-8') as f:
    jsx_code = f.read()

if 'var designMap = /*__DESIGN_MAP__*/;' not in jsx_code:
    print("❌ 模板裡沒這一行：var designMap = /*__DESIGN_MAP__*/;")
    exit()

jsx_code = jsx_code.replace('var designMap = /*__DESIGN_MAP__*/;', design_json)

with open(generated_jsx_path, 'w', encoding='utf-8') as f:
    f.write(jsx_code)

print(f"✅ 已產生腳本：{generated_jsx_path}")

subprocess.call([
    'osascript', '-e',
    f'''
    tell application "Adobe Illustrator"
        activate
        do javascript POSIX file "{os.path.abspath(generated_jsx_path)}"
    end tell
    '''
])