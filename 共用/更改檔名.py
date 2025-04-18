import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES  # 用於拖放檔案功能
from tkinter import messagebox
from tkinter import ttk
import shutil

import os
import shutil
from tkinter import messagebox

def apply_renaming(rename_info):
    """
    1. 先把所有 (old, new) 轉成絕對路徑並蒐集
    2. 用 os.path.exists 檢查一輪：只要有一個目標檔已存在，就跳錯誤並 return
    3. 無衝突時才真正執行搬移
    """
    abs_pairs = []   # list of (old_abs, dst_abs)
    conflicts = []   # list of dst_abs that 已存在

    for old, new in rename_info:
        old_abs = os.path.abspath(old)
        # 保險起見，如果 dirname 回空，就用 cwd
        dirname = os.path.dirname(old_abs) or os.getcwd()
        dst_abs = os.path.join(dirname, new)
        abs_pairs.append((old_abs, dst_abs))

        if os.path.exists(dst_abs):
            conflicts.append(dst_abs)

    if conflicts:
        # 一旦有任何衝突，就一次顯示所有衝突路徑，並中止
        msg = "以下目標檔案已存在，已中止整批改名：\n\n" + "\n".join(conflicts)
        messagebox.showerror("改名衝突", msg)
        return

    # 無衝突才開始搬
    for old_abs, dst_abs in abs_pairs:
        try:
            shutil.move(old_abs, dst_abs)
            print(f"檔案 {old_abs} 已更名為 {dst_abs}")
        except Exception as e:
            # 萬一跨磁碟、權限或其他失敗，也會印出錯誤
            print(f"改名失敗：{old_abs} → {dst_abs}，原因：{e}")


def show_preview_window_treeview(data):
    """
    data 為一個列表，每一項為 (原始檔案路徑, 改名後檔案名稱) 的 tuple。
    此函式建立一個 Toplevel 視窗，利用 Treeview 以兩個欄位呈現預覽資料，
    左側欄位顯示「舊檔案」（僅顯示檔案名稱），右側欄位顯示「新檔案」。
    視窗下方有【確定】與【取消】按鈕，使用者選擇後返回結果。
    """
    result = {"confirmed": False}
    top = tk.Toplevel(root)
    top.title("確認更名")
    top.transient(root)
    top.resizable(False, False)
    top.configure(bg="#ffffff")

    # 取得主視窗的位置，讓 Toplevel 與主視窗跳出同一位置
    x = root.winfo_rootx()
    y = root.winfo_rooty()
    top.geometry(f"700x400+{x}+{y}")

    # 設定全域字型 (14pt, Arial, Bold)
    default_font = ("Arial", 14, "bold")
    header_font = ("Arial", 14, "bold")

    # 設定 Treeview 樣式：移除邊框、採用白底黑字，並套用指定字型
    style = ttk.Style(top)
    style.theme_use("clam")
    style.configure("Custom.Treeview",
                    background="#ffffff",
                    foreground="#000000",
                    fieldbackground="#ffffff",
                    borderwidth=0,
                    font=default_font,
                    rowheight=28)
    style.layout("Custom.Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
    style.configure("Custom.Treeview.Heading",
                    background="#ffffff",
                    foreground="#000000",
                    borderwidth=0,
                    font=header_font)
    style.map("Custom.Treeview",
              background=[("selected", "#f0f0f0")],
              foreground=[("selected", "#000000")])

    # 用 grid 版面配置：Treeview 區（第0行）與按鈕區（第1行）
    top.grid_rowconfigure(0, weight=1)
    top.grid_rowconfigure(1, weight=0)
    top.grid_columnconfigure(0, weight=1)

    # 建立放置 Treeview 的框架 (無邊框)
    frame = tk.Frame(top, bg="#ffffff", bd=0, highlightthickness=0)
    frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

    # 建立 Treeview，指定自訂的 style
    columns = ("old", "new")
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=15, style="Custom.Treeview")
    tree.heading("old", text="舊檔案")
    tree.heading("new", text="新檔案")
    tree.column("old", anchor="center", width=300)
    tree.column("new", anchor="center", width=300)
    for old, new in data:
        tree.insert("", "end", values=(os.path.basename(old), new))
    tree.pack(fill="both", expand=True)

    # 建立按鈕區域 (無邊框)
    button_frame = tk.Frame(top, bg="#ffffff", bd=0, highlightthickness=0)
    button_frame.grid(row=1, column=0, pady=(10, 20))

    def confirm():
        result["confirmed"] = True
        top.destroy()

    def cancel():
        result["confirmed"] = False
        top.destroy()

    btn_confirm = tk.Button(button_frame, text="確定", command=confirm,
                            font=default_font, bg="#4CAF50", fg="#000000",
                            relief="flat", bd=0, padx=10, pady=5)
    btn_confirm.pack(side="left", padx=10)

    btn_cancel = tk.Button(button_frame, text="取消", command=cancel,
                           font=default_font, bg="#F44336", fg="#000000",
                           relief="flat", bd=0, padx=10, pady=5)
    btn_cancel.pack(side="left", padx=10)

    top.protocol("WM_DELETE_WINDOW", cancel)
    top.lift()
    top.focus_force()
    top.grab_set()
    top.wait_window()
    return result["confirmed"]

# 根據檔案、銷貨單號、起始編號與附加選項計算每個檔案的新名稱（未執行更名）
def generate_rename_list(files, order_number, start_number, add_option):
    rename_info = []  # 每個元素為 (原始檔案路徑, 預計新檔名)
    if '-' not in order_number:
        base_order_number = order_number
        for idx, file in enumerate(files):
            ext = os.path.splitext(file)[1]
            new_name = f"{base_order_number}-{start_number + idx:03d}"
            if add_option == "-@10":
                new_name += "-@10"
            elif add_option == "-New":
                new_name += "-New"
            elif add_option == "-鏡射":
                new_name += "-鏡射"
            elif add_option == "-link":
                new_name += "-link"
            rename_info.append((file, new_name + ext))
    else:
        base_order_number = '-'.join(order_number.split('-')[:-1])
        sub_number = order_number.split('-')[-1]
        for idx, file in enumerate(files):
            ext = os.path.splitext(file)[1]
            new_name = f"{base_order_number}-{sub_number}-{start_number + idx}"
            if add_option == "-@10":
                new_name += "-@10"
            elif add_option == "-New":
                new_name += "-New"
            elif add_option == "-鏡射":
                new_name += "-鏡射"
            elif add_option == "-link":
                new_name += "-link"
            rename_info.append((file, new_name + ext))
    return rename_info

def process_preview(rename_info):
    # 如果你需要顯示預覽資訊，可先處理或記錄 preview_message
    header_old = "舊檔案"
    header_new = "新檔案"
    old_names = [os.path.basename(old) for old, _ in rename_info]
    new_names = [new for _, new in rename_info]
    old_width = max(len(header_old), *(len(name) for name in old_names))
    new_width = max(len(header_new), *(len(name) for name in new_names))

    preview_message = "以下是預覽更名結果：\n\n"
    preview_message += f"{header_old.ljust(old_width)}    {header_new.ljust(new_width)}\n"
    preview_message += "-" * (old_width + new_width + 4) + "\n"
    for old, new in rename_info:
        old_name = os.path.basename(old)
        preview_message += f"{old_name.ljust(old_width)}    {new.ljust(new_width)}\n"

    print(preview_message)  # 如果需要，可以印出預覽訊息

    confirmed = show_preview_window_treeview(rename_info)
    if confirmed:
        apply_renaming(rename_info)

# 真正執行更名的函式
def apply_renaming(rename_info):
    for old, new in rename_info:
        new_path = os.path.join(os.path.dirname(old), new)
        if new_path != old:
            os.rename(old, new_path)
            print(f"檔案 {old} 已更名為 {new}")


# 拖放事件處理函式
def on_drop(event):
    # 先把拖入的檔案處理好後再呼叫預覽窗口（避免在拖放回調中建立模態窗口）
    files = root.tk.splitlist(event.data)
    order_number = entry_order_number.get().strip()

    if not entry_start_number.get().strip():
        messagebox.showerror("錯誤", "未填寫遞增起始編號")
        return
    try:
        start_number = int(entry_start_number.get().strip())
    except ValueError:
        messagebox.showerror("錯誤", "遞增起始編號必須為整數")
        return

    add_option = dropdown_option.get().strip()

    if len(order_number) == 8 and '-' not in order_number:
        base_order_number = order_number
    elif len(order_number) == 12 and order_number[8] == '-':
        base_order_number = order_number
    else:
        messagebox.showerror("錯誤", "銷貨單號格式錯誤")
        return

    rename_info = generate_rename_list(files, base_order_number, start_number, add_option)

    # 為了避免在拖放事件回調中直接建立模態窗口，
    # 利用 root.after(0, …) 延遲調用預覽窗口顯示函式
    root.after(0, lambda: process_preview(rename_info))


# 輸入框焦點事件：改變邊框顏色
def on_focus_in(event):
    event.widget.config(highlightbackground="#6AB7FF", highlightthickness=2)


def on_focus_out(event):
    event.widget.config(highlightbackground="#FFFFFF", highlightthickness=0)


# === GUI 介面設定 ===
root = TkinterDnD.Tk()
root.title("檔案批次更名工具")
window_width = 600
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2) - 100
position_right = int(screen_width / 2 - window_width / 2)
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")
root.config(bg="#F0F0F0")
font = ("Arial", 14, "bold")
title_font = ("Arial", 18, "bold")
title_label = tk.Label(root, text="檔案批次更名工具", font=title_font, bg="#F0F0F0", fg="#333")
title_label.pack(pady=10)
label_order_number = tk.Label(root, text="銷貨單號", bg="#F0F0F0", font=font, fg="#333")
label_order_number.pack(pady=5)
entry_order_number = tk.Entry(root, font=font, bg="#FFFFFF", fg="#333", width=30, bd=0, highlightthickness=0,
                              justify="center")
entry_order_number.pack(pady=5)
entry_order_number.bind("<FocusIn>", on_focus_in)
entry_order_number.bind("<FocusOut>", on_focus_out)
label_start_number = tk.Label(root, text="遞增起始編號:", bg="#F0F0F0", font=font, fg="#333")
label_start_number.pack(pady=5)
entry_start_number = tk.Entry(root, font=font, bg="#FFFFFF", fg="#333", width=30, bd=0, highlightthickness=0,
                              justify="center")
entry_start_number.insert(0, "1")
entry_start_number.pack(pady=5)
entry_start_number.bind("<FocusIn>", on_focus_in)
entry_start_number.bind("<FocusOut>", on_focus_out)
label_option = tk.Label(root, text="附加選項", bg="#F0F0F0", font=font, fg="#333")
label_option.pack(pady=5)
dropdown_option = tk.StringVar(root)
dropdown_option.set("無")
option_menu = tk.OptionMenu(root, dropdown_option, "無", "-@10", "-New", "-鏡射", "-link")
option_menu.config(font=font, bg="#FFFFFF", fg="#333", width=10, bd=0, highlightthickness=0, justify="center")
option_menu.pack(pady=5)
frame = tk.Frame(root, bg="#FFFFFF", bd=0, relief="flat", padx=20, pady=20)
frame.pack(padx=20, pady=20, expand=True)
label_drop_area = tk.Label(frame, text="丟入檔案", font=font, width=40, height=6, relief="flat", bg="#FFFFFF",
                           fg="#333")
label_drop_area.pack(pady=20)
label_drop_area.drop_target_register(DND_FILES)
label_drop_area.dnd_bind("<<Drop>>", lambda event: on_drop(event))
root.mainloop()