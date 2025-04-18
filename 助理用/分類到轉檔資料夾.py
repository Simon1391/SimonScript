import os
import shutil


def extract_and_move_files_with_txt(folder_path, destination_folder, file_extensions=None):
    if file_extensions is None:
        file_extensions = ['.ai', '.tiff', '.png', '.jpg', '.jpeg', '.gif', '.bmp']  # 需要配對的檔案類型

    # 確保目標資料夾存在
    os.makedirs(destination_folder, exist_ok=True)

    all_files = []  # 儲存已處理的檔案

    # 遍歷來源資料夾及其子資料夾
    for root, dirs, files in os.walk(folder_path):
        # 找出所有符合條件的檔案（如 .ai、.png、.jpg 等）
        move_files = [file for file in files if any(file.endswith(ext) for ext in file_extensions)]
        txt_files = [file for file in files if file.endswith('.txt')]  # 所有 .txt 檔案

        # 如果有需要處理的檔案
        if move_files and txt_files:
            random_txt_path = os.path.join(root, txt_files[0])  # 取一個隨機 .txt 作為模板

            # 針對每個檔案進行處理
            for file in move_files:
                source_path = os.path.join(root, file)
                destination_path = os.path.join(destination_folder, file)

                # **直接覆蓋** 目標資料夾中的同名檔案
                shutil.move(source_path, destination_path)
                all_files.append(destination_path)

                # 複製並重新命名 .txt 檔案
                new_txt_name = os.path.splitext(file)[0] + '.txt'  # 產生新 .txt 檔名
                new_txt_path = os.path.join(destination_folder, new_txt_name)

                try:
                    shutil.copy(random_txt_path, new_txt_path)  # **直接覆蓋已存在的 .txt**
                    all_files.append(new_txt_path)
                except FileNotFoundError as e:
                    print(f"錯誤: 找不到 .txt 檔案 {random_txt_path}，錯誤詳情: {e}")

    return all_files


# 設定來源資料夾和目標資料夾
folder_path = '/Users/onlycolor/Desktop/品名規格資料夾'  # 來源資料夾
destination_folder = '/Users/onlycolor/Desktop/轉檔佇列'  # 目標資料夾

# 執行函數
files = extract_and_move_files_with_txt(folder_path, destination_folder)

# 顯示已處理的檔案
for file in files:
    print(f"已移動: {file}")