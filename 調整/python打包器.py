import sys
import os
import subprocess
import shutil
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QMessageBox, QLineEdit
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class DropArea(QLabel):
    def __init__(self, parent):
        super().__init__("丟入python檔案", parent)
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)
        self.setStyleSheet("""
            QLabel {
                background-color: white;
                color: black;
                font: 18pt "Arial";
                font-weight: bold;
                border: none;
            }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.endswith('.py'):
                self.parent().package_py_file(file_path)
            else:
                QMessageBox.warning(self, "錯誤", "只能拖曳 .py 檔案")


class PackagerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python 應用程式打包工具")
        self.setFixedSize(500, 400)
        self.setStyleSheet("background-color: #f5f5f5;")

        layout = QVBoxLayout()
        layout.setSpacing(20)

        # 標題
        title_label = QLabel("Python 打包器", self)
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: black;")
        layout.addWidget(title_label)

        # 名稱提示文字
        name_title = QLabel("應用程式名稱：", self)
        name_title.setAlignment(Qt.AlignCenter)
        name_title.setStyleSheet("color: black; font-weight: bold; font-size: 16pt;")
        layout.addWidget(name_title)

        # 輸入欄（短一點）
        self.name_input = QLineEdit(self)
        self.name_input.setAlignment(Qt.AlignCenter)
        self.name_input.setFixedSize(250, 36)
        self.name_input.setStyleSheet("""
            background-color: white;
            border: none;
            color: black;
            font-weight: bold;
            font-size: 16pt;
        """)
        layout.addWidget(self.name_input, alignment=Qt.AlignCenter)

        # 拖曳框（高一點）
        self.drop_area = DropArea(self)
        self.drop_area.setFixedSize(300, 200)
        layout.addWidget(self.drop_area, alignment=Qt.AlignCenter)

        self.setLayout(layout)
        self.center_window()

    def center_window(self):
        screen = QApplication.primaryScreen().geometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2 - 100
        )

    def package_py_file(self, file_path):
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        custom_name = self.name_input.text().strip()
        output_name = custom_name if custom_name else base_name

        desktop_path = os.path.expanduser("~/Desktop")

        cmd = [
            "pyinstaller",
            "--noconfirm",
            "--windowed",
            "--name", output_name,
            "--distpath", desktop_path,
            file_path
        ]
        try:
            subprocess.run(cmd, check=True)
            self.clean_build_files(output_name)
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("成功")
            msg_box.setText(f"✅ 打包完成：{output_name}\n\n📁 位置：桌面")
            msg_box.setStyleSheet("QLabel{color:black; font-size:14pt;} QPushButton{color:black;}")
            msg_box.exec_()
            subprocess.run(["open", desktop_path])  # Mac 用 open，Windows 改 explorer

        except subprocess.CalledProcessError as e:
            error_box = QMessageBox(self)
            error_box.setWindowTitle("錯誤")
            error_box.setText(f"打包失敗：\n{e}")
            error_box.setStyleSheet("QLabel{color:black; font-size:14pt;} QPushButton{color:black;}")
            error_box.exec_()

    def clean_build_files(self, output_name):
        for folder in ["build", "__pycache__"]:
            if os.path.exists(folder):
                shutil.rmtree(folder)
        spec_file = f"{output_name}.spec"
        if os.path.exists(spec_file):
            os.remove(spec_file)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PackagerApp()
    window.show()
    sys.exit(app.exec_())