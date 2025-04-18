import subprocess
import threading
import time
import requests
import json
from 伺服器.app import app, socketio  # 注意此處根據你的模組結構調整引入路徑


# 定義一個函式啟動 ngrok 隧道
def start_ngrok(port):
    # 假設 ngrok 在系統 PATH 中，否則請使用絕對路徑
    cmd = ["ngrok", "http", str(port)]
    ngrok_proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    # 等待幾秒鐘以確保 ngrok 啟動
    time.sleep(3)
    # 透過 ngrok 的 API 取得公開的 URL（這邊使用 http://127.0.0.1:4040/api/tunnels）
    try:
        tunnels_url = "http://127.0.0.1:4040/api/tunnels"
        tunnels = requests.get(tunnels_url).text
        tunnels_json = json.loads(tunnels)
        public_url = tunnels_json['tunnels'][0]['public_url']
        print("ngrok 公開 URL:", public_url)
    except Exception as e:
        print("無法取得 ngrok 公開 URL:", e)
    return ngrok_proc


def run_server():
    port = 5001  # 根據你的需要，這裡設定你的 Flask 伺服器要使用的 port
    # 啟動 ngrok（放在背景線程中）
    ngrok_thread = threading.Thread(target=start_ngrok, args=(port,), daemon=True)
    ngrok_thread.start()

    # 啟動 Flask/SocketIO 伺服器，注意我們加入 allow_unsafe_werkzeug=True 參數（僅用於測試）
    socketio.run(app, port=port, debug=True, allow_unsafe_werkzeug=True)


if __name__ == '__main__':
    run_server()