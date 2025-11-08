import webview
import threading
import subprocess
import time
import os
import sys
import requests

# Fix path for PyInstaller
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

os.environ["FLASK_APP_BASE"] = base_path

def start_flask():
    print("Starting Flask server...")
    subprocess.Popen([sys.executable, "app.py"], stdout=sys.stdout, stderr=sys.stderr)

def wait_for_server(url, timeout=10):
    """Wait for the Flask server to start up."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            response = requests.get(url)
            if response.status_code == 200:
                return True
        except requests.ConnectionError:
            pass
        time.sleep(0.5)
    return False

if __name__ == '__main__':
    # Start Flask in a background thread
    flask_thread = threading.Thread(target=start_flask)
    flask_thread.daemon = True
    flask_thread.start()

    print("Waiting for Flask to start...")
    if wait_for_server("http://127.0.0.1:5000"):
        print("Flask server is up. Launching UI.")
        webview.create_window("Paysly Tool", "http://127.0.0.1:5000")
        webview.start()
    else:
        print("Error: Flask server did not start.")







