import pandas as pd
import keyboard
import pyperclip
import time
import os
import threading
from datetime import datetime

# File names for different categories
FILES = {
    "Leads": "leads.xlsx",
    "Potential": "potential.xlsx",
    "NoResponse": "no_response.xlsx"
}

lock = threading.Lock()

# Create empty files if they do not exist
for category, file_name in FILES.items():
    if not os.path.exists(file_name):
        df = pd.DataFrame(columns=["Date", category])
        df.to_excel(file_name, index=False)

def write_to_excel(category, data):
    with lock:
        file_name = FILES[category]
        df = pd.read_excel(file_name)
        new_data = pd.DataFrame({"Date": [datetime.now().strftime('%Y-%m-%d %H:%M:%S')], category: [data]})
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(file_name, index=False)
        print(f"Added '{data}' to {category}")

def on_key_press(category):
    content = pyperclip.paste()
    write_to_excel(category, content)

def run_listener():
    keyboard.wait()

def main():
    keyboard.add_hotkey("alt+j", lambda: on_key_press("Leads"))
    keyboard.add_hotkey("alt+k", lambda: on_key_press("Potential"))
    keyboard.add_hotkey("alt+l", lambda: on_key_press("NoResponse"))
    print("Listening for clipboard content... Press Alt+J, Alt+K, or Alt+L to categorize data.")
    listener_thread = threading.Thread(target=run_listener, daemon=True)
    listener_thread.start()
    while True:
        time.sleep(1)  # Keep main thread alive

if __name__ == "__main__":
    main()
