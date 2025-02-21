Excel Clipboard Logger

Overview

This script captures clipboard content and logs it into separate Excel files based on the selected category. The user can press designated keyboard shortcuts to save the clipboard content under the appropriate category.

Features

Automatically creates and manages three separate Excel files: leads.xlsx, potential.xlsx, and no_response.xlsx.

Logs data with a timestamp.

Hotkey-based categorization:

Alt+J: Save clipboard content to leads.xlsx.

Alt+K: Save clipboard content to potential.xlsx.

Alt+L: Save clipboard content to no_response.xlsx.

Runs continuously in the background.

Requirements

Python 3.x

Required Python libraries:

pandas

keyboard

pyperclip

openpyxl

To install dependencies, run:

pip install pandas keyboard pyperclip openpyxl

How to Use

Run the script:

python script.py

Copy the desired text to the clipboard.

Press the appropriate hotkey (Alt+J, Alt+K, or Alt+L) to save the content to the respective Excel file.

The script runs in the background and continues to listen for clipboard content until manually stopped.

Stopping the Script

To stop execution, use Ctrl+C in the terminal.

Notes

Ensure you have the required dependencies installed.

Excel files will be created automatically if they do not exist.

The script maintains a timestamp for each entry in the Excel files.

License

This project is open-source and free to use.
