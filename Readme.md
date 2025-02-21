# Dynamic Excel Logger

## Overview

- This Python script allows users to dynamically log clipboard content into Excel files. Users can define multiple categories, assign hotkeys, and select file paths to store their data.

## Features

- Dynamic File and Column Management: Users can assign different Excel files and column names.

- Custom Hotkey Assignments: Any key combination (e.g., alt+j) can be mapped to log clipboard content.

- Graphical User Interface (GUI): Users can add new files and hotkeys on demand.

- Multi-threaded Execution: The script runs a background listener while keeping the GUI active.

- Timestamped Entries: Each log entry includes the date and time.

## Requirements

```
 Python 3.x
```

## Required Python libraries:

```
- pandas

- keyboard

- pyperclip

- openpyxl

- tkinter
```

## To install dependencies, run:

- pip install pandas keyboard pyperclip openpyxl

## How to Use

#### Run the script:

- python script.py

#### Set up logging:

- Enter a column name.

- Select an Excel file to store the data.

- Assign a hotkey (e.g., alt+j).

#### Log clipboard content:

- Copy text to the clipboard.

- Press the assigned hotkey to store it in the selected Excel file.

#### Add more categories:

- The GUI remains open so users can add more categories and hotkeys.

- Exit:

- Click the Close button to exit the GUI, but the background logging will continue.

### Stopping the Script

```
 To stop execution, use Ctrl+C in the terminal.
```

#### Notes

- Ensure you have the required dependencies installed.

- Excel files will be created automatically if they do not exist.

- The script maintains a timestamp for each entry in the Excel files.

### License

- This project is open-source and free to use.
