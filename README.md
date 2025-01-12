# Python Screenshot Tool for Manual Testers

A simple and effective Python-based screenshot tool designed for manual testers to easily capture and save screenshots into a Word document during software testing. This tool allows testers to create a new document, add titles, and include screenshots for documentation and bug reporting.

## Features

- **Capture Screenshots**: Easily capture screenshots of the entire screen using the `Print Screen` key.
- **Create New Document**: Automatically create a new Word document to store screenshots and titles.
- **Add Titles**: Input titles for each screenshot to provide context or description.
- **Save Screenshots in Document**: Screenshots are saved directly into the Word document for easy reference.
- **Organized Folder Structure**: Automatically creates a timestamped folder to store the document and temporary screenshot files.

> **Note:** This tool currently supports creating a new document and adding screenshots. Advanced features like editing existing documents or additional annotation tools may be added in future versions.

## Requirements

- Python 3.x
- Libraries:
  - `python-docx` (for Word document manipulation)
  - `pyautogui` (for screenshot capture)
  - `pynput` (for keyboard monitoring)
  - `tkinter` (for GUI components)
  - `Pillow` (image processing)

Install the required libraries using:

```bash
pip install -r requirements.txt

