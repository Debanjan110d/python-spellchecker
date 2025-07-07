# Python System-wide Spell Checker

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**Created by Debanjan Dutta**

A lightweight, system-wide real-time spell-checking application for Windows. This utility monitors your typing across all applications and provides instant suggestions for misspelled words.

---

## Features

* **System-wide monitoring**: Captures keystrokes in any application.
* **Real-time spell checking**: Under-the-hood use of `pyspellchecker` for language accuracy.
* **Suggestions popup**: Displays a small window with correction suggestions.
* **Clipboard integration**: One-click copying of corrected words directly to the clipboard.
* **Configurable word length**: Ignore short words under a user-defined minimum length.
* **Usage statistics**: Tracks total words checked and misspelled words count.
* **Recent words buffer**: View a list of your most recently typed words.
* **Lightweight GUI**: Built with Tkinter for simplicity and low overhead.

## Requirements

* Windows OS (requires `pywin32` for focus tracking)
* Python 3.7+

### Python Packages

* `pynput`
* `pyspellchecker`
* `pyperclip`
* `pywin32`

All dependencies can be installed automatically with the provided script.

## Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/Debanjan110d/python-spellchecker.git
   cd python-spellchecker
   ```

2. **Install dependencies**

   * Automatic installation via batch script:

     ```bat
     run_spellchecker.bat
     ```
   * Or manually with pip:

     ```bash
     pip install pynput pyspellchecker pyperclip pywin32
     ```

## Usage

1. **Run the application**

   ```bash
   python main.py
   ```
2. **Start Monitoring**

   * In the GUI window, click **Start Monitoring** to begin capturing keystrokes.
3. **Stop Monitoring**

   * Click **Stop Monitoring** to pause spell checking.
4. **Adjust Settings**

   * Change the *minimum word length* in the Settings panel to ignore short words.
5. **View Statistics**

   * Check the *Words checked* and *Misspelled words* counters.
   * See recent words in the text box at the bottom.

## Configuration

* **Minimum Word Length**: Default is 3. Use the spinner control to adjust.
* **Suggestion Throttle**: Suggestions appear at most once per second to avoid interrupting typing.

## File Overview

* `main.py` — Core application logic and GUI.
* `run_spellchecker.bat` — Batch file to install dependencies and launch the app.

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to open a pull request or issue.


---

© 2025 Debanjan Dutta
