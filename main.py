#!/usr/bin/env python3
"""
System-wide Spell Checker
Monitors typing across all applications and provides real-time spell checking
with suggestions for misspelled words.
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk
import pynput
from pynput import keyboard
import pyperclip
from spellchecker import SpellChecker
import re
from collections import deque
import win32gui
import win32con
from difflib import get_close_matches

class SpellCheckerApp:
    def __init__(self):
        self.spell = SpellChecker()
        self.current_word = ""
        self.word_buffer = deque(maxlen=50)  # Store recent words
        self.suggestion_window = None
        self.is_monitoring = False
        self.last_suggestion_time = 0
        self.min_word_length = 3
        
        # Create main control window
        self.root = tk.Tk()
        self.root.title("System Spell Checker")
        self.root.geometry("400x300")
        self.root.attributes('-topmost', True)
        
        self.setup_ui()
        self.keyboard_listener = None
        
    def setup_ui(self):
        """Setup the main control interface"""
        # Title
        title_label = tk.Label(self.root, text="System-wide Spell Checker", 
                              font=('Arial', 14, 'bold'))
        title_label.pack(pady=10)
        
        # Status
        self.status_label = tk.Label(self.root, text="Status: Stopped", 
                                    font=('Arial', 10))
        self.status_label.pack(pady=5)
        
        # Control buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.start_button = tk.Button(button_frame, text="Start Monitoring", 
                                     command=self.start_monitoring, 
                                     bg='green', fg='white', font=('Arial', 10, 'bold'))
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = tk.Button(button_frame, text="Stop Monitoring", 
                                    command=self.stop_monitoring, 
                                    bg='red', fg='white', font=('Arial', 10, 'bold'))
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Settings
        settings_frame = tk.LabelFrame(self.root, text="Settings", font=('Arial', 10, 'bold'))
        settings_frame.pack(pady=10, padx=20, fill='x')
        
        # Minimum word length
        tk.Label(settings_frame, text="Minimum word length:").pack(anchor='w')
        self.min_length_var = tk.StringVar(value=str(self.min_word_length))
        length_spinbox = tk.Spinbox(settings_frame, from_=2, to=10, 
                                   textvariable=self.min_length_var, 
                                   command=self.update_min_length)
        length_spinbox.pack(anchor='w', pady=2)
        
        # Statistics
        stats_frame = tk.LabelFrame(self.root, text="Statistics", font=('Arial', 10, 'bold'))
        stats_frame.pack(pady=10, padx=20, fill='x')
        
        self.words_checked_label = tk.Label(stats_frame, text="Words checked: 0")
        self.words_checked_label.pack(anchor='w')
        
        self.misspelled_label = tk.Label(stats_frame, text="Misspelled words: 0")
        self.misspelled_label.pack(anchor='w')
        
        # Recent words
        self.recent_words_text = tk.Text(stats_frame, height=4, width=40)
        self.recent_words_text.pack(pady=5)
        
        # Initialize counters
        self.words_checked = 0
        self.misspelled_count = 0
        
    def update_min_length(self):
        """Update minimum word length setting"""
        try:
            self.min_word_length = int(self.min_length_var.get())
        except ValueError:
            self.min_word_length = 3
            
    def start_monitoring(self):
        """Start the keyboard monitoring"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.status_label.config(text="Status: Monitoring...", fg='green')
            self.start_button.config(state='disabled')
            self.stop_button.config(state='normal')
            
            # Start keyboard listener in a separate thread
            self.keyboard_listener = keyboard.Listener(on_press=self.on_key_press)
            self.keyboard_listener.start()
            
    def stop_monitoring(self):
        """Stop the keyboard monitoring"""
        if self.is_monitoring:
            self.is_monitoring = False
            self.status_label.config(text="Status: Stopped", fg='red')
            self.start_button.config(state='normal')
            self.stop_button.config(state='disabled')
            
            if self.keyboard_listener:
                self.keyboard_listener.stop()
                
            if self.suggestion_window:
                self.suggestion_window.destroy()
                self.suggestion_window = None
                
    def on_key_press(self, key):
        """Handle keyboard input"""
        if not self.is_monitoring:
            return
            
        try:
            # Handle regular characters
            if hasattr(key, 'char') and key.char is not None:
                char = key.char
                if char.isalpha():
                    self.current_word += char.lower()
                elif char in [' ', '\n', '\t'] and self.current_word:
                    self.check_word(self.current_word)
                    self.current_word = ""
                else:
                    if self.current_word:
                        self.check_word(self.current_word)
                        self.current_word = ""
                        
            # Handle special keys
            elif key == keyboard.Key.space or key == keyboard.Key.enter:
                if self.current_word:
                    self.check_word(self.current_word)
                    self.current_word = ""
                    
            elif key == keyboard.Key.backspace:
                if self.current_word:
                    self.current_word = self.current_word[:-1]
                    
        except Exception as e:
            pass  # Ignore errors to prevent interrupting typing
            
    def check_word(self, word):
        """Check if word is spelled correctly"""
        if len(word) < self.min_word_length:
            return
            
        # Clean the word
        word = re.sub(r'[^a-zA-Z]', '', word).lower()
        if not word:
            return
            
        # Update statistics
        self.words_checked += 1
        self.words_checked_label.config(text=f"Words checked: {self.words_checked}")
        
        # Add to recent words
        self.word_buffer.append(word)
        self.update_recent_words_display()
        
        # Check spelling
        if word not in self.spell:
            self.misspelled_count += 1
            self.misspelled_label.config(text=f"Misspelled words: {self.misspelled_count}")
            
            # Get suggestions
            suggestions = list(self.spell.candidates(word))
            if not suggestions:
                suggestions = get_close_matches(word, self.spell.word_frequency.keys(), n=3, cutoff=0.6)
                
            # Show suggestion window
            current_time = time.time()
            if current_time - self.last_suggestion_time > 1.0:  # Throttle suggestions
                self.show_suggestion(word, suggestions)
                self.last_suggestion_time = current_time
                
    def update_recent_words_display(self):
        """Update the display of recent words"""
        self.recent_words_text.delete(1.0, tk.END)
        recent_words = list(self.word_buffer)
        if recent_words:
            self.recent_words_text.insert(tk.END, ", ".join(recent_words[-20:]))
            
    def show_suggestion(self, misspelled_word, suggestions):
        """Show suggestion window"""
        if self.suggestion_window:
            self.suggestion_window.destroy()
            
        if not suggestions:
            return
            
        # Create suggestion window
        self.suggestion_window = tk.Toplevel(self.root)
        self.suggestion_window.title("Spell Check Suggestion")
        self.suggestion_window.geometry("300x150")
        self.suggestion_window.attributes('-topmost', True)
        
        # Position near cursor (approximate)
        self.suggestion_window.geometry("+200+200")
        
        # Content
        tk.Label(self.suggestion_window, 
                text=f"Misspelled: '{misspelled_word}'", 
                font=('Arial', 10, 'bold'), fg='red').pack(pady=5)
        
        tk.Label(self.suggestion_window, 
                text="Suggestions:", 
                font=('Arial', 9)).pack()
        
        # Create buttons for suggestions
        for i, suggestion in enumerate(suggestions[:3]):
            btn = tk.Button(self.suggestion_window, 
                           text=suggestion, 
                           command=lambda s=suggestion: self.copy_suggestion(s))
            btn.pack(pady=2)
            
        # Auto-close after 3 seconds
        self.suggestion_window.after(3000, self.close_suggestion_window)
        
    def copy_suggestion(self, suggestion):
        """Copy suggestion to clipboard"""
        try:
            pyperclip.copy(suggestion)
            if self.suggestion_window:
                # Show brief confirmation
                tk.Label(self.suggestion_window, 
                        text=f"Copied '{suggestion}' to clipboard!", 
                        fg='green').pack()
                self.suggestion_window.after(1000, self.close_suggestion_window)
        except Exception as e:
            pass
            
    def close_suggestion_window(self):
        """Close the suggestion window"""
        if self.suggestion_window:
            self.suggestion_window.destroy()
            self.suggestion_window = None
            
    def on_closing(self):
        """Handle window closing"""
        self.stop_monitoring()
        self.root.destroy()
        
    def run(self):
        """Start the application"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

def install_dependencies():
    """Install required dependencies"""
    dependencies = [
        'pynput',
        'pyspellchecker', 
        'pyperclip',
        'pywin32'
    ]
    
    print("Installing required dependencies...")
    for dep in dependencies:
        try:
            os.system(f'pip install {dep}')
            print(f"✓ {dep} installed successfully")
        except Exception as e:
            print(f"✗ Failed to install {dep}: {e}")
            
    print("\nDependencies installation complete!")

if __name__ == "__main__":
    print("System-wide Spell Checker")
    print("=" * 30)
    
    # Check if running with admin privileges (recommended for system-wide monitoring)
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        if not is_admin:
            print("⚠️  Note: Running without admin privileges. Some applications may not be monitored.")
    except:
        pass
    
    # Install dependencies if needed
    try:
        import pynput
        from spellchecker import SpellChecker
        import pyperclip
        import win32gui
    except ImportError:
        print("Missing dependencies. Installing...")
        install_dependencies()
        print("Please restart the application after installation.")
        sys.exit(1)
    
    # Start the application
    app = SpellCheckerApp()
    app.run()
