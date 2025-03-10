import tkinter as tk
from tkinter import filedialog, messagebox
import time
import torch
import os
import sys
import whisper
import winshell
from win32com.client import Dispatch
import win32com.client

def create_shortcut():
    try:
        # Get the path to the executable
        if getattr(sys, 'frozen', False):
            application_path = sys.executable
        else:
            application_path = os.path.abspath(__file__)

        # Get desktop path
        desktop = winshell.desktop()
        
        # Path for the shortcut
        shortcut_path = os.path.join(desktop, "Posten Transkriber.lnk")

        # Only create shortcut if it doesn't exist
        if not os.path.exists(shortcut_path):
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = application_path
            shortcut.WorkingDirectory = os.path.dirname(application_path)
            shortcut.IconLocation = application_path
            shortcut.save()
    except Exception as e:
        print(f"Could not create shortcut: {e}")

def get_model_path():
    """Get path to the model file, works both in development and when bundled"""
    if getattr(sys, 'frozen', False):
        # If bundled with PyInstaller
        return os.path.join(sys._MEIPASS, "whisper_model")
    else:
        # During development
        return None

class TranscriberApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Posten Transkriber")
        self.root.geometry("600x500")
        self.file_path = ""
        
        # Create desktop shortcut
        create_shortcut()
        
        # Initialize GUI elements
        self.init_gui()

    def init_gui(self):
        # Browse button
        self.browse_button = tk.Button(
            self.root, 
            text="Last inn lydfil", 
            command=self.browse_file
        )
        self.browse_button.pack(pady=10)

        # File label
        self.file_label = tk.Label(
            self.root, 
            text="Ingen fil valgt", 
            wraplength=500
        )
        self.file_label.pack(pady=5)

        # Run button
        self.run_button = tk.Button(
            self.root, 
            text="Start transkribering", 
            command=self.run_transcription
        )
        self.run_button.pack(pady=10)

        # Text widget for transcription
        self.transcription_text = tk.Text(
            self.root, 
            wrap=tk.WORD, 
            height=15, 
            width=70
        )
        self.transcription_text.pack(pady=10)

        # Copy button
        self.copy_button = tk.Button(
            self.root, 
            text="Kopier til utklippstavle", 
            command=self.copy_to_clipboard
        )
        self.copy_button.pack(pady=10)

        # Time label
        self.time_label = tk.Label(
            self.root, 
            text="Tid brukt: Ikke beregnet ennå"
        )
        self.time_label.pack(pady=5)

        # Status label
        self.status_label = tk.Label(
            self.root, 
            text="Klar", 
            fg="green"
        )
        self.status_label.pack(pady=5)

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[
                ("Lyd/Video filer", 
                "*.mp3 *.wav *.m4a *.mp4 *.avi *.mov *.wmv *.flac")
            ]
        )
        if self.file_path:
            self.file_label.config(
                text=f"Valgt fil: {os.path.basename(self.file_path)}"
            )
            self.status_label.config(text="Klar til å transkribere", fg="green")

    def run_transcription(self):
        if not self.file_path:
            messagebox.showwarning("Advarsel", "Vennligst velg en fil først!")
            return

        try:
            self.status_label.config(text="Laster inn modell...", fg="orange")
            self.root.update()
            
            # Load model with offline support
            model = whisper.load_model("medium", download_root=get_model_path())
            
            self.status_label.config(text="Transkriberer...", fg="orange")
            self.root.update()
            
            # Start timer
            start_time = time.time()
            
            # First detect the language
            initial_result = model.transcribe(
                self.file_path,
                language="no",     # Prefer Norwegian detection
                fp16=False,
                verbose=False
            )
            
            # If the detected language is not Norwegian, translate to Norwegian
            if initial_result.get("language", "").lower() not in ["no", "nor", "norwegian"]:
                result = model.transcribe(
                    self.file_path,
                    task="translate",
                    language="no",
                    fp16=False,
                    verbose=False
                )
            else:
                # If it's already Norwegian, just transcribe normally
                result = model.transcribe(
                    self.file_path,
                    language="no",
                    fp16=False,
                    verbose=False
                )
            
            # Calculate time
            elapsed_time = time.time() - start_time
            
            # Update UI
            self.transcription_text.delete(1.0, tk.END)
            self.transcription_text.insert(tk.END, result["text"])
            self.time_label.config(
                text=f"Tid brukt: {elapsed_time:.2f} sekunder"
            )
            self.status_label.config(text="Transkribering fullført!", fg="green")
            
        except Exception as e:
            self.status_label.config(text="Feil oppstod", fg="red")
            messagebox.showerror("Feil", f"En feil oppstod:\n{str(e)}")

    def copy_to_clipboard(self):
        text = self.transcription_text.get(1.0, tk.END).strip()
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()
            messagebox.showinfo("Suksess", "Tekst kopiert til utklippstavlen!")
        else:
            messagebox.showwarning("Advarsel", "Ingen tekst å kopiere!")

def main():
    try:
        root = tk.Tk()
        app = TranscriberApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Feil ved oppstart av applikasjon: {e}")

if __name__ == "__main__":
    main()