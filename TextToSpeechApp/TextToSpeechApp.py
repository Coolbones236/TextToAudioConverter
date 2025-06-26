import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyttsx3
import pygame
import threading
import time
import os
from pydub import AudioSegment
import docx
import PyPDF2
import pypandoc

def extract_text_from_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    elif ext == ".docx":
        doc = docx.Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    elif ext == ".pdf":
        text = ""
        with open(path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text += page.extract_text() or ''
        return text
    elif ext in [".rtf", ".odt"]:
        return pypandoc.convert_file(path, 'plain')
    else:
        raise ValueError(f"Unsupported file format: {ext}")

class TextToSpeechApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("TTS Audio")
        self.geometry("520x620")
        self.configure(bg="#1e1e1e")
        self.resizable(False, False)

        self.engine = pyttsx3.init()
        self.voices = self.engine.getProperty("voices")
        self.theme = "dark"
        self.audio_file = "output_audio.wav"
        self.is_playing = False
        self.is_paused = False
        self.duration = 0

        pygame.mixer.init()

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.on_drop_file)

        self.create_widgets()
        self.apply_theme()

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('default')
        style.configure("TButton", relief="flat", padding=6)

        top_frame = tk.Frame(self, bg=self.cget("bg"))
        top_frame.pack(pady=10)

        self.theme_btn = ttk.Button(top_frame, text="🌓", command=self.toggle_theme)
        self.theme_btn.pack(side=tk.LEFT, padx=4)

        self.voice_var = tk.StringVar()
        self.voice_combo = ttk.Combobox(top_frame, textvariable=self.voice_var, width=25)
        self.voice_combo['values'] = [f"{i}: {v.name}" for i, v in enumerate(self.voices)]
        self.voice_combo.current(0)
        self.voice_combo.pack(side=tk.LEFT, padx=4)
        self.voice_combo.tooltip = tk.Label(self, text="Select voice", bg=self.cget("bg"), fg="gray")

        self.format_var = tk.StringVar(value="wav")
        self.format_combo = ttk.Combobox(top_frame, textvariable=self.format_var, values=["wav", "mp3"], width=5)
        self.format_combo.pack(side=tk.LEFT, padx=4)

        self.browse_btn = ttk.Button(top_frame, text="📂", width=3, command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT, padx=4)

        self.convert_btn = ttk.Button(self, text="🎧 Convert & Play", command=self.convert_text)
        self.convert_btn.pack(pady=10)

        self.text_entry = tk.Text(self, height=12, width=60, wrap=tk.WORD)
        self.text_entry.pack(padx=10, pady=5)

        self.progress_bar = ttk.Progressbar(self, length=480, mode='determinate')
        self.progress_bar.pack(pady=6)

        self.progress_label = tk.Label(self, text="Progress: 0%", bg=self.cget("bg"), fg="gray")
        self.progress_label.pack()

        control_frame = tk.Frame(self, bg=self.cget("bg"))
        control_frame.pack(pady=8)

        self.play_btn = ttk.Button(control_frame, text="▶️", command=self.play_audio, width=4, state="disabled")
        self.pause_btn = ttk.Button(control_frame, text="⏸️", command=self.pause_audio, width=4, state="disabled")
        self.resume_btn = ttk.Button(control_frame, text="🔁", command=self.resume_audio, width=4, state="disabled")
        self.stop_btn = ttk.Button(control_frame, text="⏹️", command=self.stop_audio, width=4, state="disabled")

        self.play_btn.pack(side=tk.LEFT, padx=5)
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        self.resume_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(self, text="", bg=self.cget("bg"), fg="gray")
        self.status_label.pack(pady=6)

    def apply_theme(self):
        bg = "#1e1e1e" if self.theme == "dark" else "#f9f9f9"
        fg = "#ffffff" if self.theme == "dark" else "#000000"
        self.configure(bg=bg)
        for widget in self.winfo_children():
            try:
                widget.configure(bg=bg, fg=fg)
            except:
                pass

    def toggle_theme(self):
        self.theme = "light" if self.theme == "dark" else "dark"
        self.apply_theme()

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt *.docx *.pdf *.rtf *.odt")])
        if path:
            self.load_file(path)

    def load_file(self, path):
        try:
            text = extract_text_from_file(path)
            self.text_entry.delete("1.0", tk.END)
            self.text_entry.insert(tk.END, text.strip())
        except Exception as e:
            messagebox.showerror("File Error", str(e))

    def on_drop_file(self, event):
        paths = self.tk.splitlist(event.data)
        content = ""
        for path in paths:
            if os.path.isfile(path):
                try:
                    content += extract_text_from_file(path).strip() + "\n\n"
                except Exception as e:
                    messagebox.showerror("Drop Error", str(e))
        if content:
            self.text_entry.delete("1.0", tk.END)
            self.text_entry.insert(tk.END, content.strip())

    def convert_text(self):
        text = self.text_entry.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Missing Text", "Please input or drop a file.")
            return

        voice_index = int(self.voice_var.get().split(":")[0])
        self.engine.setProperty("voice", self.voices[voice_index].id)

        fmt = self.format_var.get()
        save_path = filedialog.asksaveasfilename(defaultextension=f".{fmt}", filetypes=[(f"{fmt.upper()} files", f"*.{fmt}")])
        if not save_path:
            return

        self.engine.save_to_file(text, "temp_audio.wav")
        self.engine.runAndWait()

        if fmt == "mp3":
            try:
                audio = AudioSegment.from_wav("temp_audio.wav")
                audio.export(save_path, format="mp3")
                os.remove("temp_audio.wav")
            except Exception:
                messagebox.showerror("MP3 Error", "Ensure FFmpeg is installed and on PATH.")
                return
        else:
            os.rename("temp_audio.wav", save_path)

        self.audio_file = save_path
        self.play_btn.config(state="normal")
        self.status_label.config(text="✅ Audio ready.")

    def play_audio(self):
        if not os.path.exists(self.audio_file):
            return
        self.is_playing = True
        self.is_paused = False

        pygame.mixer.music.load(self.audio_file)
        pygame.mixer.music.play()

        self.pause_btn.config(state="normal")
        self.resume_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.status_label.config(text="▶️ Playing...")

        sound = pygame.mixer.Sound(self.audio_file)
        self.duration = sound.get_length()
        start = time.time()
        while pygame.mixer.music.get_busy() or self.is_paused:
            self.update()
            if not self.is_paused:
                self.update_progress(start)
            time.sleep(0.1)

        self.status_label.config(text="✅ Finished")
        self.reset_controls()

    def pause_audio(self):
        pygame.mixer.music.pause()
        self.is_paused = True
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="normal")

    def resume_audio(self):
        pygame.mixer.music.unpause()
        self.is_paused = False
        self.pause_btn.config(state="normal")
        self.resume_btn.config(state="disabled")

    def stop_audio(self):
        pygame.mixer.music.stop()
        self.reset_controls()
        self.status_label.config(text="⏹️ Stopped")

    def update_progress(self, start):
        elapsed = time.time() - start
        percent = min((elapsed / self.duration) * 100, 100)
        self.progress_bar['value'] = percent
        self.progress_label.config(text=f"Progress: {int(percent)}%")

    def reset_controls(self):
        self.is_playing = False
        self.is_paused = False
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="disabled")
        self.stop_btn.config(state="disabled")
        self.progress_bar['value'] = 0
        self.progress_label.config(text="Progress: 0%")

if __name__ == "__main__":
    app = TextToSpeechApp()
    app.mainloop()
