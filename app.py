import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import os
import queue
import threading

from config_manager import ConfigManager
from automation_worker import AutomationWorker

class ToolTip(ctk.CTkToplevel):
    def __init__(self, widget, text):
        super().__init__(widget)
        self.widget = widget
        self.text = text
        self.withdraw()
        self.overrideredirect(True)
        self.label = ctk.CTkLabel(self, text=self.text, corner_radius=5, fg_color="#404040", text_color="white", padx=10, pady=5)
        self.label.pack()
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.geometry(f"+{x}+{y}")
        self.deiconify()

    def hide_tip(self, event):
        self.withdraw()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Pay Period Report Automation")
        self.geometry("800x650")

        self.config = ConfigManager()
        self.status_queue = queue.Queue()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme(self.config.get('Application', 'theme').lower())
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.settings_frame.grid_columnconfigure(1, weight=1)
        self.create_settings_widgets()

        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.create_log_widgets()

        self.load_settings()

    def create_settings_widgets(self):
        ctk.CTkLabel(self.settings_frame, text="Email Address:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.email_entry = ctk.CTkEntry(self.settings_frame, width=300)
        self.email_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ToolTip(self.email_entry, "The email address of the inbox to monitor.")

        ctk.CTkLabel(self.settings_frame, text="Root Folder Path:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.root_path_entry = ctk.CTkEntry(self.settings_frame, width=300)
        self.root_path_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.browse_button = ctk.CTkButton(self.settings_frame, text="Browse...", command=self.browse_folder)
        self.browse_button.grid(row=1, column=2, padx=10, pady=5)
        ToolTip(self.root_path_entry, "The main folder where all Pay Period subfolders will be created.")

        ctk.CTkLabel(self.settings_frame, text="Pay Period File:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.pp_file_entry = ctk.CTkEntry(self.settings_frame, width=300)
        self.pp_file_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.browse_pp_button = ctk.CTkButton(self.settings_frame, text="Browse...", command=self.browse_pp_file)
        self.browse_pp_button.grid(row=2, column=2, padx=10, pady=5)
        ToolTip(self.pp_file_entry, "The CSV file containing the pay period schedule.")

        self.save_button = ctk.CTkButton(self.settings_frame, text="Save Settings", command=self.save_settings)
        self.save_button.grid(row=3, column=1, columnspan=2, padx=10, pady=10, sticky="e")
    def create_log_widgets(self):
        self.start_button = ctk.CTkButton(self.log_frame, text="Start Automation", command=self.start_automation, height=40, font=ctk.CTkFont(size=14, weight="bold"))
        self.start_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ToolTip(self.start_button, "Connects to the email inbox and starts processing reports.")
        self.log_textbox = ctk.CTkTextbox(self.log_frame, state="disabled", wrap="word")
        self.log_textbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    def log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path: self.root_path_entry.delete(0, "end"); self.root_path_entry.insert(0, folder_path)

    def browse_pp_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path: self.pp_file_entry.delete(0, "end"); self.pp_file_entry.insert(0, file_path)

    def load_settings(self):
        self.email_entry.insert(0, self.config.get('Email', 'address'))
        self.root_path_entry.insert(0, self.config.get('Folders', 'root_path'))
        self.pp_file_entry.insert(0, self.config.get('Folders', 'pay_period_schedule_csv'))

    def save_settings(self):
        self.config.set('Email', 'address', self.email_entry.get())
        self.config.set('Folders', 'root_path', self.root_path_entry.get())
        self.config.set('Folders', 'pay_period_schedule_csv', self.pp_file_entry.get())
        self.config.save_config()
        messagebox.showinfo("Success", "Settings have been saved successfully!")

    def start_automation(self):
        self.save_settings()
        password = simpledialog.askstring("Password", "Please enter the email password:", show='*')
        if not password:
            messagebox.showwarning("Cancelled", "Password not provided. Automation cancelled.")
            return
        self.log_textbox.configure(state="normal"); self.log_textbox.delete("1.0", "end"); self.log_textbox.configure(state="disabled")
        self.start_button.configure(state="disabled", text="Running...")
        self.worker = AutomationWorker(self.config, password, self.status_queue)
        self.thread = threading.Thread(target=self.worker.run, daemon=True)
        self.thread.start()
        self.after(100, self.check_queue)

    def check_queue(self):
        try:
            message = self.status_queue.get_nowait()
            if message == "DONE":
                self.start_button.configure(state="normal", text="Start Automation")
            else:
                self.log_message(message)
            self.after(100, self.check_queue)
        except queue.Empty:
            if self.thread.is_alive():
                self.after(100, self.check_queue)
            else: # Thread finished but no DONE message might indicate a crash
                self.start_button.configure(state="normal", text="Start Automation")

if __name__ == "__main__":
    if not os.path.exists('config.ini'):
         messagebox.showerror("Error", "config.ini not found! Please create it before running the application.")
    else:
        from datetime import datetime
        app = App()
        app.mainloop()
