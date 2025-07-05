import tkinter as tk
from tkinter import messagebox, Text, font
import json
import win32com.client as win32
import threading
import os

class EmailApp:
    """
    A desktop application for sending batch emails via Microsoft Outlook.
    """
    def __init__(self, root):
        """
        Initializes the application's user interface and loads saved data.
        """
        self.root = root
        self.root.title("Email Automation Tool")
        self.root.geometry("850x750")

        # --- Modern UI Styling ---
        self.colors = {
            'bg': '#2B2B2B',
            'frame_bg': '#3C3C3C',
            'text': '#E0E0E0',
            'button_bg': '#007ACC',
            'button_fg': '#FFFFFF',
            'button_hover': '#005F9E',
            'entry_bg': '#505050',
            'border': '#555555'
        }
        self.root.configure(bg=self.colors['bg'])
        self.font_normal = font.Font(family="Segoe UI", size=10)
        self.font_bold = font.Font(family="Segoe UI", size=11, weight="bold")

        # --- Data Storage ---
        self.data_file = "email_data.json"
        self.to_emails = []
        self.cc_emails = []
        self.load_data()

        # --- UI Setup ---
        self.create_widgets()
        self.populate_fields()
        
        # --- Cancellation Flag ---
        self.cancel_sending = False

    def create_widgets(self):
        """
        Creates and arranges all the UI components in the main window.
        """
        # --- Main Frames ---
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(padx=15, pady=15, fill="both", expand=True)

        main_frame.grid_rowconfigure(0, weight=1) 
        main_frame.grid_rowconfigure(1, weight=5) 
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # --- "To" Emails Frame (with Scrollbar) ---
        to_frame_container = tk.LabelFrame(main_frame, text="Receiver's Email Addresses (To)", 
                                           bg=self.colors['frame_bg'], fg=self.colors['text'], 
                                           font=self.font_bold, relief='flat', borderwidth=0)
        to_frame_container.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        to_frame_container.grid_rowconfigure(0, weight=1)
        to_frame_container.grid_columnconfigure(0, weight=1)

        to_canvas = tk.Canvas(to_frame_container, borderwidth=0, bg=self.colors['frame_bg'], highlightthickness=0)
        to_scrollbar = tk.Scrollbar(to_frame_container, orient="vertical", command=to_canvas.yview, 
                                    bg=self.colors['frame_bg'], troughcolor=self.colors['bg'], 
                                    activebackground=self.colors['button_hover'], relief='flat')
        to_scrollable_frame = tk.Frame(to_canvas, bg=self.colors['frame_bg'])

        to_scrollable_frame.bind("<Configure>", lambda e: to_canvas.configure(scrollregion=to_canvas.bbox("all")))
        
        to_canvas.create_window((0, 0), window=to_scrollable_frame, anchor="nw")
        to_canvas.configure(yscrollcommand=to_scrollbar.set)
        
        to_canvas.grid(row=0, column=0, sticky="nsew")
        to_scrollbar.grid(row=0, column=1, sticky="ns")
        to_scrollable_frame.grid_columnconfigure(0, weight=1)

        self.to_entries = []
        for i in range(24):
            entry = tk.Entry(to_scrollable_frame, width=30, bg=self.colors['entry_bg'], fg=self.colors['text'],
                             relief='flat', font=self.font_normal, insertbackground=self.colors['text'],
                             highlightthickness=1, highlightbackground=self.colors['border'])
            entry.grid(row=i, column=0, padx=(10,15), pady=3, sticky="ew")
            self.to_entries.append(entry)

        # --- "CC" Emails Frame ---
        cc_frame = tk.LabelFrame(main_frame, text="CC Email Addresses", 
                                 bg=self.colors['frame_bg'], fg=self.colors['text'], 
                                 font=self.font_bold, relief='flat', borderwidth=0)
        cc_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        cc_frame.grid_columnconfigure(0, weight=1)

        self.cc_entries = []
        for i in range(6):
            entry = tk.Entry(cc_frame, width=30, bg=self.colors['entry_bg'], fg=self.colors['text'],
                             relief='flat', font=self.font_normal, insertbackground=self.colors['text'],
                             highlightthickness=1, highlightbackground=self.colors['border'])
            entry.grid(row=i, column=0, padx=10, pady=3, sticky="ew")
            self.cc_entries.append(entry)

        # --- Email Body Frame ---
        body_frame = tk.LabelFrame(main_frame, text="Email Body", 
                                   bg=self.colors['frame_bg'], fg=self.colors['text'], 
                                   font=self.font_bold, relief='flat', borderwidth=0)
        body_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        body_frame.grid_rowconfigure(0, weight=1)
        body_frame.grid_columnconfigure(0, weight=1)

        self.email_body = Text(body_frame, width=80, bg=self.colors['entry_bg'], fg=self.colors['text'],
                               relief='flat', font=self.font_normal, insertbackground=self.colors['text'],
                               highlightthickness=1, highlightbackground=self.colors['border'], wrap='word')
        self.email_body.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # --- Buttons Frame ---
        button_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        button_frame.grid(row=2, column=0, columnspan=2, pady=15)

        self.create_modern_button(button_frame, "Send in Batch", self.start_sending_thread).grid(row=0, column=0, padx=10)
        self.create_modern_button(button_frame, "Cancel", self.cancel_send).grid(row=0, column=1, padx=10)
        self.create_modern_button(button_frame, "Clear", self.clear_fields).grid(row=0, column=2, padx=10)

    def create_modern_button(self, parent, text, command):
        """Helper function to create styled buttons with hover effects."""
        button = tk.Button(parent, text=text, command=command,
                           bg=self.colors['button_bg'], fg=self.colors['button_fg'],
                           font=self.font_bold, relief='flat', borderwidth=0,
                           activebackground=self.colors['button_hover'],
                           activeforeground=self.colors['button_fg'],
                           pady=5, padx=10)
        
        button.bind("<Enter>", lambda e: e.widget.config(bg=self.colors['button_hover']))
        button.bind("<Leave>", lambda e: e.widget.config(bg=self.colors['button_bg']))
        return button

    def load_data(self):
        """
        Loads email addresses from the JSON data file.
        """
        if os.path.exists(self.data_file):
            with open(self.data_file, 'r') as f:
                try:
                    data = json.load(f)
                    self.to_emails = data.get("to_emails", [])
                    self.cc_emails = data.get("cc_emails", [])
                except json.JSONDecodeError:
                    self.to_emails, self.cc_emails = [], []

    def save_data(self):
        """
        Saves the current email addresses to the JSON data file.
        """
        data = {
            "to_emails": [entry.get() for entry in self.to_entries],
            "cc_emails": [entry.get() for entry in self.cc_entries]
        }
        with open(self.data_file, 'w') as f:
            json.dump(data, f, indent=4)

    def populate_fields(self):
        """
        Fills the entry fields with the loaded email addresses.
        """
        for i, email in enumerate(self.to_emails):
            if i < len(self.to_entries):
                self.to_entries[i].insert(0, email)

        for i, email in enumerate(self.cc_emails):
            if i < len(self.cc_entries):
                self.cc_entries[i].insert(0, email)

    def start_sending_thread(self):
        """
        Starts the email sending process in a separate thread to keep the UI responsive.
        """
        self.save_data()
        self.cancel_sending = False
        
        send_thread = threading.Thread(target=self.send_emails)
        send_thread.start()

    def send_emails(self):
        """
        Iterates through the recipient list and sends emails via Outlook.
        """
        try:
            outlook = win32.Dispatch('outlook.application')
        except Exception as e:
            messagebox.showerror("Outlook Error", f"Microsoft Outlook is not running or could not be started.\nError: {e}")
            return

        to_list = [entry.get() for entry in self.to_entries if entry.get()]
        cc_list = ";".join([entry.get() for entry in self.cc_entries if entry.get()])
        email_content = self.email_body.get("1.0", "end-1c")

        if not to_list:
            messagebox.showwarning("Input Error", "Please enter at least one recipient.")
            return

        for recipient in to_list:
            if self.cancel_sending:
                messagebox.showinfo("Cancelled", "Email sending has been cancelled.")
                break
            
            try:
                mail = outlook.CreateItem(0)
                mail.To = recipient
                mail.CC = cc_list
                mail.Subject = "Your Subject Here" 
                mail.Body = email_content
                mail.Send()
            except Exception as e:
                messagebox.showerror("Email Error", f"Could not send email to {recipient}.\nError: {e}")
        else: 
            if not self.cancel_sending:
                messagebox.showinfo("Success", "All emails have been sent successfully.")

    def cancel_send(self):
        """
        Sets the cancellation flag to stop the sending process.
        """
        self.cancel_sending = True

    def clear_fields(self):
        """
        Clears all input fields in the application.
        """
        if messagebox.askokcancel("Confirm Clear", "Are you sure you want to clear all fields?"):
            for entry in self.to_entries:
                entry.delete(0, 'end')
            for entry in self.cc_entries:
                entry.delete(0, 'end')
            self.email_body.delete("1.0", 'end')
            self.save_data()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
