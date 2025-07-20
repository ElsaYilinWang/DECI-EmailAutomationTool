import tkinter as tk
from tkinter import messagebox, scrolledtext
import win32com.client
import os
import json
import logging
import re
from logging.handlers import RotatingFileHandler

# --- Configuration ---
APP_NAME = "Email Automation Tool v1.8"
CONFIG_FILE = os.path.join(os.getenv('APPDATA'), 'EmailAutomationTool', 'config.json')
LOG_DIR = os.path.join(os.getenv('APPDATA'), 'EmailAutomationTool', 'logs')

# --- Setup Logging ---
os.makedirs(LOG_DIR, exist_ok=True)
log_file_path = os.path.join(LOG_DIR, 'app_log.log')
json_log_file_path = os.path.join(LOG_DIR, 'app_log.jsonl')

# Human-readable log
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler = RotatingFileHandler(log_file_path, maxBytes=1024*1024, backupCount=5)
log_handler.setFormatter(log_formatter)

# JSON log for database import
json_log_formatter = logging.Formatter('{"timestamp": "%(asctime)s", "level": "%(levelname)s", "message": "%(message)s"}')
json_log_handler = RotatingFileHandler(json_log_file_path, maxBytes=1024*1024, backupCount=5)
json_log_handler.setFormatter(json_log_formatter)


logger = logging.getLogger()
logger.setLevel(logging.INFO)
if not logger.handlers:
    logger.addHandler(log_handler)
    logger.addHandler(json_log_handler)


class EmailApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("650x700")
        self.configure(bg="#f0f0f0")

        # --- Variables ---
        self.sender_email_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.default_cc_var = tk.StringVar(value="qa.team@deci-ltd.com") # Default CC

        # --- Main Frame ---
        main_frame = tk.Frame(self, bg="#f0f0f0", padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # --- Sender and Subject Section ---
        header_frame = tk.Frame(main_frame, bg="#f0f0f0")
        header_frame.pack(fill="x", pady=(0, 15))
        
        # NEW: Sender's Email Field
        tk.Label(header_frame, text="Sender's Email (leave blank for default):", font=("Segoe UI", 10, "bold"), bg="#f0f0f0").pack(fill="x")
        self.sender_entry = tk.Entry(header_frame, textvariable=self.sender_email_var, font=("Segoe UI", 10), relief="solid", bd=1)
        self.sender_entry.pack(fill="x", pady=(2, 10))

        tk.Label(header_frame, text="Subject of Draft Email:", font=("Segoe UI", 10, "bold"), bg="#f0f0f0").pack(fill="x")
        self.subject_entry = tk.Entry(header_frame, textvariable=self.subject_var, font=("Segoe UI", 10), relief="solid", bd=1)
        self.subject_entry.pack(fill="x", pady=(2, 0))

        # --- Email Lists Section ---
        lists_frame = tk.Frame(main_frame, bg="#f0f0f0")
        lists_frame.pack(fill="both", expand=True)
        lists_frame.grid_columnconfigure(0, weight=1)
        lists_frame.grid_columnconfigure(1, weight=1)

        # 'To' List
        tk.Label(lists_frame, text="To:", font=("Segoe UI", 10, "bold"), bg="#f0f0f0").grid(row=0, column=0, sticky="w")
        self.to_text = scrolledtext.ScrolledText(lists_frame, height=15, width=35, font=("Segoe UI", 10), relief="solid", bd=1)
        self.to_text.grid(row=1, column=0, sticky="nsew", padx=(0, 5))

        # 'Cc' List
        tk.Label(lists_frame, text="Cc:", font=("Segoe UI", 10, "bold"), bg="#f0f0f0").grid(row=0, column=1, sticky="w", padx=(5, 0))
        self.cc_text = scrolledtext.ScrolledText(lists_frame, height=15, width=35, font=("Segoe UI", 10), relief="solid", bd=1)
        self.cc_text.grid(row=1, column=1, sticky="nsew", padx=(5, 0))

        # --- Buttons Section ---
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill="x", pady=(15, 0))

        self.send_button = tk.Button(button_frame, text="Send in Batch", command=self.send_emails, font=("Segoe UI", 10, "bold"), bg="#0078d4", fg="white", relief="flat", padx=10, pady=5)
        self.send_button.pack(side="right", padx=(5, 0))

        self.clear_button = tk.Button(button_frame, text="Clear", command=self.clear_fields, font=("Segoe UI", 10), bg="#e0e0e0", relief="flat", padx=10, pady=5)
        self.clear_button.pack(side="right", padx=(5, 0))
        
        self.review_button = tk.Button(button_frame, text="Review List", command=self.review_list, font=("Segoe UI", 10), bg="#e0e0e0", relief="flat", padx=10, pady=5)
        self.review_button.pack(side="right")

        # --- Load and Save ---
        self.load_state()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.log_message("Application started.")

    def log_message(self, message, level='info'):
        escaped_message = json.dumps(message)
        if level == 'info':
            logger.info(escaped_message)
        elif level == 'warning':
            logger.warning(escaped_message)
        elif level == 'error':
            logger.error(escaped_message)

    def get_emails_from_text(self, text_widget):
        text_content = text_widget.get("1.0", tk.END)
        # Regex to find email addresses
        email_regex = r'[\w\.-]+@[\w\.-]+'
        emails = re.findall(email_regex, text_content)
        # Clean up and remove duplicates
        return sorted(list(set(email.strip() for email in emails if email.strip())))

    def clear_fields(self):
        if messagebox.askokcancel("Confirm Clear", "Are you sure you want to clear all fields?"):
            self.subject_var.set("")
            self.to_text.delete("1.0", tk.END)
            self.cc_text.delete("1.0", tk.END)
            # Do not clear the sender email, as it's often reused
            self.log_message("Input fields cleared by user.")

    def review_list(self):
        to_emails = self.get_emails_from_text(self.to_text)
        cc_emails = self.get_emails_from_text(self.cc_text)

        review_message = f"Valid 'To' Emails ({len(to_emails)}):\n" + "\n".join(to_emails)
        review_message += f"\n\nValid 'Cc' Emails ({len(cc_emails)}):\n" + "\n".join(cc_emails)
        
        messagebox.showinfo("Review Email Lists", review_message)

    def send_emails(self):
        subject = self.subject_var.get().strip()
        if not subject:
            messagebox.showerror("Error", "Subject of the draft email cannot be empty.")
            return

        to_emails = self.get_emails_from_text(self.to_text)
        cc_emails = self.get_emails_from_text(self.cc_text)

        if not to_emails:
            messagebox.showerror("Error", "The 'To' list is empty. Please provide at least one recipient.")
            return

        # --- CC Confirmation ---
        default_cc = self.default_cc_var.get()
        if default_cc and default_cc not in cc_emails:
            if messagebox.askyesno("Confirm CC", f"The default CC '{default_cc}' is missing. Do you want to add it?"):
                cc_emails.append(default_cc)
                self.cc_text.insert(tk.END, f"\n{default_cc}")
        
        cc_string = "; ".join(cc_emails)
        
        confirm_msg = (f"You are about to send an email with the subject:\n'{subject}'\n\n"
                       f"To: {len(to_emails)} recipients\n"
                       f"Cc: {cc_string or 'None'}\n\n"
                       "Do you want to proceed?")

        if not messagebox.askyesno("Confirm Sending", confirm_msg):
            self.log_message("Sending process cancelled by user.")
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # --- Get Sending Account ---
            sender_email_str = self.sender_email_var.get().strip()
            sending_account = None
            if sender_email_str:
                for acc in namespace.Accounts:
                    if acc.SmtpAddress.lower() == sender_email_str.lower():
                        sending_account = acc
                        break
                if not sending_account:
                    messagebox.showwarning("Warning", f"Sender email '{sender_email_str}' not found in Outlook. The default account will be used instead.")
                    self.log_message(f"Warning: Sender email '{sender_email_str}' not found. Using default account.", 'warning')
            
            # --- Find Draft Email ---
            drafts_folder = namespace.GetDefaultFolder(16) # 16 is the constant for the Drafts folder
            draft_item = None
            for item in drafts_folder.Items:
                if item.Subject == subject:
                    draft_item = item
                    # This is a robust way to handle both opened (inline) and closed drafts
                    if hasattr(item, 'Display'): item.Display()
                    if hasattr(item, 'Close'): item.Close(0) # olDiscard
                    draft_item = outlook.ActiveInspector().CurrentItem
                    break
            
            if not draft_item:
                messagebox.showerror("Error", f"Could not find a draft with the subject: '{subject}'.\nPlease make sure it's saved in your Drafts folder.")
                self.log_message(f"Error: Draft not found with subject '{subject}'.", 'error')
                return

            self.log_message(f"Starting batch send. Subject: '{subject}'. To: {len(to_emails)}. Cc: {len(cc_emails)}.")

            # --- Loop and Send ---
            for recipient in to_emails:
                new_mail = draft_item.Copy()
                new_mail.To = recipient
                if cc_string:
                    new_mail.CC = cc_string
                
                # Set the sending account if one was found
                if sending_account:
                    new_mail.SendUsingAccount = sending_account
                
                new_mail.Send()
                self.log_message(f"Email sent to: {recipient}")

            messagebox.showinfo("Success", f"All {len(to_emails)} emails have been sent successfully!")
            self.log_message("Batch send completed successfully.")

        except Exception as e:
            messagebox.showerror("An Error Occurred", f"An unexpected error occurred:\n{e}")
            self.log_message(f"Fatal Error during sending: {e}", 'error')

    def save_state(self):
        state = {
            "sender_email": self.sender_email_var.get(),
            "subject": self.subject_var.get(),
            "to_list": self.to_text.get("1.0", tk.END),
            "cc_list": self.cc_text.get("1.0", tk.END)
        }
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        with open(CONFIG_FILE, 'w') as f:
            json.dump(state, f)
        self.log_message("Application state saved.")

    def load_state(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                try:
                    state = json.load(f)
                    self.sender_email_var.set(state.get("sender_email", ""))
                    self.subject_var.set(state.get("subject", ""))
                    self.to_text.insert("1.0", state.get("to_list", ""))
                    self.cc_text.insert("1.0", state.get("cc_list", ""))
                    self.log_message("Application state loaded.")
                except json.JSONDecodeError:
                    self.log_message("Warning: Could not decode JSON from config file.", 'warning')

    def on_closing(self):
        self.save_state()
        self.destroy()

if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()