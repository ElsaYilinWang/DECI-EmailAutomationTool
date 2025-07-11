import tkinter as tk
from tkinter import messagebox, Text, font, simpledialog
import json
import win32com.client as win32
import threading
import os
import re
import logging
from logging.handlers import RotatingFileHandler

# --- Setup System Logging ---
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_file = 'app_log.log'
# Use a rotating file handler to keep log files from growing too large
# This will create up to 3 backup log files, each up to 1MB in size.
my_handler = RotatingFileHandler(log_file, mode='a', maxBytes=1*1024*1024, 
                                 backupCount=3, encoding=None, delay=0)
my_handler.setFormatter(log_formatter)
my_handler.setLevel(logging.INFO)

app_log = logging.getLogger('root')
app_log.setLevel(logging.INFO)
app_log.addHandler(my_handler)


class EmailApp:
    """
    A desktop application for sending batch emails via Microsoft Outlook,
    using a pre-formatted draft as a template.
    """
    def __init__(self, root):
        app_log.info("Application starting up.")
        self.root = root
        self.root.title("Email Automation Tool")
        self.root.geometry("900x650")

        # --- UI Styling ---
        self.colors = {
            'bg': '#F5F5F5', 'frame_bg': '#FFFFFF', 'text': '#242424',
            'secondary_text': '#605E5C', 'button_bg': '#0078D4', 'button_fg': '#FFFFFF',
            'button_hover': '#106EBE', 'button_secondary_bg': '#FFFFFF',
            'button_secondary_fg': '#242424', 'button_secondary_hover': '#F0F0F0',
            'entry_bg': '#FFFFFF', 'border': '#C8C6C4', 'border_focus': '#0078D4'
        }
        self.root.configure(bg=self.colors['bg'])
        self.font_normal = font.Font(family="Segoe UI", size=10)
        self.font_bold = font.Font(family="Segoe UI", size=11, weight="bold")
        self.font_title = font.Font(family="Segoe UI Semibold", size=12)

        # --- Data Storage ---
        self.data_file = "email_data.json"
        self.to_emails_str = ""
        self.cc_emails_str = ""
        self.draft_subject = ""
        self.sender_email = "Detecting..."
        self.load_data()

        # --- UI Setup ---
        self.create_widgets()
        self.populate_fields()
        
        self.cancel_sending = False
        threading.Thread(target=self.get_sender_email, daemon=True).start()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Handle the window closing event."""
        app_log.info("Application shutting down.")
        self.root.destroy()

    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=0)
        main_frame.grid_columnconfigure(0, weight=2) # Give To field more space
        main_frame.grid_columnconfigure(1, weight=1)

        # --- "To" Emails Frame ---
        to_frame_container = tk.LabelFrame(main_frame, text="To (Paste Email List)", 
                                           bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                           font=self.font_title, relief='flat', borderwidth=0)
        to_frame_container.grid(row=0, column=0, padx=(0, 10), pady=5, sticky="nsew")
        to_frame_container.grid_rowconfigure(0, weight=1)
        to_frame_container.grid_columnconfigure(0, weight=1)

        self.to_emails_text = Text(to_frame_container, bg=self.colors['entry_bg'], fg=self.colors['text'],
                                   relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                                   borderwidth=1, highlightthickness=1, wrap='word', padx=5, pady=5)
        self.to_emails_text.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
        self.to_emails_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=(5,0))

        review_to_button = self.create_modern_button(to_frame_container, "Review 'To' List", lambda: self.review_emails(self.to_emails_text, "'To' Recipients"), 'secondary')
        review_to_button.grid(row=1, column=0, sticky="e", padx=10, pady=10)
        
        # --- "CC" Emails Frame ---
        cc_frame = tk.LabelFrame(main_frame, text="Cc (Paste Email List)", 
                                 bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                 font=self.font_title, relief='flat', borderwidth=0)
        cc_frame.grid(row=0, column=1, padx=(10, 0), pady=5, sticky="nsew")
        cc_frame.grid_rowconfigure(0, weight=1)
        cc_frame.grid_columnconfigure(0, weight=1)

        self.cc_emails_text = Text(cc_frame, bg=self.colors['entry_bg'], fg=self.colors['text'],
                                   relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                                   borderwidth=1, highlightthickness=1, wrap='word', padx=5, pady=5)
        self.cc_emails_text.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
        self.cc_emails_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=(5,0))
        
        review_cc_button = self.create_modern_button(cc_frame, "Review 'Cc' List", lambda: self.review_emails(self.cc_emails_text, "'Cc' Recipients"), 'secondary')
        review_cc_button.grid(row=1, column=0, sticky="e", padx=10, pady=10)

        # --- Draft Subject Frame ---
        draft_subject_frame = tk.LabelFrame(main_frame, text="Subject of Draft Template in Outlook", 
                                   bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                   font=self.font_title, relief='flat', borderwidth=0)
        draft_subject_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="nsew")
        draft_subject_frame.grid_columnconfigure(0, weight=1)

        self.draft_subject_entry = tk.Entry(draft_subject_frame, bg=self.colors['entry_bg'], fg=self.colors['text'],
                             relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                             borderwidth=1, highlightthickness=1)
        self.draft_subject_entry.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
        self.draft_subject_entry.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # --- Status and Buttons Frame ---
        bottom_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        bottom_frame.grid(row=3, column=0, columnspan=2, pady=(15,0), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)

        self.sender_label = tk.Label(bottom_frame, text=f"Sending from: {self.sender_email}", 
                                     font=self.font_normal, bg=self.colors['bg'], 
                                     fg=self.colors['secondary_text'], anchor='w')
        self.sender_label.grid(row=0, column=0, sticky='w')

        button_container = tk.Frame(bottom_frame, bg=self.colors['bg'])
        button_container.grid(row=0, column=1, sticky='e')

        self.create_modern_button(button_container, "Clear", self.clear_fields, 'secondary').pack(side='left', padx=(0,10))
        self.create_modern_button(button_container, "Cancel", self.cancel_send, 'secondary').pack(side='left', padx=(0,10))
        self.create_modern_button(button_container, "Send in Batch", self.start_sending_thread, 'primary').pack(side='left')

    def _get_validated_emails(self, text_content):
        email_regex = r'[\w\.\-]+@[\w\.\-]+'
        potential_emails = re.findall(email_regex, text_content)
        return sorted(list(set(email.lower() for email in potential_emails)))

    def review_emails(self, text_widget, title):
        text_content = text_widget.get("1.0", "end")
        valid_emails = self._get_validated_emails(text_content)
        
        review_window = tk.Toplevel(self.root)
        review_window.title(f"{title} Review")
        review_window.geometry("400x500")
        review_window.configure(bg=self.colors['frame_bg'])

        tk.Label(review_window, text=f"Found {len(valid_emails)} valid & unique emails:", 
                 font=self.font_bold, bg=self.colors['frame_bg'], fg=self.colors['text']).pack(pady=10)

        list_frame = tk.Frame(review_window)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=self.font_normal, 
                             bg=self.colors['entry_bg'], fg=self.colors['text'], relief='flat')
        scrollbar.config(command=listbox.yview)
        
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)

        for email in valid_emails:
            listbox.insert("end", email)

        tk.Button(review_window, text="Close", command=review_window.destroy, 
                  font=self.font_bold, bg=self.colors['button_bg'], fg=self.colors['button_fg'], 
                  relief='flat', padx=10).pack(pady=10)
        review_window.transient(self.root)
        review_window.grab_set()

    def create_modern_button(self, parent, text, command, style='primary'):
        if style == 'primary':
            bg_color, fg_color, hover_color = self.colors['button_bg'], self.colors['button_fg'], self.colors['button_hover']
        else:
            bg_color, fg_color, hover_color = self.colors['button_secondary_bg'], self.colors['button_secondary_fg'], self.colors['button_secondary_hover']

        button = tk.Button(parent, text=text, command=command, bg=bg_color, fg=fg_color, font=self.font_bold, 
                           relief='flat', borderwidth=1, activebackground=hover_color, activeforeground=fg_color, pady=8, padx=15)
        button.bind("<Enter>", lambda e, h=hover_color: e.widget.config(bg=h))
        button.bind("<Leave>", lambda e, b=bg_color: e.widget.config(bg=b))
        return button

    def get_sender_email(self):
        try:
            outlook = win32.Dispatch('outlook.application')
            self.sender_email = outlook.Session.Accounts[0].SmtpAddress
            app_log.info(f"Successfully detected sender email: {self.sender_email}")
        except Exception as e:
            self.sender_email = "Outlook not running or no account found."
            app_log.error(f"Failed to detect sender email. Error: {e}")
        
        self.sender_label.config(text=f"Sending from: {self.sender_email}")

    def load_data(self):
        if os.path.exists(self.data_file):
            with open(self.data_file, 'r') as f:
                try:
                    data = json.load(f)
                    self.to_emails_str = data.get("to_emails_str", "")
                    self.cc_emails_str = data.get("cc_emails_str", "")
                    self.draft_subject = data.get("draft_subject", "")
                    app_log.info("User data loaded successfully.")
                except json.JSONDecodeError:
                    app_log.error("Failed to decode JSON data file.")
                    pass

    def save_data(self):
        data = {
            "to_emails_str": self.to_emails_text.get("1.0", "end-1c"),
            "cc_emails_str": self.cc_emails_text.get("1.0", "end-1c"),
            "draft_subject": self.draft_subject_entry.get()
        }
        with open(self.data_file, 'w') as f:
            json.dump(data, f, indent=4)
        app_log.info("User data saved.")

    def populate_fields(self):
        self.to_emails_text.insert("1.0", self.to_emails_str)
        
        if not self.cc_emails_str.strip():
            self.cc_emails_text.insert("1.0", 'mro@deci-ltd.com')
        else:
            self.cc_emails_text.insert("1.0", self.cc_emails_str)
        
        self.draft_subject_entry.insert(0, self.draft_subject)

    def start_sending_thread(self):
        app_log.info("'Send in Batch' button clicked.")
        self.save_data()
        self.cancel_sending = False
        threading.Thread(target=self.send_emails).start()

    def send_emails(self):
        cc_text_content = self.cc_emails_text.get("1.0", "end")
        cc_list = self._get_validated_emails(cc_text_content)
        cc_list_str = "; ".join(cc_list)
        
        if not messagebox.askokcancel("Confirm CC Addresses", f"You are about to send this email with the following in CC:\n\n{cc_list_str if cc_list_str else 'NOBODY'}\n\nDo you want to proceed?"):
            app_log.warning("User cancelled sending at CC confirmation.")
            return

        to_list = self._get_validated_emails(self.to_emails_text.get("1.0", "end"))
        template_subject = self.draft_subject_entry.get()
        
        app_log.info(f"Starting batch send. Template Subject: '{template_subject}'. Number of recipients: {len(to_list)}. CC list: '{cc_list_str}'")

        if not to_list:
            messagebox.showwarning("Input Error", "No valid recipient emails found.")
            app_log.warning("Sending aborted: No valid 'To' recipients.")
            return
        if not template_subject:
            messagebox.showwarning("Input Error", "Please enter the subject of the draft template.")
            app_log.warning("Sending aborted: No draft subject provided.")
            return

        try:
            outlook = win32.Dispatch('outlook.application')
            namespace = outlook.GetNamespace("MAPI")
            drafts_folder = namespace.GetDefaultFolder(16)
            template_email = next((item for item in drafts_folder.Items if item.Subject == template_subject), None)
            
            if template_email is None:
                messagebox.showerror("Template Not Found", f"Could not find a draft with the subject: '{template_subject}'")
                app_log.error(f"Draft template not found with subject: '{template_subject}'")
                return

            template_body = template_email.HTMLBody
            app_log.info("Successfully found and loaded draft template.")
            
            for i, recipient in enumerate(to_list):
                if self.cancel_sending:
                    messagebox.showinfo("Cancelled", "Email sending has been cancelled.")
                    app_log.warning(f"User cancelled sending after {i} emails were sent.")
                    break
                
                new_mail = outlook.CreateItem(0)
                new_mail.To = recipient
                new_mail.CC = cc_list_str
                new_mail.Subject = template_subject
                new_mail.HTMLBody = template_body
                new_mail.Send()
                app_log.info(f"Email sent to: {recipient}")
            else: 
                if not self.cancel_sending:
                    messagebox.showinfo("Success", "All emails have been sent successfully.")
                    app_log.info("Batch send process completed successfully.")
        except Exception as e:
            messagebox.showerror("Email Error", f"An error occurred during sending.\nError: {e}")
            app_log.error(f"An exception occurred during email sending: {e}")

    def cancel_send(self):
        app_log.info("'Cancel' button clicked.")
        self.cancel_sending = True

    def clear_fields(self):
        app_log.info("'Clear' button clicked.")
        if messagebox.askokcancel("Confirm Clear", "Are you sure you want to clear all fields?"):
            app_log.info("User confirmed clearing all fields.")
            self.to_emails_text.delete("1.0", 'end')
            self.cc_emails_text.delete("1.0", 'end')
            self.draft_subject_entry.delete(0, 'end')
            self.populate_fields()
            self.save_data()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = EmailApp(root)
        root.mainloop()
    except Exception as e:
        app_log.critical(f"A critical error occurred, causing the application to crash: {e}")
        raise
