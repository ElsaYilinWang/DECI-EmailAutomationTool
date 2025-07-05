import tkinter as tk
from tkinter import messagebox, Text, font
import json
import win32com.client as win32
import threading
import os

class EmailApp:
    """
    A desktop application for sending batch emails via Microsoft Outlook,
    styled to resemble the new Outlook interface.
    """
    def __init__(self, root):
        """
        Initializes the application's user interface and loads saved data.
        """
        self.root = root
        self.root.title("Email Automation Tool")
        self.root.geometry("900x780")

        # --- New Outlook UI Styling ---
        self.colors = {
            'bg': '#F5F5F5',
            'frame_bg': '#FFFFFF',
            'text': '#242424',
            'secondary_text': '#605E5C',
            'button_bg': '#0078D4',
            'button_fg': '#FFFFFF',
            'button_hover': '#106EBE',
            'button_secondary_bg': '#FFFFFF',
            'button_secondary_fg': '#242424',
            'button_secondary_hover': '#F0F0F0',
            'entry_bg': '#FFFFFF',
            'border': '#C8C6C4',
            'border_focus': '#0078D4'
        }
        self.root.configure(bg=self.colors['bg'])
        self.font_normal = font.Font(family="Segoe UI", size=10)
        self.font_bold = font.Font(family="Segoe UI", size=11, weight="bold")
        self.font_title = font.Font(family="Segoe UI Semibold", size=12)

        # --- Data Storage ---
        self.data_file = "email_data.json"
        self.to_emails = []
        self.cc_emails = []
        self.subject = ""
        self.sender_email = "Detecting..." # Default value
        self.load_data()

        # --- UI Setup ---
        self.create_widgets()
        self.populate_fields()
        
        # --- Cancellation Flag ---
        self.cancel_sending = False
        
        # --- Get Sender Email ---
        # Run in a separate thread to not freeze UI on startup
        threading.Thread(target=self.get_sender_email, daemon=True).start()


    def create_widgets(self):
        """
        Creates and arranges all the UI components in the main window.
        """
        # --- Main Frames ---
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        main_frame.grid_rowconfigure(0, weight=1) 
        main_frame.grid_rowconfigure(1, weight=0) # Subject row
        main_frame.grid_rowconfigure(2, weight=5) # Message row
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # --- "To" Emails Frame (with Scrollbar) ---
        to_frame_container = tk.LabelFrame(main_frame, text="To", 
                                           bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                           font=self.font_title, relief='flat', borderwidth=0)
        to_frame_container.grid(row=0, column=0, padx=(0, 10), pady=5, sticky="nsew")
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
                             relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                             borderwidth=1, highlightthickness=1)
            entry.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
            entry.grid(row=i, column=0, padx=(10,15), pady=3, sticky="ew")
            self.to_entries.append(entry)

        # --- "CC" Emails Frame ---
        cc_frame = tk.LabelFrame(main_frame, text="Cc", 
                                 bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                 font=self.font_title, relief='flat', borderwidth=0)
        cc_frame.grid(row=0, column=1, padx=(10, 0), pady=5, sticky="nsew")
        cc_frame.grid_columnconfigure(0, weight=1)

        self.cc_entries = []
        for i in range(6):
            entry = tk.Entry(cc_frame, width=30, bg=self.colors['entry_bg'], fg=self.colors['text'],
                             relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                             borderwidth=1, highlightthickness=1)
            entry.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
            entry.grid(row=i, column=0, padx=10, pady=3, sticky="ew")
            self.cc_entries.append(entry)

        # --- Email Subject Frame ---
        subject_frame = tk.LabelFrame(main_frame, text="Subject", 
                                   bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                   font=self.font_title, relief='flat', borderwidth=0)
        subject_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky="nsew")
        subject_frame.grid_columnconfigure(0, weight=1)

        self.subject_entry = tk.Entry(subject_frame, bg=self.colors['entry_bg'], fg=self.colors['text'],
                             relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                             borderwidth=1, highlightthickness=1)
        self.subject_entry.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
        self.subject_entry.grid(row=0, column=0, padx=10, pady=5, sticky="ew")


        # --- Email Body Frame ---
        body_frame = tk.LabelFrame(main_frame, text="Message", 
                                   bg=self.colors['frame_bg'], fg=self.colors['secondary_text'], 
                                   font=self.font_title, relief='flat', borderwidth=0)
        body_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="nsew")
        body_frame.grid_rowconfigure(0, weight=1)
        body_frame.grid_columnconfigure(0, weight=1)

        self.email_body = Text(body_frame, width=80, bg=self.colors['entry_bg'], fg=self.colors['text'],
                               relief='solid', font=self.font_normal, insertbackground=self.colors['text'],
                               borderwidth=1, highlightthickness=1, wrap='word', padx=5, pady=5)
        self.email_body.config(highlightbackground=self.colors['border'], highlightcolor=self.colors['border_focus'])
        self.email_body.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

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

    def create_modern_button(self, parent, text, command, style='primary'):
        """Helper function to create styled buttons with hover effects."""
        if style == 'primary':
            bg_color = self.colors['button_bg']
            fg_color = self.colors['button_fg']
            hover_color = self.colors['button_hover']
        else: # secondary
            bg_color = self.colors['button_secondary_bg']
            fg_color = self.colors['button_secondary_fg']
            hover_color = self.colors['button_secondary_hover']

        button = tk.Button(parent, text=text, command=command,
                           bg=bg_color, fg=fg_color,
                           font=self.font_bold, relief='flat', borderwidth=1,
                           activebackground=hover_color,
                           activeforeground=fg_color,
                           pady=8, padx=15)
        
        button.bind("<Enter>", lambda e, h=hover_color: e.widget.config(bg=h))
        button.bind("<Leave>", lambda e, b=bg_color: e.widget.config(bg=b))
        return button

    def get_sender_email(self):
        """Detects the default Outlook account email and updates the UI."""
        try:
            outlook = win32.Dispatch('outlook.application')
            # This is a reliable way to get the primary account's SMTP address
            self.sender_email = outlook.Session.Accounts[0].SmtpAddress
        except Exception:
            self.sender_email = "Outlook not running or no account found."
        
        # Update the label in the UI
        self.sender_label.config(text=f"Sending from: {self.sender_email}")

    def load_data(self):
        """
        Loads email addresses and subject from the JSON data file.
        """
        if os.path.exists(self.data_file):
            with open(self.data_file, 'r') as f:
                try:
                    data = json.load(f)
                    self.to_emails = data.get("to_emails", [])
                    self.cc_emails = data.get("cc_emails", [])
                    self.subject = data.get("subject", "")
                except json.JSONDecodeError:
                    self.to_emails, self.cc_emails, self.subject = [], [], ""

    def save_data(self):
        """
        Saves the current email addresses and subject to the JSON data file.
        """
        data = {
            "to_emails": [entry.get() for entry in self.to_entries],
            "cc_emails": [entry.get() for entry in self.cc_entries],
            "subject": self.subject_entry.get()
        }
        with open(self.data_file, 'w') as f:
            json.dump(data, f, indent=4)

    def populate_fields(self):
        """
        Fills the entry fields with the loaded email addresses and subject.
        """
        for i, email in enumerate(self.to_emails):
            if i < len(self.to_entries):
                self.to_entries[i].insert(0, email)

        for i, email in enumerate(self.cc_emails):
            if i < len(self.cc_entries):
                self.cc_entries[i].insert(0, email)
        
        self.subject_entry.insert(0, self.subject)

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
        Iterates through the recipient list and sends emails via Outlook, including the default signature.
        """
        try:
            outlook = win32.Dispatch('outlook.application')
        except Exception as e:
            messagebox.showerror("Outlook Error", f"Microsoft Outlook is not running or could not be started.\nError: {e}")
            return

        to_list = [entry.get() for entry in self.to_entries if entry.get()]
        cc_list = ";".join([entry.get() for entry in self.cc_entries if entry.get()])
        subject_text = self.subject_entry.get()
        user_content = self.email_body.get("1.0", "end-1c")

        if not to_list:
            messagebox.showwarning("Input Error", "Please enter at least one recipient.")
            return

        for recipient in to_list:
            if self.cancel_sending:
                messagebox.showinfo("Cancelled", "Email sending has been cancelled.")
                break
            
            try:
                mail = outlook.CreateItem(0)
                
                mail.GetInspector 

                signature = mail.HTMLBody

                user_content_html = f"<p>{user_content.replace(os.linesep, '<br>')}</p>"

                mail.To = recipient
                mail.CC = cc_list
                mail.Subject = subject_text if subject_text else "No Subject"
                mail.HTMLBody = user_content_html + signature
                
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
            self.subject_entry.delete(0, 'end')
            self.email_body.delete("1.0", 'end')
            self.save_data()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
