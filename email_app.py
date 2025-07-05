import tkinter as tk
from tkinter import messagebox, Text
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
        self.root.geometry("800x600")

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
        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # --- "To" Emails Frame ---
        to_frame = tk.LabelFrame(main_frame, text="Receiver's Email Addresses (To)")
        to_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.to_entries = []
        for i in range(24):
            entry = tk.Entry(to_frame, width=30)
            entry.grid(row=i, column=0, padx=5, pady=2)
            self.to_entries.append(entry)

        # --- "CC" Emails Frame ---
        cc_frame = tk.LabelFrame(main_frame, text="CC Email Addresses")
        cc_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.cc_entries = []
        for i in range(6):
            entry = tk.Entry(cc_frame, width=30)
            entry.grid(row=i, column=0, padx=5, pady=2)
            self.cc_entries.append(entry)

        # --- Email Body Frame ---
        body_frame = tk.LabelFrame(main_frame, text="Email Body")
        body_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        self.email_body = Text(body_frame, height=15, width=80)
        self.email_body.pack(padx=5, pady=5, fill="both", expand=True)

        # --- Buttons Frame ---
        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        send_button = tk.Button(button_frame, text="Send in Batch", command=self.start_sending_thread)
        send_button.grid(row=0, column=0, padx=10)

        cancel_button = tk.Button(button_frame, text="Cancel", command=self.cancel_send)
        cancel_button.grid(row=0, column=1, padx=10)

        clear_button = tk.Button(button_frame, text="Clear", command=self.clear_fields)
        clear_button.grid(row=0, column=2, padx=10)

    def load_data(self):
        """
        Loads email addresses from the JSON data file.
        """
        if os.path.exists(self.data_file):
            with open(self.data_file, 'r') as f:
                data = json.load(f)
                self.to_emails = data.get("to_emails", [])
                self.cc_emails = data.get("cc_emails", [])

    def save_data(self):
        """
        Saves the current email addresses to the JSON data file.
        """
        data = {
            "to_emails": [entry.get() for entry in self.to_entries],
            "cc_emails": [entry.get() for entry in self.cc_entries]
        }
        with open(self.data_file, 'w') as f:
            json.dump(data, f)

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
        
        # Create and start the email sending thread
        send_thread = threading.Thread(target=self.send_emails)
        send_thread.start()

    def send_emails(self):
        """
        Iterates through the recipient list and sends emails via Outlook.
        """
        try:
            outlook = win32.Dispatch('outlook.application')
        except Exception as e:
            messagebox.showerror("Outlook Error", "Microsoft Outlook is not running or could not be started.")
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
                mail.Subject = "Your Subject Here" # You can customize the subject
                mail.Body = email_content
                mail.Send()
            except Exception as e:
                messagebox.showerror("Email Error", f"Could not send email to {recipient}.\nError: {e}")
        else: # This 'else' belongs to the 'for' loop and executes if the loop completes without a 'break'
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
        for entry in self.to_entries:
            entry.delete(0, 'end')
        for entry in self.cc_entries:
            entry.delete(0, 'end')
        self.email_body.delete("1.0", 'end')
        self.save_data() # Save the cleared fields

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
