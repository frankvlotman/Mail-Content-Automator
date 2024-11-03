import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import win32com.client as win32
from PIL import Image
import os
import pandas as pd
import io
import win32clipboard  # For clipboard operations
import html  # For escaping HTML characters
from tkhtmlview import HTMLScrolledText  # Importing HTMLScrolledText for HTML rendering with scrollbars
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter.font as tkFont  # Import the font module
import logging  # For debugging

# ------------------- Configure Logging -------------------
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# ------------------- Icon Handling -------------------

# Define the path for the blank icon
# Update this path to point to your actual icon file if desired
icon_path = 'Your\\path\\blank.ico'  # Example path

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    if not os.path.exists(path):
        size = (16, 16)  # Size of the icon
        image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
        image.save(path, format="ICO")
        logging.debug(f"Created blank icon at {path}")
    else:
        logging.debug(f"Icon already exists at {path}")

# Create the blank ICO file
create_blank_ico(icon_path)

# ------------------- Initialize Root Window -------------------

# Initialize the root window with drag-and-drop support
root = TkinterDnD.Tk()
root.title("Mail Content Automator")
try:
    root.iconbitmap(icon_path)  # Set the icon for the main window
    logging.debug(f"Set icon from {icon_path}")
except Exception as e:
    logging.warning(f"Unable to set icon. {e}")

# <--- Change: Adjusted window geometry to reduce overall height
root.geometry("1000x740")

# Initialize ttk.Style
style = ttk.Style()
style.theme_use("clam")  # Use 'clam' theme for better customization

# Define custom style for buttons
style.configure("Custom.TButton",
                background="#d0e8f1",
                foreground="black",
                borderwidth=1,
                focusthickness=3,
                focuscolor='none')

# Define style map for hover (active) state
style.map("Custom.TButton",
          background=[('active', '#87CEFA')],
          foreground=[('active', 'black')])

# ------------------- Initialize Selected Recipients -------------------

selected_recipients = {'to': [], 'cc': []}

# ------------------- Create Notebook and Tabs -------------------

# Create a Notebook (tabbed interface)
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Define the first tab (Copy Paste Excel Emailer)
tab_emailer = ttk.Frame(notebook)
notebook.add(tab_emailer, text="Copy Paste Excel Emailer")

# ------------------- Tab Content: Copy Paste Excel Emailer -------------------

def focus_next_widget(event):
    """Move focus to the next widget in the tab order."""
    event.widget.tk_focusNext().focus()
    return "break"

# ------------------- Subject Input -------------------

def add_subject_input(frame):
    """Add an input field for the email subject."""
    subject_label = tk.Label(frame, text="Email Subject:")
    subject_label.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
    subject_entry = tk.Entry(frame, width=50)
    subject_entry.grid(row=0, column=1, padx=(0, 10), pady=5)
    subject_entry.bind("<Return>", focus_next_widget)  # Bind 'Enter' key to move focus
    return subject_entry

# ------------------- Greeting and Email Body -------------------

def add_greeting_email_body_inputs(frame):
    """Add input fields for greeting and email body."""
    greeting_label = tk.Label(frame, text="Greeting")
    greeting_label.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
    greeting_entry = tk.Entry(frame, width=50)
    greeting_entry.grid(row=1, column=1, padx=(0, 10), pady=5)
    greeting_entry.bind("<Return>", focus_next_widget)  # Bind 'Enter' key to move focus

    email_body_label = tk.Label(frame, text="Email Body:")
    email_body_label.grid(row=2, column=0, sticky="nw", padx=(0, 10), pady=5)

    # Define the Aptos (Body) font with size 11
    try:
        aptos_font = tkFont.Font(family="Aptos (Body)", size=11)
        logging.debug("Aptos (Body) font found and set.")
    except:
        # Fallback to a default font if Aptos is not available
        aptos_font = tkFont.Font(family="Arial", size=11)
        logging.warning("Aptos (Body) font not found. Falling back to Arial.")

    email_body_text = tk.Text(frame, width=50, height=10, wrap='word', font=aptos_font)
    email_body_text.grid(row=2, column=1, padx=(0, 10), pady=5)
    email_body_text.insert(tk.END, "Please see below for details about a delivery expected to arrive at the company.")
    # Removed the Enter key binding to allow normal line breaks

    return greeting_entry, email_body_text

# ------------------- Recipient Selection Button -------------------

def add_recipient_selection_button(frame):
    """Add a button to open the recipient selection window."""
    def open_recipient_window():
        RecipientWindow(frame, selected_recipients)  # Pass 'frame' as parent

    recipient_button = ttk.Button(
        frame,
        text="Add Recipients",
        command=open_recipient_window,
        style="Custom.TButton"
    )
    recipient_button.grid(row=3, column=0, padx=5, pady=5, sticky="w")

    # Label to display selected recipients summary
    # <--- Change: Adjust wraplength and justification for better display
    summary_label = tk.Label(frame, text="No recipients selected.", fg="grey",
                             wraplength=800, justify="left")
    summary_label.grid(row=4, column=0, columnspan=2, sticky="we", padx=5, pady=(0, 10))

    # Function to update the summary label
    def update_summary():
        to = "; ".join(selected_recipients['to'])
        cc = "; ".join(selected_recipients['cc'])
        summary_text = ""
        if to:
            summary_text += f"To: {to}\n"
        if cc:
            summary_text += f"CC: {cc}"
        if summary_text:
            summary_label.config(text=summary_text, fg="black")
        else:
            summary_label.config(text="No recipients selected.", fg="grey")

    # Store the update function for later use
    frame.update_summary = update_summary

    return summary_label

# ------------------- Recipient Selection Window Class -------------------

class RecipientWindow(tk.Toplevel):
    def __init__(self, parent, selected_recipients):
        super().__init__(parent)
        self.title("Select Recipients")
        self.geometry("600x400")
        self.resizable(False, False)
        self.selected_recipients = selected_recipients  # Dictionary to store selections

        # Define the list of email addresses
        self.email_options = [
            'user1@example.com',
            'user2@example.com',
            'user3@example.com',
            'user4@example.com',
            'user5@example.com',
            'user6@example.com',
            'user7@example.com',
            'user8@example.com',
            'user9@example.com',
            'user10@example.com'        
        ]

        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        # Frames for To and CC
        to_frame = tk.LabelFrame(self, text="To", padx=10, pady=10)
        to_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        cc_frame = tk.LabelFrame(self, text="CC", padx=10, pady=10)
        cc_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        # Listboxes for To and CC
        self.to_listbox = tk.Listbox(to_frame, selectmode=tk.MULTIPLE, width=30, exportselection=False)
        self.to_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for email in self.email_options:
            self.to_listbox.insert(tk.END, email)

        self.to_scrollbar = ttk.Scrollbar(to_frame, orient="vertical", command=self.to_listbox.yview)
        self.to_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.to_listbox.configure(yscrollcommand=self.to_scrollbar.set)

        self.cc_listbox = tk.Listbox(cc_frame, selectmode=tk.MULTIPLE, width=30, exportselection=False)
        self.cc_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for email in self.email_options:
            self.cc_listbox.insert(tk.END, email)

        self.cc_scrollbar = ttk.Scrollbar(cc_frame, orient="vertical", command=self.cc_listbox.yview)
        self.cc_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.cc_listbox.configure(yscrollcommand=self.cc_scrollbar.set)

        # Pre-select previously selected recipients
        for i, email in enumerate(self.email_options):
            if email in self.selected_recipients['to']:
                self.to_listbox.selection_set(i)
            if email in self.selected_recipients['cc']:
                self.cc_listbox.selection_set(i)

        # Buttons for OK and Cancel
        button_frame = tk.Frame(self)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)

        ok_button = ttk.Button(button_frame, text="OK", command=self.on_ok, style="Custom.TButton")
        ok_button.pack(side=tk.LEFT, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.destroy, style="Custom.TButton")
        cancel_button.pack(side=tk.LEFT, padx=5)

    def on_ok(self):
        # Get selected To recipients
        to_indices = self.to_listbox.curselection()
        to_selected = [self.email_options[i] for i in to_indices]

        # Get selected CC recipients
        cc_indices = self.cc_listbox.curselection()
        cc_selected = [self.email_options[i] for i in cc_indices]

        # Update the selected_recipients dictionary
        self.selected_recipients['to'] = to_selected
        self.selected_recipients['cc'] = cc_selected

        # Update the summary in the main window
        if hasattr(self.master, 'update_summary'):
            self.master.update_summary()

        self.destroy()

# ------------------- Attachment Handling -------------------

def add_attachment_section(frame):
    """Add a section for handling attachments, including add/remove buttons and drag-and-drop."""
    attachment_label = tk.Label(frame, text="Attachments:")
    attachment_label.grid(row=5, column=0, sticky="nw", padx=(0, 10), pady=5)

    # Frame for attachment list and buttons
    attachment_frame = tk.Frame(frame, relief=tk.SUNKEN, borderwidth=1)
    attachment_frame.grid(row=5, column=1, columnspan=3, padx=(0, 10), pady=5, sticky="ew")  # Changed sticky to 'ew'

    # Configure grid weights for resizing
    frame.grid_rowconfigure(5, weight=0)  # Changed from 1 to 0 to prevent vertical expansion
    frame.grid_columnconfigure(1, weight=1)

    # Listbox to display attachments
    attachment_listbox = tk.Listbox(attachment_frame, selectmode=tk.MULTIPLE, width=80, height=2)  # Set height=2  # <--- Change
    attachment_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5,0), pady=5)  # Changed fill to 'X'

    # Scrollbar for the attachment listbox
    attachment_scrollbar = ttk.Scrollbar(attachment_frame, orient="vertical", command=attachment_listbox.yview)
    attachment_scrollbar.pack(side=tk.RIGHT, fill="y")
    attachment_listbox.configure(yscrollcommand=attachment_scrollbar.set)

    # Frame for buttons
    button_frame = tk.Frame(frame)
    button_frame.grid(row=6, column=1, columnspan=3, pady=(0,10), sticky="w", padx=(0,10))

    # Add Attachment Button
    add_attachment_button = ttk.Button(button_frame, text="Add Attachment", command=lambda: add_attachments(attachment_listbox), style="Custom.TButton")
    add_attachment_button.pack(side=tk.LEFT, padx=5)

    # Remove Attachment Button
    remove_attachment_button = ttk.Button(button_frame, text="Remove Selected", command=lambda: remove_attachments(attachment_listbox), style="Custom.TButton")
    remove_attachment_button.pack(side=tk.LEFT, padx=5)

    # Drag-and-Drop Area Label
    dnd_label = tk.Label(attachment_frame, text="Drag and drop files here to attach", fg="grey")
    dnd_label.place(relx=0.5, rely=0.5, anchor="center")

    # Bind drag-and-drop events
    attachment_listbox.drop_target_register(DND_FILES)
    attachment_listbox.dnd_bind('<<Drop>>', lambda event: drop_files(event, attachment_listbox))

    # Change label appearance on drag enter and leave
    attachment_listbox.dnd_bind('<<DragEnter>>', lambda event: on_drag_enter(event, dnd_label))
    attachment_listbox.dnd_bind('<<DragLeave>>', lambda event: on_drag_leave(event, dnd_label))

    return attachment_listbox

def add_attachments(listbox):
    """Open a file dialog to select files and add them to the listbox."""
    files = filedialog.askopenfilenames(title="Select files to attach")
    for file in files:
        if file not in listbox.get(0, tk.END):
            listbox.insert(tk.END, file)
            logging.debug(f"Added attachment: {file}")

def remove_attachments(listbox):
    """Remove selected attachments from the listbox."""
    selected_indices = listbox.curselection()
    for index in reversed(selected_indices):
        file = listbox.get(index)
        listbox.delete(index)
        logging.debug(f"Removed attachment: {file}")

def drop_files(event, listbox):
    """Handle files dropped into the listbox."""
    files = root.splitlist(event.data)
    for file in files:
        if os.path.isfile(file):
            if file not in listbox.get(0, tk.END):
                listbox.insert(tk.END, file)
                logging.debug(f"Drag-and-Dropped attachment: {file}")

def on_drag_enter(event, label):
    """Provide visual feedback when dragging enters the attachment area."""
    label.config(fg="black")
    label.config(font=("Arial", 10, "bold"))

def on_drag_leave(event, label):
    """Revert visual feedback when dragging leaves the attachment area."""
    label.config(fg="grey")
    label.config(font=("Arial", 10))

# ------------------- Buttons for Pasting, Clearing, Sending, and Previewing Email -------------------

def add_action_buttons(frame, attachment_listbox):
    """Add buttons for pasting data, clearing data, sending email, and previewing email."""
    paste_button = ttk.Button(frame, text="Paste from Clipboard", command=paste_table_data, style="Custom.TButton")
    paste_button.grid(row=7, column=0, padx=5, pady=5, sticky="w")

    clear_button = ttk.Button(frame, text="Clear Data", command=clear_table_data, style="Custom.TButton")
    clear_button.grid(row=7, column=1, padx=5, pady=5, sticky="w")

    send_email_button = ttk.Button(
        frame, text="Send Email",
        command=lambda: send_email(attachment_listbox), style="Custom.TButton"
    )
    send_email_button.grid(row=7, column=2, padx=5, pady=5, sticky="w")

    # Preview Email Button
    preview_email_button = ttk.Button(
        frame, text="Preview Email",
        command=preview_email, style="Custom.TButton"
    )
    preview_email_button.grid(row=7, column=3, padx=5, pady=5, sticky="w")

# ------------------- Data Table -------------------

def add_data_table(frame):
    """Add a Treeview table for displaying table data."""
    table_frame = tk.Frame(frame)
    table_frame.grid(row=8, column=0, columnspan=4, sticky="nsew", padx=20, pady=20)

    # Configure grid weights for resizing
    frame.grid_rowconfigure(8, weight=1)
    frame.grid_columnconfigure(1, weight=1)

    # <--- Change: Set height to 10 to reduce vertical space
    data_table = ttk.Treeview(table_frame, show="headings", style="Custom.Treeview", height=10)  # Set height=10

    data_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add scrollbar to the data table
    scrollbar_data = ttk.Scrollbar(table_frame, orient="vertical", command=data_table.yview)
    scrollbar_data.pack(side=tk.RIGHT, fill="y")
    data_table.configure(yscrollcommand=scrollbar_data.set)

    # Define tags for alternating row colors
    data_table.tag_configure('oddrow', background='#FFFFFF')
    data_table.tag_configure('evenrow', background='#D3D3D3')

    return data_table

# ------------------- Email Composition and Sending -------------------

def compose_email_content():
    """
    Compose the email content in both plain text and HTML formats.
    Returns:
        plain_text (str): The plain text version of the email.
        html_content (str): The HTML version of the email.
        subject (str): The email subject.
    """
    # Get the email subject from the input field
    subject = subject_entry.get().strip()

    # Get the custom greeting or use a default if empty
    custom_greeting = greeting_entry.get().strip()
    greeting = custom_greeting if custom_greeting else "Hi all,"

    # Get the Email Body input
    email_body_raw = email_body_text.get("1.0", tk.END).strip()
    if not email_body_raw:
        email_body_raw = "Please see below for details about a delivery expected to arrive at the company."

    # Escape HTML characters to prevent HTML injection
    greeting_escaped = html.escape(greeting)
    email_body_escaped = html.escape(email_body_raw)

    # Process the Email Body to preserve paragraph breaks
    # Split the text into paragraphs based on double newlines
    paragraphs = email_body_escaped.split('\n\n')
    # Replace single newlines with <br> and wrap paragraphs in <p> tags
    processed_paragraphs = ["<p>{}</p>".format(para.replace('\n', '<br>')) for para in paragraphs]
    email_body_html = "<br>".join(processed_paragraphs)

    # Start composing the HTML content
    html_content = "<html><body>"
    html_content += "<p>{}</p><br>{}<br>".format(greeting_escaped, email_body_html)  # Using .format()

    # Compose plain text content
    plain_text = f"{greeting}\n\n{email_body_raw}\n\n"

    # ------------------- Include Table Data Section -------------------
    if data_table.get_children():
        # HTML Table without width: 100%
        html_content += """
        <table style="border-collapse: collapse; table-layout: auto; text-align: left;">
            <thead style="background-color: lightgreen;">
                <tr>
        """
        for col in data_table["columns"]:
            html_content += f"<th style='border: 1px solid black; padding: 8px; text-align: left;'>{html.escape(str(col))}</th>"
        html_content += """
                </tr>
            </thead>
            <tbody>
        """
        for index, item in enumerate(data_table.get_children()):
            row_style = "background-color: #f9f9f9;" if index % 2 != 0 else "background-color: #ffffff;"
            html_content += f"<tr style='{row_style}'>"
            for value in data_table.item(item)['values']:
                # Ensure that cell data is properly escaped
                value_escaped = html.escape(str(value))
                html_content += f"<td style='border: 1px solid black; padding: 8px; text-align: left;'>{value_escaped}</td>"
            html_content += "</tr>"
        html_content += """
            </tbody>
        </table>
        """

        # Include the table data in plain text without the heading
        headers = data_table["columns"]
        if headers:
            # Create a header row
            header_row = "\t".join(headers)
            plain_text += header_row + "\n"
            # Create separator
            plain_text += "\t".join(['-' * len(header) for header in headers]) + "\n"
            # Add each data row
            for item in data_table.get_children():
                row = data_table.item(item)['values']
                row_text = "\t".join(str(value) for value in row)
                plain_text += row_text + "\n"
            plain_text += "\n"
    else:
        # Optional: If no data, you can choose to include a message or skip the table
        pass  # Do nothing if no data

    # Add closing remarks
    html_content += "<br>"
    html_content += "<p>Regards,</p>"
    html_content += """
    <p style='font-family: Georgia, serif; font-size: 11pt; color: #022a4d; margin-bottom: 0;'>
        <strong>Your Name</strong><br>
        <span style='font-size: 8pt;'>Your Position</span>
    </p>
    <p style='font-family: Georgia, serif; font-size: 11pt; color: #022a4d; margin-bottom: 5px;'>
        <strong>Company Name</strong><br>
        Address Line 1<br>
        Address Line 2<br>
        Tel: [Your Phone Number]<br>
        Fax: [Your Fax Number]<br>
        www.yourcompanywebsite.com<br>
        info@yourcompanywebsite.com
    </p>
    </body></html>
    """

    # Plain text closing remarks
    plain_text += "Regards,\n"
    plain_text += "Your Name\n"
    plain_text += "Your Position\n\n"
    plain_text += "Company Name\n"
    plain_text += "Address Line 1\n"
    plain_text += "Address Line 2\n"
    plain_text += "Tel: [Your Phone Number]\n"
    plain_text += "Fax: [Your Fax Number]\n"
    plain_text += "www.yourcompanywebsite.com\n"
    plain_text += "info@yourcompanywebsite.com\n"

    # Debugging: Print the composed HTML content
    logging.debug("----- Composed HTML Content -----")
    logging.debug(html_content)
    logging.debug("----- End of HTML Content -----\n")

    return plain_text, html_content, subject

def send_email(attachment_listbox):
    """Compose and send an email including the Table data and attachments."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Failed to initialize Outlook: {str(e)}")
        return

    # Compose email content
    plain_text, html_content, subject = compose_email_content()

    # Debugging: Print the subject and recipients
    logging.debug(f"Subject: {subject}")
    logging.debug(f"HTML Content Length: {len(html_content)}")
    logging.debug(f"Plain Text Content Length: {len(plain_text)}")

    if not subject:
        messagebox.showerror("Input Error", "Please enter the email subject.")
        return

    # Set email properties
    mail.Subject = subject
    mail.BodyFormat = 2  # olFormatHTML = 2
    mail.HTMLBody = html_content
    # mail.Body = plain_text  # Removed to ensure email is sent as HTML

    # Gather selected recipients from the selected_recipients dictionary
    to_recipients = selected_recipients['to']
    cc_recipients = selected_recipients['cc']

    # Debugging: Print the recipients
    logging.debug(f"To Recipients: {to_recipients}")
    logging.debug(f"CC Recipients: {cc_recipients}")

    if not to_recipients and not cc_recipients:
        messagebox.showerror("Recipient Error", "Please select at least one email recipient in To or CC.")
        return

    mail.To = ";".join(to_recipients)
    mail.CC = ";".join(cc_recipients)

    # Attach files
    attachments = list(attachment_listbox.get(0, tk.END))
    for file_path in attachments:
        if os.path.isfile(file_path):
            try:
                mail.Attachments.Add(Source=file_path)
                logging.debug(f"Attached file: {file_path}")
            except Exception as e:
                logging.error(f"Failed to attach file {file_path}: {e}")
                messagebox.showwarning("Attachment Error", f"Failed to attach file: {file_path}\n{e}")

    # Debugging: Confirm email properties before sending
    logging.debug(f"Sending Email to: {mail.To}")
    logging.debug(f"CC: {mail.CC}")
    logging.debug(f"Number of Attachments: {len(attachments)}")

    # Send the email
    try:
        mail.Send()
        messagebox.showinfo("Success", "Email sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {str(e)}")

# ------------------- Preview Email Function -------------------

def preview_email():
    """Preview the composed email in plain text and HTML formats with vertical scrollbars."""
    # Compose email content
    plain_text, html_content, subject = compose_email_content()

    # Create a new top-level window for preview
    preview_window = tk.Toplevel(root)
    preview_window.title("Email Preview")
    # <--- Change: Adjusted window geometry to reduce overall height
    preview_window.geometry("800x600")  # Previously "800x600"

    # Create a frame for the toggle buttons
    toggle_frame = tk.Frame(preview_window)
    toggle_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

    # Variable to track the selected view
    view_var = tk.StringVar(value="Plain Text")

    # Function to update the preview based on the selected view
    def update_preview():
        selected_view = view_var.get()
        if selected_view == "Plain Text":
            preview_html.pack_forget()
            preview_text.pack(fill=tk.BOTH, expand=True)
        else:
            preview_text.pack_forget()
            preview_html.pack(fill=tk.BOTH, expand=True)

    # Radio buttons for toggling views
    plain_text_rb = tk.Radiobutton(toggle_frame, text="Plain Text", variable=view_var, value="Plain Text", command=update_preview)
    plain_text_rb.pack(side=tk.LEFT, padx=5)

    html_rb = tk.Radiobutton(toggle_frame, text="HTML", variable=view_var, value="HTML", command=update_preview)
    html_rb.pack(side=tk.LEFT, padx=5)

    # Create a frame for the preview content
    content_frame = tk.Frame(preview_window)
    content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # ------------------- Plain Text Preview with Scrollbar -------------------
    # Frame to hold Text widget and scrollbar
    text_frame = tk.Frame(content_frame)
    text_frame.pack(fill=tk.BOTH, expand=True)

    # Text widget for plain text preview
    preview_text = tk.Text(text_frame, wrap='word', state='disabled')
    preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Scrollbar for plain text
    text_scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=preview_text.yview)
    text_scrollbar.pack(side=tk.RIGHT, fill="y")
    preview_text.configure(yscrollcommand=text_scrollbar.set)

    # Insert plain text content
    preview_text.configure(state='normal')
    preview_text.delete("1.0", tk.END)
    preview_text.insert(tk.END, plain_text)
    preview_text.configure(state='disabled')

    # ------------------- HTML Preview with Scrollbar -------------------
    # HTMLScrolledText widget for HTML preview (comes with built-in scrollbars)
    preview_html = HTMLScrolledText(content_frame, html=html_content, background="white")
    preview_html.pack(fill=tk.BOTH, expand=True)
    preview_html.pack_forget()  # Hide initially

    # Initially show plain text
    update_preview()

    # Make sure the preview window is above the main window
    preview_window.transient(root)
    preview_window.grab_set()
    root.wait_window(preview_window)

# ------------------- Data Pasting and Clearing -------------------

def paste_table_data():
    """Paste data from clipboard into a table."""
    try:
        # Get data from clipboard using win32clipboard
        data = get_clipboard_text()

        # Check if data is empty
        if not data.strip():
            messagebox.showerror("Paste Error", "Clipboard is empty.")
            return

        # Attempt to read data as tab-separated values (TSV)
        df = pd.read_csv(io.StringIO(data), sep='\t')

        # Clear existing data and columns in the table
        data_table.delete(*data_table.get_children())
        data_table["columns"] = list(df.columns)

        # Set up new columns
        for col in df.columns:
            data_table.heading(col, text=col, anchor="center")
            # Dynamically adjust column width based on content
            max_length = max(df[col].astype(str).map(len).max(), len(col))
            data_table.column(col, anchor="center", width=max(100, max_length * 10))

        # Insert new data with alternating row colors
        for index, row in df.iterrows():
            tag = 'oddrow' if index % 2 == 0 else 'evenrow'
            data_table.insert("", "end", values=list(row), tags=(tag,))

        messagebox.showinfo("Success", "Data pasted successfully from clipboard!")
    except pd.errors.EmptyDataError:
        messagebox.showerror("Paste Error", "No data found in clipboard.")
    except pd.errors.ParserError:
        messagebox.showerror("Paste Error", "Failed to parse clipboard data. Ensure it's tab-separated.")
    except ValueError as ve:
        messagebox.showerror("Paste Error", str(ve))
    except Exception as e:
        messagebox.showerror("Paste Error", f"Failed to paste data: {str(e)}")

def clear_table_data():
    """Clear all data from the table."""
    data_table.delete(*data_table.get_children())
    data_table["columns"] = []
    messagebox.showinfo("Clear Data", "Data table has been cleared.")

# ------------------- Helper Functions -------------------

def get_clipboard_text():
    """Retrieve text from the clipboard using win32clipboard."""
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        elif win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT).decode('utf-8')
        else:
            raise ValueError("Unsupported clipboard format.")
    except Exception as e:
        raise e
    finally:
        win32clipboard.CloseClipboard()
    return data

# ------------------- Assemble Tab Content -------------------

# Create input frame for the first tab
input_frame = tk.Frame(tab_emailer)
input_frame.grid(row=0, column=0, columnspan=4, sticky="nsew", padx=20, pady=20)

# Configure grid weights for resizing
tab_emailer.grid_rowconfigure(8, weight=1)
tab_emailer.grid_columnconfigure(1, weight=1)

# Add Subject input
subject_entry = add_subject_input(input_frame)

# Add Greeting and Email Body inputs
greeting_entry, email_body_text = add_greeting_email_body_inputs(input_frame)

# Add Recipient Selection Button and Summary
summary_label = add_recipient_selection_button(input_frame)

# Add Attachment Handling section
attachment_listbox = add_attachment_section(input_frame)

# Add Action Buttons (including Preview Email)
add_action_buttons(input_frame, attachment_listbox)

# Add Data Table
data_table = add_data_table(tab_emailer)

# ------------------- Create Templates Tab -------------------

# Create the Templates tab
templates_tab = ttk.Frame(notebook)
notebook.add(templates_tab, text="Templates")

# Define the template texts with multiple paragraphs
templates = [
    """Please see below for details about a delivery expected to arrive.

We are committed to ensuring timely delivery and will notify you of any changes promptly.""",

    """Please see below for order info.

If you have any questions or require further assistance, please don't hesitate to reach out to our support team.""",

    """Please see below for tracking information as requested.

You can track your package using the provided link or contact our support for more details."""
]

# Function to copy template text to Email Body
def copy_template_to_email_body(template_text):
    email_body_text.delete("1.0", tk.END)  # Clear existing text in Email Body
    email_body_text.insert(tk.END, template_text)  # Insert selected template text

# Create a label for the Templates tab
templates_label = tk.Label(templates_tab, text="Select a template to insert into the Email Body:", font=("Arial", 12))
templates_label.pack(pady=10, padx=20, anchor='w')

# Create buttons for each template
for i, template in enumerate(templates, start=1):
    template_button = ttk.Button(
        templates_tab,
        text=f"Template {i}",
        command=lambda text=template: copy_template_to_email_body(text),
        style="Custom.TButton"
    )
    template_button.pack(pady=10, padx=20, anchor='w')  # Adjust layout as needed

# Optionally, display the template text for user reference
templates_display_frame = tk.Frame(templates_tab)
templates_display_frame.pack(pady=10, padx=20, fill='both', expand=True)

templates_display_label = tk.Label(templates_display_frame, text="Available Templates:", font=("Arial", 12, "bold"))
templates_display_label.pack(anchor='w')

for i, template in enumerate(templates, start=1):
    template_text_label = tk.Label(
        templates_display_frame,
        text=f"Template {i}:\n{template}",
        wraplength=800,
        justify="left"
    )
    template_text_label.pack(anchor='w', pady=5)

# ------------------- Run the Application -------------------

root.mainloop()
