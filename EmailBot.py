"""
A simple email sender utility

Author: Yujun Zeng & Chuan Zhang
Email: yujunchuanz@gmail.com
Date: 2023-10-29
"""
# TODO: wrap up functions to a class
import pandas as pd
import smtplib
from smtplib import SMTPAuthenticationError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openai
import json

# Set your OpenAI API key here
openai.api_key = "XXXXXXXXXXXXXXXXXXXXXXX"

# Function to load templates
def load_templates(name):
    try:
        with open(name+".json", "r", encoding="utf-8") as file:
            templates = json.load(file)
        return templates
    except Exception as e:
        print("Error loading email templates:", str(e))
        return {}

# Function to generate email content using ChatGPT
def generate_email_content(prompt):
    try:
        response = openai.Completion.create(
            engine="text-davinci-002",
            prompt=prompt,
            max_tokens=100
        )
        return response.choices[0].text
    except Exception as e:
        messagebox.showerror("Error", "Error generating email content:" + str(e))
        return ""

def generate_email_content_from_ui():
    prompt = prompt_text.get("1.0", "end")
    if not prompt:
        messagebox.showwarning("Warning", "Please enter a prompt.")
        return

    generated_content = generate_email_content(prompt)
    if not generated_content:
        messagebox.showerror("Error", "Failed to generate email content.")
        return

    message_text.delete("1.0", "end")
    message_text.insert("1.0", generated_content)

# Function to read Excel data
def read_excel(file_path):
    try:
        data = pd.read_excel(file_path)
        return data
    except Exception as e:
        print("Error reading Excel file:", str(e))
        return None

# Create a function to update the message module when a template is selected
def select_template(event):
    selected_template = template_combobox.get()
    message_text.delete("1.0", "end")  # Clear the message module
    message_text.insert("1.0", templates.get(selected_template, ""))

# Function to update the subject when a template is selected
def select_subject_template(event):
    selected_template = subject_combobox.get()
    subject_entry.delete(0, "end")  # Clear the subject entry
    subject_entry.insert(0, selected_template)

def process_excel(message_template,data,recipients,subjects):
    filled_messages, n = [], len(data)
    if len(subjects) == 0:
        if "<subject>" in data.columns:
            subjects = data["<subject>"]
        else:
            raise "need title(s)! either choose one or have <subject> in the file"
    else:
        subjects = [subjects for _ in range(n)]

    for index, row in data.iterrows():
        filled_message = message_template
        filled_subject = subjects[index]
        # Replace placeholders in the message with corresponding values from the DataFrame
        for column in data.columns:
            placeholder = "<" + column[1:-1] + ">"  # Extract column name from "<...>"
            value = row[column]
            filled_message = filled_message.replace(placeholder, str(value))
            filled_subject = filled_subject.replace(placeholder, str(value))
        filled_messages.append(filled_message)
        subjects[index] = filled_subject

    if len(recipients) == 0:
        if "<recipient>" in data.columns:
            recipients = data["<recipient"]
        elif "<email>" in data.columns:
            recipients = data["<email>"]
        elif "<to>" in data.columns:
            recipients = data["<to>"]
    else:
        recipients = [recipients for _ in range(n)]

    return recipients, subjects, filled_messages

# Function to send emails
def send_emails(server, sender_email, password, subjects, messages, recipients, signature_html):
    try:
        server, code = server.split(" ")
        with smtplib.SMTP_SSL(server, int(code)) as smtp:
            # smtp.starttls()
            smtp.login(sender_email, password)

            for i in range(len(recipients)):
                recipient_email = recipients[i]
                print("Sending to "+recipient_email)
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = subjects[i]
                email_body_plain = messages[i]
                email_body_plain = email_body_plain.replace("\n", '<br>')
                # Combine the plain text email body and HTML signature into a single HTML content
                combined_html = f"""\
                <html>
                <head></head>
                <body>
                <p>{email_body_plain}</p>
                {signature_html}
                </body>
                </html>
                """
                msg.attach(MIMEText(combined_html, 'html'))
                smtp.sendmail(sender_email, recipient_email, msg.as_string(), 'html')

        messagebox.showinfo("Success", "Emails sent successfully!")
    except SMTPAuthenticationError:
        messagebox.showerror("Error", f"Check your username and password!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to handle generating email content and sending emails
def send_emails_from_ui():
    server = smtp_server_entry.get()
    sender_email = from_entry.get()
    recipients = to_entry.get()
    password = password_entry.get()
    subject = subject_entry.get()
    file_path = file_path_entry.get()
    message_content = message_text.get("1.0", "end")
    sender_name_cn = sender_name_cn_entry.get()
    sender_name_en = sender_name_en_entry.get()
    sender_mobile = sender_mobile_entry.get()

    if not (server and sender_email and password and sender_name_en and sender_name_cn and sender_mobile):
        messagebox.showwarning("Warning", "Please fill server/sender_email/password/signature.")
        return

    with open('signature.html', 'rb') as signature_file:
        signature_content_bytes = signature_file.read()
        signature_content = signature_content_bytes.decode('utf-8', errors='ignore')
        signature_content = signature_content.replace("sender_name_cn", str(sender_name_cn))
        signature_content = signature_content.replace("sender_name_en", str(sender_name_en))
        signature_content = signature_content.replace("sender_mobile", str(sender_mobile))
        signature_content = signature_content.replace("sender_mail", str(sender_email))
    signature_html = "<p>" + signature_content + "</p>"

    if len(file_path) > 0:
        data = read_excel(file_path)
        [recipients, subject, message_content] = process_excel(message_content, data, recipients, subject)
    else:
        recipients, subject, message_content = [recipients], [subject], [message_content]
    send_emails(server, sender_email, password, subject, message_content, recipients, signature_html)


# Create the main UI window
root = tk.Tk()
root.title("Email Sender")

# Create UI components
default_value = tk.StringVar()
default_value.set("smtp.exmail.qq.com 465")
smtp_server_label = tk.Label(root, text="SMTP Server - Port:")
smtp_server_label.pack()
smtp_server_entry = tk.Entry(root, textvariable=default_value)
smtp_server_entry.pack()

default_value = tk.StringVar()
default_value.set("yujun.zeng@future-cap.com")
from_label = tk.Label(root, text="From:")
from_label.pack()
from_entry = tk.Entry(root, textvariable=default_value)
from_entry.pack()

default_value = tk.StringVar()
default_value.set("XXX")
password_label = tk.Label(root, text="Password:")
password_label.pack()
password_entry = tk.Entry(root, show="*", textvariable=default_value)
password_entry.pack()


to_label = tk.Label(root, text="To:")
to_label.pack()
to_entry = tk.Entry(root)
to_entry.pack()

subjects = load_templates("subjects")
# Create a combobox to select templates
subject_label = tk.Label(root, text="Select Subject:")
subject_label.pack()
subject_combobox = ttk.Combobox(root, values=list(subjects.keys()))
subject_combobox.pack()
# Create an entry widget for typing the subject
subject_entry_label = tk.Label(root, text="or Type Your Subject:")
subject_entry_label.pack()
subject_entry = tk.Entry(root, width=80)
subject_entry.pack()

# Create a combobox to select templates
templates = load_templates("templates")
template_label = tk.Label(root, text="Select Template:")
template_label.pack()
template_combobox = ttk.Combobox(root, values=list(templates.keys()))
template_combobox.pack()

file_path_label = tk.Label(root, text="Excel File:")
file_path_label.pack()
file_path_entry = tk.Entry(root, width=80)
file_path_entry.pack()
browse_button = tk.Button(root, text="Browse", command=lambda: file_path_entry.insert(0, filedialog.askopenfilename()))
browse_button.pack()

default_value = tk.StringVar()
default_value.set("曾宇君")
sender_name_cn_label = tk.Label(root, text="Signature Name (CN):")
sender_name_cn_label.pack()
sender_name_cn_entry = tk.Entry(root, textvariable=default_value)
sender_name_cn_entry.pack()

default_value = tk.StringVar()
default_value.set("Lulu Zeng")
sender_name_en_label = tk.Label(root, text="Signature Name (EN):")
sender_name_en_label.pack()
sender_name_en_entry = tk.Entry(root, textvariable=default_value)
sender_name_en_entry.pack()

default_value = tk.StringVar()
default_value.set("+86 18801069906")
sender_mobile_label = tk.Label(root, text="Signature Mobile:")
sender_mobile_label.pack()
sender_mobile_entry = tk.Entry(root, textvariable=default_value)
sender_mobile_entry.pack()

send_button = tk.Button(root, text="Send Emails", command=send_emails_from_ui)
send_button.pack()

prompt_label = tk.Label(root, text="Prompt:")
prompt_label.pack()
prompt_text = tk.Text(root, height=1, width=80)
prompt_text.pack()

message_label = tk.Label(root, text="Message:")
message_label.pack()
message_text = tk.Text(root, height=10, width=80)
message_text.pack()
message_text.insert("1.0", templates["Instruction"])

send_button = tk.Button(root, text="Generate", command=generate_email_content_from_ui)
send_button.pack()

# Bind the events to update the subject and message modules
subject_combobox.bind("<<ComboboxSelected>>", select_subject_template)
template_combobox.bind("<<ComboboxSelected>>", select_template)

# Start the GUI main loop
root.mainloop()
