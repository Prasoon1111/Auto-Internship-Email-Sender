import tkinter as tk
from tkinter import filedialog, messagebox
import smtplib
import pandas as pd
import random
import time
from email.message import EmailMessage
import os

# ================= EMAIL FUNCTION =================

def send_emails():

    sender_email = email_entry.get()
    app_password = password_entry.get()
    excel_path = excel_file_path.get()
    resume_path = resume_file_path.get()

    if not sender_email or not app_password or not excel_path or not resume_path:
        messagebox.showerror("Error", "Please fill all fields")
        return

    try:
        data = pd.read_excel(excel_path)

        templates = [

f"""Dear Hiring Team at {{company}},

My name is Prasoon Ranjan, an Electronics and Communication Engineering student.
I am highly interested in exploring internship opportunities at {{company}}.

I am eager to contribute, learn, and grow within your organization.
Please find my resume attached for your review.

Thank you for your time and consideration.

Best regards,
Prasoon Ranjan
""",

f"""Hello {{company}} Team,

I hope this message finds you well.
I am Prasoon Ranjan, currently pursuing ECE and actively seeking internship opportunities.

I would be grateful for the opportunity to contribute to {{company}} while gaining valuable industry experience.

Kindly find my resume attached.

Looking forward to your response.

Sincerely,
Prasoon Ranjan
""",

f"""Respected {{company}} Recruitment Team,

I am writing to express my interest in internship roles at {{company}}.
As an ECE student passionate about technology and innovation, I am eager to apply my knowledge in a practical environment.

Please find my resume attached for your consideration.

Thank you for your time.

Warm regards,
Prasoon Ranjan
"""
        ]

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, app_password)

        for index, row in data.iterrows():

            company = row["company"]
            receiver = row["email"]

            subject = f"Internship Application | {company} | Prasoon Ranjan"

            body_template = random.choice(templates)
            body = body_template.format(company=company)

            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = sender_email
            msg["To"] = receiver
            msg.set_content(body)

            with open(resume_path, "rb") as f:
                file_data = f.read()
                file_name = os.path.basename(resume_path)

            msg.add_attachment(file_data,
                               maintype="application",
                               subtype="pdf",
                               filename=file_name)

            server.send_message(msg)

            delay = random.randint(25, 45)
            time.sleep(delay)

        server.quit()
        messagebox.showinfo("Success", "All emails sent successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# ================= GUI =================

root = tk.Tk()
root.title("Prasoon Auto Internship Email Sender")
root.geometry("520x380")

tk.Label(root, text="Your Gmail").pack()
email_entry = tk.Entry(root, width=45)
email_entry.pack()

tk.Label(root, text="App Password").pack()
password_entry = tk.Entry(root, width=45, show="*")
password_entry.pack()

tk.Label(root, text="Excel File (company, email columns)").pack()
excel_file_path = tk.StringVar()

def browse_excel():
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_file_path.set(filename)

tk.Entry(root, textvariable=excel_file_path, width=45).pack()
tk.Button(root, text="Browse Excel", command=browse_excel).pack(pady=5)

tk.Label(root, text="Resume PDF File").pack()
resume_file_path = tk.StringVar()

def browse_resume():
    filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    resume_file_path.set(filename)

tk.Entry(root, textvariable=resume_file_path, width=45).pack()
tk.Button(root, text="Browse Resume", command=browse_resume).pack(pady=5)

tk.Button(root, text="SEND EMAILS", bg="green", fg="white",
          font=("Arial", 12), command=send_emails).pack(pady=15)

root.mainloop()