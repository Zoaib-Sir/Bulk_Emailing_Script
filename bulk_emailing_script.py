import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd


def send_email_with_pdf(sender_emails, sender_passwords, receiver_emails, subject, body, pdf_path, rotation_size):
    
    with open(pdf_path, 'rb') as f:
        pdf_data = f.read()

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        
        length = len(sender_emails)
        count = 0
        i = -1
        for recipient_email in receiver_emails:

            if(i==-1):
                i = 0
                server.quit()
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(sender_emails[i], sender_passwords[i])

            if(count == rotation_size):
                i += 1
                count = 0
            
                if(i==length):
                    i = 0

                server.quit()
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(sender_emails[i], sender_passwords[i]) 


            msg = MIMEMultipart()

            msg['Content-Type'] = 'multipart/mixed'

            msg['From'] = sender_emails[i]
            msg['To'] = recipient_email
            msg['Subject'] = subject

            text_part = MIMEText(body, 'plain')
            msg.attach(text_part)

            pdf_attachment = MIMEApplication(pdf_data, _subtype="pdf")
            pdf_attachment.add_header('Content-Disposition', 'attachment', filename='test_file.pdf')
            msg.attach(pdf_attachment)

            server.send_message(msg)

            print(f"Email with PDF attachment sent successfully to {recipient_email} from {sender_emails[i]}")

            count += 1

        server.quit()

    except Exception as e:
        print(f"Error sending email to {recipient_email}", e)


emails_path = input("Enter text file name containing email adrresses of receivers : ") + ".txt"

with open(emails_path, 'r') as file:
    receiver_emails = [line.strip() for line in file.readlines()]

file_path = input("Enter excel file name containing sender emails and passwords : ") + ".xlsx"
df = pd.read_excel(file_path) 
sender_emails = df.iloc[:, 0].tolist()
sender_passwords = df.iloc[:, 1].tolist()
rotation_size = int(input("Enter rotation size : "))
subject = input("Enter subject of the email : ")
body = input("Enter contents of the email : ")
pdf_path = input("Enter file name of the pdf attachment you want to send : ") + ".pdf"

send_email_with_pdf(sender_emails, sender_passwords, receiver_emails, subject, body, pdf_path, rotation_size)
