import smtplib
import pandas as pd
import os
import time
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')

def send_results():
    data = pd.read_csv('data/Student_Summary_Results.csv')
    with open('templates/email_template.html', 'r') as f:
        html_content = f.read()

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        print("Connected to server. Starting email loop...")
        
        for index, row in data.iterrows():
            recipient_email = row['Enter Your email']
            student_name = row['Enter Your Name']

            msg = MIMEMultipart()
            msg['From'] = f"Evaluation Portal <{EMAIL_USER}>"
            msg['To'] = recipient_email
            msg['Subject'] = f"Assessment Results for {student_name}"

            personalized_html = html_content.replace('{{name}}', str(student_name)) \
                .replace('{{reg_no}}', str(row['Enter Your RegNo'])) \
                .replace('{{total_marks}}', str(row['Total_Score'])) \
                .replace('{{total_percent}}', f"{row['Total_Percentage']:.1f}") \
                .replace('{{fe_marks}}', str(row['Frontend'])) \
                .replace('{{fe_percent}}', f"{row['Frontend_Percentage']:.1f}") \
                .replace('{{be_marks}}', str(row['Backend'])) \
                .replace('{{be_percent}}', f"{row['Backend_Percentage']:.1f}") \
                .replace('{{db_marks}}', str(row['DB'])) \
                .replace('{{db_percent}}', f"{row['DB_Percentage']:.1f}")

            msg.attach(MIMEText(personalized_html, 'html'))

            server.send_message(msg)
            print(f"[{index+1}/{len(data)}] Email sent to {student_name} ({recipient_email})")
            
            time.sleep(2)

        server.quit()
        print("\nAll emails sent successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    send_results()