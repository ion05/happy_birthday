import pandas as pd  # reading our excel file
import datetime  # for intracting with time and year
import smtplib  # for a connection between our gmail & python script
from email.message import EmailMessage  # sending emails
import os  # intracting with the system
from dotenv import load_dotenv

load_dotenv()
password = os.getenv("PASSWORD")


def sendEmail(to, sub, msg):
    print(f"email to {to} \nsend with subject: {sub} \n message: {msg}")
    email = EmailMessage()
    email['from'] = "Aayan Agarwal"
    email['to'] = f"{to}"
    email['subject'] = f'{sub}'
    email.set_content(f"{msg}")

    with smtplib.SMTP(host="smtp.gmail.com", port=587) as smtp:

        smtp.ehlo()
        smtp.starttls()
        smtp.login("Email", password)
        smtp.send_message(email)
        print("Email Send")
    pass


if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    print(df)
