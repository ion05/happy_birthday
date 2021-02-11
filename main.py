import pandas as pd  # reading our excel file
import datetime  # for intracting with time and year
import smtplib  # for a connection between our gmail & python script
from email.message import EmailMessage  # sending emails
import os  # intracting with the system
from dotenv import load_dotenv

load_dotenv()
password = os.getenv("PASSWORD")
emaill = os.getenv("EMAIL")


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
        smtp.login(email, password)
        smtp.ehlo()
        smtp.send_message(email)
        print("Email Send")
    pass


# Here if __name__ == “main” block is used to prevent (certain) code from being run when the module is imported in another program.
if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    # print(type(today))
    update = []
    yearnow = datetime.datetime.now().strftime("%Y")
    for index, item in df.iterrows():
        bday = item["Birthday"].strftime("%d-%m")
        if (bday == today) and yearnow not in str(item["Year"]):
            sendEmail(item["email"], 'Happy Birthday' +
                      item["name"], item["message"])
            update.append(index)
        for i in update:
            yr = df.loc[i, "Year"]
            df.loc[i, "Year"] = f"{yr}, {yearnow}"
        df.to_excel("data.xlsx", index=False)
