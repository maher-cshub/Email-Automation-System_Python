import smtplib, pandas,time
from secret_file import SENDER_NAME,SENDER_EMAIL,PASSWORD
import smtplib, ssl
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart # New line
from datetime import date
import re

TODAY = date.today().strftime("%B %d, %Y")
SUBJECT = "Web Developer Application"
JOB_TITLE = "Re: Application for the position of web developer"

def readExcel(dest):
    return pandas.read_excel(dest)

##email body
html_file = "./index.html"


# Creating a SMTP session | use 587 with TLS, 465 SSL and 25
server = smtplib.SMTP('smtp.gmail.com', 587)
# Encrypts the email
context = ssl.create_default_context()
server.starttls(context=context)
# We log in into our Google account
server.login(SENDER_EMAIL, PASSWORD)

# Try to log in to server and send email

def SendEmail(sender_email,company_email,email_content):
    try:
        # Sending email from sender, to receiver with the email body
        server.sendmail(sender_email, company_email, email_content.as_string())
        print(f'Email sent! to {company_email}')
    except Exception as e:
        print(f'Oh no! Something bad happened!\n {e}')


def main():
    email_content = ""
    with open(html_file,mode="r",encoding="utf-8") as f:
        email_content = f.read()

    ##get emails
    receivers = readExcel("./receivers.xlsx").to_records()  

    ##send email for each one
    for receiver in receivers:
        if (len(receiver) == 7):
            print(f"....Sending email to {receiver[5]}")
            ##email setup
            receiver_name = receiver[1]
            company_name = receiver[2]
            company_location = receiver[3]
            company_phone = receiver[4]
            company_email = receiver[5]
            date_sent = receiver[6]
            if (date_sent == "none"):

                ##suit email for receiver
                email_content = re.sub("TODAY_DATE",TODAY,email_content)
                email_content = re.sub("JOB_TITLE",JOB_TITLE,email_content)
                email_content = re.sub("Receiver Name",receiver_name,email_content)
                email_content = re.sub("Hiring Manager",receiver_name,email_content)
                email_content = re.sub("Company Name",company_name,email_content)
                email_content = re.sub("company_name",company_name,email_content)
                email_content = re.sub("Company Location",company_location,email_content)
                email_content = re.sub("Company Phone",company_phone,email_content)
                email_content = re.sub("Company Email",company_email,email_content)
                email = MIMEMultipart()
                email['To'] = formataddr((receiver_name, company_email))
                email['From'] = formataddr((SENDER_NAME, SENDER_EMAIL))
                email['Subject'] = SUBJECT
                email.attach(MIMEText(email_content, 'html'))
                SendEmail(SENDER_EMAIL,company_email,email)
                receivers[receiver[0]][6] = TODAY
            
            time.sleep(2)

    ##save changes to excel
    receivers = pandas.DataFrame(receivers)
    receivers.drop(columns=receivers.columns[0], axis=1, inplace=True)
    receivers.to_excel("receivers.xlsx",index=False)

    print("Closing Server")
    server.quit()

main()

