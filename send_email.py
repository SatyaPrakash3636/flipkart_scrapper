from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from flipkart import constant as const
import smtplib as s

def send_email(password, file_name, seller_name):
    host = "smtp.gmail.com"
    port = 587
    username = "buffermailid@gmail.com"
    password = password
    from_addr = username
    # Comma seperated email ids
    to_addr = "satyaprakash3636@gmail.com,buffermailid@gmail.com"

    # To send content of file
    # msg_content = open(FileName).read()
    msg_content = """This is an automated email sent by a <b>Python Bot</b>
    \n
    \n
    Regards
    """

    msg = MIMEMultipart('alternative')
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = f"Earn More Report - {seller_name.upper()}"

    msg.attach(MIMEText(msg_content, 'html'))
    attachment = MIMEApplication(open(file_name, "rb").read())
    attachment.add_header('Content-Disposition','attachment', filename=file_name)
    msg.attach(attachment)
    email_data = msg.as_string()

    try:
        conn = s.SMTP(host, port)
        conn.ehlo()
        conn.starttls()
        conn.login(username, password)
        conn.sendmail(from_addr, to_addr.split(','), email_data)
        conn.quit()
    except s.SMTPAuthenticationError:
        print("Incorrect username/password combination. Please try again...")
    except s.SMTPException as error:
        print(error)
    else:
        print("Email sent successfully...")

downloaded_file = "F:\Python_learning\selenium\earn_more_reports\odyosonic\odyosonic-weekly-2022-04-25_00-37.xlsx"

send_email(const.GMAIL_PASS, downloaded_file, "odyosonic")