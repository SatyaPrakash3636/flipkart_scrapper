from datetime import datetime
from flipkart.dev import Flipkart
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from flipkart import constant as const
import smtplib as s
import glob
import os
import pandas as pd
import pygsheets


def create_directories(dir_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)

def get_seller_details():
    bot = Flipkart()
    seller_details = bot.get_all_seller()
    bot.close()
    return seller_details

def get_reports(seller_name, save_to=None, seller_id=1, report_type="latest"):
    bot = Flipkart(save_to)
    bot.landing_page()
    bot.login_page()
    bot.select_seller(seller_id)
    downloaded_file = bot.earn_more(seller_name, report_type)
    bot.close()
    return downloaded_file

# Function for sending email with attachment
def send_email(password, file_name, seller_name):
    host = "smtp.gmail.com"
    port = 587
    username = "buffermailid@gmail.com"
    password = password
    from_addr = "no-reply@blooprint.in"
    # Comma seperated email ids
    # to_addr = "satyaprakash3636@gmail.com,ajay.m@blooprint.in"
    to_addr = "buffermailid@gmail.com"
    # To send content of file
    # msg_content = open(FileName).read()
    msg_content = "This is an automated email sent by <b>Python Bot</b>"
    now = datetime.now()
    date_today = now.strftime("%Y-%m-%d")
    msg = MIMEMultipart('alternative')
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = f"{seller_name.upper()} | Earn More Report - {date_today}"

    msg.attach(MIMEText(msg_content, 'html'))
    attachment = MIMEApplication(open(file_name, "rb").read())
    attachment.add_header('Content-Disposition','attachment', filename=file_name.split("\\")[-1])
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

def create_update_gsheet(seller, excel_file):
    gc = pygsheets.authorize(service_file='creds.json')
    df = pd.read_excel(excel_file, engine="openpyxl")
    files_in_drive = gc.spreadsheet_titles()
    print(f"Existing files: {files_in_drive}")
    if seller in files_in_drive:
        print(f"{seller} already present in drive, deleting and recreating....")
        sheet = gc.open(seller)
        sheet.delete()
    else:
        print(f"{seller} not present in drive, creating...")
    sheet = gc.create(title=seller, folder_name="python-test")
    wksheet = sheet.worksheet(0)
    wksheet.set_dataframe(df, (1,1), extent=True)
    wksheet.title = "earn-more-report"

# Directory where all the reports will be downloaded
main_dir = os.getcwd() + "\earn_more_reports"

# Check if main dir exixt or not, if not then create
create_directories(main_dir)

# Get all seller details 
# Sample output - [[1, 'ODYOSONIC', 'Odyosonic Private Limited'], [2, 'Express-Group', 'Krishan LaL & Co.'],....]
seller_details = get_seller_details()
print(seller_details)

# Modify names
# Sample - ['odyosonic', 'express_group', 'defsxhn', 'kolorr',....]
seller_names = []
seller_name_dir = []
i = 0
for seller in seller_details:
    i += 1
    s_name = seller[1].lower().replace(" ", "_").replace("-", "_")
    seller_names.append(s_name)
    s_dir = main_dir + "\\" + s_name
    seller_name_dir.append([i, s_dir, s_name])

# Create sub directories
for i in seller_names:
    seller_dir = main_dir +  "\\" + i
    create_directories(seller_dir)

# Delete any old excel workbook with name earn_more_report*.xlsx
old_files = glob.glob("earn_more_reports/*/*.xlsx")
if old_files:
    for f in old_files:
        os.remove(f)

# Getting Reports
for seller in seller_name_dir:
    if seller[0] < 11:
        downloaded_file = get_reports(seller_name=seller[2], save_to=seller[1], seller_id=seller[0], report_type="weekly")
        print(downloaded_file)
        # send_email(const.GMAIL_PASS, downloaded_file, seller[2])
        create_update_gsheet(seller[2], downloaded_file)

    # get_reports(save_to=seller[1], seller_id=seller[0])

