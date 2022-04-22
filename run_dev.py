from flipkart.dev import Flipkart
import os

def create_directories(dir_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)

def get_seller_details():
    bot = Flipkart()
    seller_details = bot.get_all_seller()
    bot.close()
    return seller_details

def get_reports(save_to=None, seller_id=1, report_type="latest"):
    bot = Flipkart(save_to)
    bot.landing_page()
    bot.login_page()
    bot.select_seller(seller_id)
    bot.earn_more(report_type)
    bot.close()

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
    seller_name_dir.append([i, s_dir])

# Create sub directories
for i in seller_names:
    seller_dir = main_dir +  "\\" + i
    create_directories(seller_dir)

# Getting Reports
for seller in seller_name_dir:
    if seller[0] < 2:
        get_reports(save_to=seller[1], seller_id=seller[0])
    # get_reports(save_to=seller[1], seller_id=seller[0])

