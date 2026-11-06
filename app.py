
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = select_file('Please select a file in the folder you want to use.').split('/')
    del directory[-1]
    directory = '\\'.join(directory) + '\\'
    filenames = ['ProductData.xlsx', 'MasterInvoice.pdf', "Order1.xlsx", 'Order2.xlsx', 'Order3.xlsx','Order4.xlsx']
    file_cwd = os.listdir(directory)
    completed_list = []
    for file in file_cwd:
        if file in filenames:
            completed_list.append(file)
        else:
            pass
    if len(completed_list)<6:
        show_message('Error', 'Please select a valid folder.')
    else:
        break

# INSERT ANSWER FOR QUESTION 2 HERE
prod_desc = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in prod_desc.iterrows():
    description = row.Description
    unit_number = row['Unit number']
    product_descriptions[unit_number] = description
prod_price = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for index, row in prod_price.iterrows():
    price = row['Price per unit']
    unit_number = row['Unit number']
    product_prices[unit_number] = price

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num, filename, directory):
    update_dict = {}
    workbook = openpyxl.load_workbook(directory + filename, data_only=True)
    sheet = workbook['OrderForm']
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = f'{sheet['C9'].value.month}/{sheet['C9'].value.day}/{sheet['C9'].value.year}'
    update_dict['CustomerAddress1'] = sheet['C16'].value
    update_dict['CustomerName'] = sheet['C14'].value
    update_dict['CustomerCompany'] = str(sheet['C15'].value)
    update_dict['CustomerAddress2'] = f'{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}'
    update_dict['Email'] = sheet['C13'].value
    update_dict['Total'] = '0'
    if sheet['B24'].value != None:
        update_dict['Price0'] = str(product_prices[str(sheet['B24'].value)])
        update_dict['Quantity0'] = str(sheet['C24'].value)
        update_dict['Amount0'] = str(float(product_prices[str(sheet['B24'].value)]) * sheet['C24'].value)
        update_dict['Description0'] = product_descriptions[sheet['B24'].value]
        update_dict['Total'] = str((float(update_dict['Total']) + float(update_dict['Amount0'])))
        update_dict['Price0'] = f"${float(update_dict['Price0']):,.2f}"
        update_dict['Amount0'] = f"${float(update_dict['Amount0']):,.2f}"
        if sheet['B25'].value != None:
            update_dict['Price1'] = str(product_prices[str(sheet['B25'].value)])
            update_dict['Quantity1'] = str(sheet['C25'].value)
            update_dict['Amount1'] = str(float(product_prices[str(sheet['B25'].value)]) * sheet['C25'].value)
            update_dict['Description1'] = product_descriptions[sheet['B25'].value]
            update_dict['Total'] = str((float(update_dict['Total']) + float(update_dict['Amount1'])))
            update_dict['Price1'] = f"${float(update_dict['Price1']):,.2f}"
            update_dict['Amount1'] = f"${float(update_dict['Amount1']):,.2f}"
            if sheet['B26'].value != None:
                update_dict['Price2'] = str(product_prices[str(sheet['B26'].value)])
                update_dict['Quantity2'] = str(sheet['C26'].value)
                update_dict['Amount2'] = str(float(product_prices[str(sheet['B26'].value)]) * sheet['C26'].value)
                update_dict['Description2'] = product_descriptions[sheet['B26'].value]
                update_dict['Total'] = str((float(update_dict['Total']) + float(update_dict['Amount2'])))
                update_dict['Price2'] = f"${float(update_dict['Price2']):,.2f}"
                update_dict['Amount2'] = f"${float(update_dict['Amount2']):,.2f}"            
                if sheet['B27'].value != None:
                    update_dict['Price3'] = str(product_prices[str(sheet['B27'].value)])
                    update_dict['Quantity3'] = str(sheet['C27'].value)
                    update_dict['Amount3'] = str(float(product_prices[str(sheet['B27'].value)]) * sheet['C27'].value)
                    update_dict['Description3'] = product_descriptions[sheet['B27'].value]
                    update_dict['Total'] = str((float(update_dict['Total']) + float(update_dict['Amount3'])))
                    update_dict['Price3'] = f"${float(update_dict['Price3']):,.2f}"
                    update_dict['Amount3'] = f"${float(update_dict['Amount3']):,.2f}"
    update_dict['Total'] = f"${float(update_dict['Total']):,.2f}"
    return update_dict, update_dict['CustomerName'], update_dict['Email'], update_dict['Date']
order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = single_input('Please type in your email address.')
    gmail_password_path = gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()
    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84602
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

# INSERT ANSWER FOR QUESTION 6 HERE
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1

