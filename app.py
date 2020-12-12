import tkinter as tk
import pandas as pd

# The tkinter.ttk module provides access to the Tk themed widget set
from tkinter import ttk, INSERT 
import xlsxwriter
from datetime import datetime
import os
from os import system
import logging

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
from apiclient import errors
from email import encoders
from datetime import datetime


window = tk.Tk()

logging.basicConfig(filename='app_log.log', level=logging.INFO)

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
#['https://mail.google.com/'] #, 'https://www.googleapis.com/auth/gmail.send']
#https://www.googleapis.com/auth/gmail.readonly

# functions from quickstart
def create_message(sender, to, subject, message_text):
    """Create a message for an email.

    Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.

    Returns:
        An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    
    return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}

def create_message_with_attachment(
    sender, to, subject, message_text, file):
    """Create a message for an email.

    Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

    Returns:
    An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    
    if main_type == 'text':
        fp = open(file, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(file, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(file, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(file, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(file)

    logging.info("Відправка Order " + filename + " на " + to)

    encoders.encode_base64(msg)

    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)
    raw = base64.urlsafe_b64encode(message.as_bytes())
    raw = raw.decode()

    return {'raw': raw} #

def send_message(service, user_id, message):
    """Send an email message.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message: Message to be sent.

    Returns:
        Sent Message.
    """
    try:
        message = (service.users().messages().send(userId=user_id, body=message)
                .execute())
        print ('Message Id: %s' % message['id'])
        logging.info("%s: %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), message)

        return message
    except errors.HttpError as error:
        print ('An error occurred: %s' % error)

def quickstart(file_to_send, email_to_send):

    
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
            
    #https://github.com/googleapis/google-api-python-client/issues/299
    service = build('gmail', 'v1', credentials=creds, cache_discovery=False)

    #open file with data for email
    f = open('data.txt', 'r')
    
    #open file with text for email 
    ft = open('email_text.txt', 'r')    
    
    #read lines from a file    
    lines = f.readlines()
    text = ft.read()

    my_email_sender = lines[1]
    to = email_to_send
    my_subject = lines[1]
    my_message_text = text

    f.close
    ft.close

    my_file = file_to_send

    sent = send_message(service,'me', create_message_with_attachment(my_email_sender, to, my_subject, my_message_text, my_file))        

    return sent


# function to make orders 
def order(list_order):
    # dict to store data from test.xlsx, reading using pd.read_excel 
    data = {}
    # read data from data.xlsx.
    with pd.ExcelFile('data.xlsx') as xls:
        data['Data'] = pd.read_excel(xls, 'Data', index_col=None, na_values=['NA', 'NAN'])
        data['Data'] = data['Data'].fillna(0)
        data['PriceList'] = pd.read_excel(xls, 'PriceList', index_col=None, na_values=['NA'])
        data['Reserve'] = pd.read_excel(xls, 'Reserve', index_col=None, na_values=['NA'])
        
        # https://stackoverflow.com/questions/36684013/extract-column-value-based-on-another-column-pandas-dataframe
        # extract column value based on another column pandas dataframe
    
    # create dict to store emails of supliers for sending orders 
    list_order_dict = {}

    # find an index for every suplier in list of supliers, labeled by user 
    for i in list_order:

        # find index of user labled supplier  
        supl_index = data1.loc[data1['Produser'] == i, 'Index'].iloc[0]

        # making new data to create new order. 
        data_order = []

        # add columns to the data_order
        for k in range(data['Reserve'].index.size):

            # if supplier index and safery stock is in data, list is created 
            # если установлены индекс поставщика и Страховой запас в листе Reserve, формируется list со списком продуктов с двумя цифрами 
            if supl_index == data['Reserve'].loc[k, 'Supplier'] and data['Reserve'].loc[k, 'ResData'] >= 0:
                data_order.append({'SCU name': None,
                                           'Barcode': data['Reserve'].loc[k, 'Barcode'],
                                            'Code': data['Reserve'].loc[k, 'Code'],
                                           'ResData': data['Reserve'].loc[k, 'ResData'],
                                           'Multipl': data['Reserve'].loc[k, 'Multipl'], #Multiplість. Додано нове поле
                                           'Order': None,
                                           'Price': None,
                                           'Sum': None,
                                           '№': None
                                            })

        # calculate Order 
        for l in range(len(data_order)):
            for m in range(data['Data'].index.size):
                if data_order[l]['Code'] == data['Data'].loc[m, 'Code']:
                    data_order[l]['Order'] = data['Data'].loc[m, 'Sale'] + data_order[l]['ResData'] - data['Data'].loc[m, 'Balance']
                    # check if 
                    if data_order[l]['Order'] > data_order[l]['Multipl']:
                        data_order[l]['Order'] = ((data_order[l]['Order'] // data_order[l]['Multipl'])+1)*data_order[l]['Multipl']

        # if 'Order' == None,  'Order' = 'ResData'
        for y in range(len(data_order)):
            if data_order[y]['Order'] == None:
                data_order[y]['Order'] = data_order[y]['ResData']
            
        # remove dicts from list of data where data_order[k]['Order'] <= 0
        data_order = list(filter(lambda x : x['Order'] > 0, data_order))

        for yt in range(len(data_order)):
            if data_order[yt]['Multipl'] != 'nan' and data_order[yt]['Order'] < data_order[yt]['Multipl']:
                data_order[yt]['Order'] = data_order[yt]['Multipl']

        # if no items for order, next suplier 
        if len(data_order) == 0: 
            continue
        
        data_to_order_file = open('data.txt', 'r')
        lines_data = data_to_order_file.readlines()
        
        f = lines_data[10]

        # create file for order, named by suplier        
        path = (f+i+'.xlsx')
        writer = pd.ExcelWriter(path, engine='xlsxwriter')

        # find email of of user labled supplier and store it in a dict with key E:/'+ i +'.xlsx', 
        # where i is name of suplier
        list_order_dict[path] = data1.loc[data1['Produser'] == i, 'email'].iloc[0]

        # variable to calculate sum of order             
        order_sum = 0

        # fill in data_order whith data 
        for n in range(len(data_order)):
            for o in range(data['PriceList'].index.size):
                if data_order[n]['Code'] == data['PriceList'].loc[o, 'Code']:

                    # add prices to ordered items
                    data_order[n]['Price'] = data['PriceList'].loc[o, 'whPrice']

                    # calculate Sum for every item in order 
                    data_order[n]['Sum'] = data_order[n]['Price'] * data_order[n]['Order']

                    # calc overall sum of order
                    order_sum += data_order[n]['Sum']

                    # fill '№' for every item in order
                    data_order[n]['№'] = n + 1

                    # add SCU name from PriceList. and remove spaces before SCU name
                    data_order[n]['SCU name'] = data['PriceList'].loc[o, 'Name'].strip()

        # add row with sum of overall order to a list of dicts 
        data_order.append({'SCU name': 'Total', 'Sum': order_sum})            

        # making DataFrame from list of dicts   
        data_order_frame = pd.DataFrame(data=data_order)

        # arange an order of columns in dataframe
        data_order_frame = data_order_frame[['№', 'SCU name', 'Barcode', 'Order', 'Price', 'Sum']]
        
        # pass DataFrame to function "to_excel" with diferent parameters
        # index = False - stop write index column to a excel
        # startcol = 1 - start write frome 1 col in excel sheet
        startrow = 10
        startcol = 0
        data_order_frame.to_excel(writer, sheet_name='Order',
                                  header=False,
                                  startrow = startrow, #start writing dataframe from row
                                  startcol=startcol,
                                  index = False)
      
        # add formating to excel file
        # Get the xlsxwriter objects from the dataframe writer object.
        #http://xlsxwriter.readthedocs.io/working_with_pandas.html
        workbook  = writer.book
        worksheet = writer.sheets['Order']

        # Add a header format for a dataframe.
        header_format = workbook.add_format({'bold': True, 'text_wrap': True,'valign': 'center',
                                               'fg_color': '#D7E4BC', 'border': 1})

        # Add a header format for headear in order.    
        header_format1 = workbook.add_format({'bold': True, 'text_wrap': False,'valign': 'left',
                                               'font_size': 15})

        header_info = workbook.add_format({'text_wrap': False,'valign': 'left',
                                               'font_size': 10})

        # sett datetime for now
        date_time = datetime.now()
        date_time_str = date_time.strftime("%d.%m.%y, %H:%M")
        
        # data from txt files
        a = lines_data[5]
        b = lines_data[6]
        c = lines_data[7]
        e = lines_data[8]
        d = lines_data[9]

        #worksheet.write(row, column, format)
        worksheet.write(0, startcol, a, header_format1) # ФОП Білич 
        worksheet.write(1, startcol, b, header_format1) # Адреса 
        worksheet.write(3, startcol, d + date_time_str, header_info) # Order створено 
        #worksheet.write_datetime(4, startcol+1, date_time, date_format) # дата/час 
        worksheet.write(len(data_order_frame)+16, startcol, c, header_info) #
        worksheet.write(len(data_order_frame)+17, startcol, e, header_info) #

        # Write the column headers with the defined format in 5 row.
        for col_num, value in enumerate(data_order_frame.columns.values):
            worksheet.write(startrow-1, col_num, value, header_format)

        # Set the alignment for data in the cell
        cell_format = workbook.add_format()
        cell_format.set_align('center')

        # Set the number format for a Price і Sum column
        cell_format04 = workbook.add_format()
        cell_format04.set_num_format('0.00')

        # Set the number format for a cell Barcodeи
        cell_format05 = workbook.add_format()
        cell_format05.set_num_format('0')
        cell_format05.set_align('center')


        # Set format format for last row Total 
        last_row_format = workbook.add_format({'bold': True, 'font_size': 12})
        last_row_format.set_num_format('0.00')
        worksheet.set_row(len(data_order_frame)+startrow-1, 20, last_row_format)
    
        worksheet.set_row(4, 0.1)
        worksheet.set_row(5, 0.1)
        worksheet.set_row(6, 0.1)
        worksheet.set_row(7, 0.1)
        worksheet.set_row(8, 0.1)
 
        worksheet.set_column(0, 0, 5, cell_format)   # 1 Column '№' width set to 5
        worksheet.set_column(1, 1, 50)   # 2 Columns Назва width set to 50
        worksheet.set_column(2, 2, 15, cell_format05)   # # Column  Barcode width set to 16
        worksheet.set_column(3, 3, 9, cell_format)  # Columns Order width set to 12
        worksheet.set_column('E:F', 9, cell_format04)  # Columns Price і Sum E-F width set to 8
       
        writer.save()

        # This will create a LabelFrame for every order 
        label_frame = tk.LabelFrame(main_label_frame, text = i, font=("Times New Roman", 13)) 
        label_frame.grid(column=0, row=list_order.index(i), sticky='W')

        # Buttons for LabelFrame. This will create two buttons for every order. 
        #https://stackoverflow.com/questions/16969168/how-can-i-get-the-name-of-a-widget-in-tkinter   - button_path = path
        # button_path take value of path for every loop
        btn1 = tk.Button(label_frame, text = 'Open', command = lambda button_path = path: os.startfile(button_path)) #os.startfile
        btn1.grid(column=0, row=2, sticky='W') #https://stackoverflow.com/questions/1585322/is-there-a-way-to-perform-if-in-pythons-lambda #list_order.index(i), button_path_v #format(list_order.index(i) + 1, '1.1f')
        btn2 = tk.Button(label_frame, text = 'Send', command = lambda button_path_v = path: lbl_success.insert(format(list_order.index(i) + 1, '1.1f'), 'Order sended ' + button_path_v +'\n') if quickstart(button_path_v, list_order_dict[button_path_v])['labelIds'] == ['SENT'] else print('False'))# , print(lambda: data1.loc['email'] == 'vsa'))
        btn2.grid(column=1, row=2, sticky='W') 
       
# This will creare text widget 
lbl_success = tk.Text(bg="white", width = 50, height = 20)
lbl_success.grid(column=0, row=30, rowspan=5,) 

# function to create a list of marked suppiers 
def callback():
    order_list= []

    for item in buttons:
        v, n = item
        if v.get():
            v.set(1)
            order_list.append(item[1])
        else:
            v.set(0)       

    order(order_list)
    
# columns in GUI
def rowcheck(a):
    if a > 9:
        return a%18 + 8  
    else:
        return a + 8  

# reset all data
def update():
    #clear all checkbuttons
    for b in buttons:
        b[0].set(0)    

    for widget in main_label_frame.winfo_children():
        widget.destroy()
    pass

# Window Title 
window.title("Calculate orders")

window.geometry("1400x800")

#ttk style pack for tkinter
# stile of all buttons
ttk.Style().configure("TButton", padding=6, relief="flat", background="#ccc")

# what are buttons does when pressed 
style = ttk.Style()
style.map("C.TButton", foreground=[('pressed', 'red'), ('active', 'blue')],
            background=[('pressed', '!disabled', 'black'), ('active', 'white')])

# creating Checkbutton for every suplier
buttons=[]

var1=tk.IntVar()
var1.set(0)

sett={}
sett['Produsers'] = pd.read_excel(pd.ExcelFile('data.xlsx'), 'Produsers', index_col=None, na_values=['NA'])

data1 = sett['Produsers']

for i in range(data1.index.size):
    var = tk.IntVar() 
    var.set(0)
    cb = tk.Checkbutton(text=data1.loc[i, 'Produser'], variable=var)
    cb.grid(column=i%5 + 1, row=rowcheck(i), sticky='W')
    buttons.append((var, data1.loc[i, 'Produser']))
    
# button to close app
button0 = ttk.Button(window, text="Close", command=window.destroy, style="C.TButton")
# position of a button
button0.grid(column=1, row=0)  
   
# LABEL
title = tk.Label(text="Create Order", font=("Times New Roman", 20), width = 25)
title.grid(column=0, row=0)

# button to make order
submit = ttk.Button(text="Make order", command=callback)
submit.grid(column=1, row=1)

# button to update
reset = ttk.Button(text="Clear", command=update)
reset.grid(column=2, row=1)

#text = "Створено нові файли"#, text = "Створено нові файли", font=("Times New Roman", 20), 
main_label_frame = tk.Frame(width = 15) #, yscrollcommand = scrollbar.set) 
main_label_frame.grid(column=0, row=2, rowspan = 200, sticky='N')

window.mainloop()