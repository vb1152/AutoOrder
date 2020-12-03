
# Description

This is python GUI app, that read data.xlsx, calculate order and then create/send xlsx files with order data for small groсery shop. Calculations are made for every scu, labeled in "Reserve" sheet with insurance reserve and supplier index.

## Table of Contents

- data.xlsx - data to use in calculation.
  - "Data" sheet: balance and turnovers of all scu (from 1C program)
    - "PriceList" sheet: price list.
    - "Reserve" sheet: Scu with codes and barcodes.
      - ResData column is insurance reserve.
      - Multipl column is multiplicity.
      - Supplier col is supplier index from Produsers sheet.
    - Produsers sheet: produsers data. Index, name and email to send order.

- data.txt. Data for emails. Sender email, letter subject ets.
- email_text.txt - text for emails.

## Usage

This app is used with data from 1C business application (1С:Retail Automated retail outlet and store accounting) for retail store.
Data about balance and turnovers should be copied to Data sheet. Price list to PriceList sheet.
In Reserve sheet user write insurance reserve, multiplicity and supplier index for every scu.
When app is opened, it load data about produsers
![Open app](/images/main.png)

User can choose producer to create order.
![Choose](/images/create.png)

When order created, user can open file and change it, if needed.
When file is sended successfully - user gets a notification.
![Send](/images/sended.png)

To send emails by this app from user_email@gmail.com user need **token.pickle and credentials** in one folder with app.py.
Use this link for details: <https://developers.google.com/gmail/api/quickstart/python>
Store these files in the same folder as app.py

Order for every scu calculated by formula:
order = Sales + insurance reserve – final balance
If scu has multiplicity, the order is calculated accordingly.
