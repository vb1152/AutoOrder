Description. This is python GUI app, that read data.xlsx, calculate order and then create/send xlsx files with order data for small grossery shop. 

Table of Contents. 
-	data.xlsx - data to use in calculaton. 
    - "Data" sheet: balance and turnovers of all scu (from 1C program)
    - "PriceList" sheet: price list. 
    - "Reserve" sheet: Scu with codes and  barcodes. 
        - ResData column is insurance reserve. 
        - Multipl column is multiplicity. 
        - Supplier col is supplier index from Produsers sheet.
    - Produsers sheet: produsers data. Index, name and email to send order. 

- data.txt. Data for emails. Sender email, letter subject ets. 
- email_text.txt - text for emails. 

Usage. These app is used with data from 1C business application (1С:Retail Automated retail outlet and store accounting).
Data about balance and turnovers should be copied to Data sheet. Price list to PriceList sheet. 
In Reserve sheet user puts down insurance reserve, multiplicity and supplier index. 

To send emails from ...@gmail.com user need files token.pickle and credentials. 
https://developers.google.com/gmail/api/quickstart/python 
Store these files in the same folder as app.py 

Order for every scu calculates by formula: 
order = Sales + insurance reserve – final balance
If scu has multiplicity, the order is calculated accordingly 


