# Overview
This vba code is intended to prepare emails with the same content and attachments to a large number of recipients. In the excel file, in dedicated cells, enter the title of the e-mail, the path of the word file with the e-mail content and the path to attachments. An unlimited number of attachments is allowed. In the "sender" cell, you can enter an e-mail address other than the main one - if we are authorized to send e-mails, the sender will be the given e-mail. Otherwise, the mail will be sent as: "on behalf of". Enter the e-mail recipients in a separate sheet. VBA downloads the email text, title and attachments and creates a model email in Outlook. In the next stage, it creates copies of the model e-mail with one recipient downloaded from Excel. In the next step, it is possible to send mass mailings of all emails generated with the code and to delete working folders in Outlook.

# Technology
To use this code need to install Outlook and Excel.

# How to run code
Dowload Excel file. Fill in the appropriate fields and press the dedicated buttons.

