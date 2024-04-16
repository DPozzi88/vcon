import win32com.client
import time
import re

def check_for_new_vcon_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items 

    for message in messages:
        if re.search(r'VCON TICKET:', message.Subject): 
            print("New VCON Ticket Email:")
            email_body = message.Body
            print(email_body)

            dealer_pattern = re.search(r'\((\w+)\)', email_body)
            if dealer_pattern:
                text_in_brackets = dealer_pattern.group(1)
                print(text_in_brackets)
                  
            else:
                print("No text found within round brackets")
            
        

check_for_new_vcon_emails()
