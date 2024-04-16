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
            
            try:
                # Regular expressions for extraction

                currency_pattern = r'Principal\s*:\s*(\w+)' 
                principal_pattern = r'Principal\s*:\s*EUR\s*(\d+,\d+,\d+\.\d+)'
                issue_date_pattern = r'Dated\s*:\s*(\d{2}/\d{2}/\d{2})'
                trade_date_pattern = r'Trade Date\s*:\s*(\d{2}/\d{2}/\d{2})'
                maturity_date_pattern = r'ENI\s+0\s+(\d{2}/\d{2}/\d{2})'  
                yield_pattern = r'Yield\s*:\s*([\d\.]+)'
                price_pattern = r'Price\s*:\s*([\d\.]+)'
                proceeds_pattern = r'Proceeds Payable to the Issuer\s*\:\s*([\d\.,]+)\s+EUR'

                # Extract and store data in a dictionary
                trade_data = {}

                trade_data['currency'] = re.search(currency_pattern, email_body).group(1)
                trade_data['principal'] = re.search(principal_pattern, email_body).group(1)
                trade_data['issue_date'] = re.search(issue_date_pattern, email_body).group(1)
                trade_data['trade_date'] = re.search(trade_date_pattern, email_body).group(1)
                trade_data['maturity_date'] = re.search(maturity_date_pattern, email_body).group(1)
                trade_data['yield'] = re.search(yield_pattern, email_body).group(1)
                trade_data['price'] = re.search(price_pattern, email_body).group(1)
            #    trade_data['proceeds'] = re.search(proceeds_pattern, email_body).group(1) 

                for key, value in trade_data.items():
                    print(f"{key}: {value}")

            
            except AttributeError:
                print("Error: Email format may have changed. Could not extract data.")

check_for_new_vcon_emails()



#proceeds = TOTAL?
#anagrafica ENI 0 + maturity sempre?