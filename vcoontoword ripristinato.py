import win32com.client
import datetime
import re
from docxtpl import DocxTemplate

def check_for_new_vcon_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items 


    def format_date(date_str, original_format='%m/%d/%Y', target_format='%d/%m/%y'):
        """Converts a date string from one format to another.

        Args:
            date_str: The input date string.
            original_format: The format of the input date string.
            target_format: The desired output format.

        Returns:
            The reformatted date string.
        """
        try:
            dt = datetime.datetime.strptime(date_str, original_format)
            return dt.strftime(target_format)
        except ValueError:
            print(f"Error: Could not convert date '{date_str}' with format '{original_format}'")
            return date_str  # Return the original value if conversion fails


    def convert_total(total_str):
        if total_str.endswith('M'):
            return "{:,.2f}".format(float(total_str[:-1]) * 1000000)  # Convert millions
        else:
            return "{:,.2f}".format(float(total_str.replace(',', '')) * 1000)  # Convert thousands


    for message in reversed(messages):
        if re.search(r'VCON', message.Subject): 
            #print("New VCON Ticket Email:")
            email_body = message.Body
            #print(email_body)

            
            
            try:
                
                # Regular expressions for extraction

                currency_pattern = r'\s*(EUR|USD)\s*'
                principal_pattern = r'\s*Principal\s*[:\-]*\s*(?:EUR|USD)?\s*(\d[\d,\.]*)\b'
    #            issue_date_pattern = r'(?:Dated(?:\sDate)?\s*:?\s*|Daté\s*:?\s*|(?:Dated\sDate)?\s*:?\s*)(\d{2}/\d{2}/\d{2})'
                # issue_date_pattern = r'(?:Dated|Daté|(?:Dated\sDate))\s*:\s*(\d{2}/\d{2}/\d{2})'
               # settle_date_pattern = r'((?:Dated\sDate)|?:Dated|Daté)\s*:\s*(\d{2}/\d{2}/\d{2,4})'
                settle_date_pattern = r'(?:Settlement|Règlement)\s*:\s*(\d{2}/\d{2}/\d{2,4})'
                # trade_date_pattern = r'Trade Date\s*:\s*(\d{2}/\d{2}/\d{2})'
                trade_date_pattern = r'(?:Trade\sDate|(?:As\sof\sDate)|(?:Transaction))\s*:\s*(\d{2}/\d{2}/\d{2,4})'
                total_pattern = r'(?:BUYS|ACHETE)\s*:\s*(\d+(?:,\d+)*M?)\b' 
                maturity_date_pattern = r'ENI\s+0\s+(\d{2}/\d{2}/\d{2})'  
                yield_pattern = r'\s*(?:Yield|Rdt)\s*:\s*([\d\.]+)'
                price_pattern = r'\s*(?:Price|Prix)\s*:\s*([\d\.]+)' 
                


                dealer_pattern = re.search(r'\((.*?)\)', email_body)

 

                proceeds_pattern = r'Proceeds Payable to the Issuer\s*\:\s*([\d\.,]+)\s+EUR'






                
                # Extract and store data in a dictionary
                trade_data = {}

                trade_data['currency'] = re.search(currency_pattern, email_body).group(1)
                trade_data['principal'] = re.search(principal_pattern, email_body).group(1)
                trade_data['settle_date'] = datetime.datetime.strptime(re.search(settle_date_pattern, email_body).group(1), "%m/%d/%y").strftime("%d/%m/%y") 
                # trade_data['issue_date'] = re.search(issue_date_pattern, email_body).group(1)
                #trade_data['issue_date'] = datetime.datetime.strptime(re.search(issue_date_pattern, email_body).group(1), '%d/%m/%Y').strftime('%d/%m/%y')
                # trade_data['trade_date'] = re.search(trade_date_pattern, email_body).group(1)
                trade_data['trade_date'] = datetime.datetime.strptime(format_date(re.search(trade_date_pattern, email_body).group(1)), "%m/%d/%y").strftime("%d/%m/%y") 
                trade_data['total'] = convert_total(re.search(total_pattern, email_body).group(1))
                # trade_data['maturity_date'] = re.search(maturity_date_pattern, email_body).group(1)
                trade_data['maturity_date'] = datetime.datetime.strptime(format_date(re.search(maturity_date_pattern, email_body).group(1)), "%m/%d/%y").strftime("%d/%m/%y") 
                trade_data['yield'] = f"{float(re.search(yield_pattern, email_body).group(1))}%"
                trade_data['price'] = re.search(price_pattern, email_body).group(1)


                # Logic to set value based on text_in_brackets
                if dealer_pattern.group(1) == "GOLDMAN SACHS INTL":
                    trade_data['dealerCode'] = "Euroclear 94589"
                    trade_data['dealerFull'] = "Goldman Sachs International"
                    trade_data['dealerShort'] = "GS"
                elif dealer_pattern.group(1) == "BNP PARIBAS":
                    trade_data['dealerCode'] = "Euroclear 99290"
                    trade_data['dealerFull'] = "BNP Paribas"
                    trade_data['dealerShort'] = "BNP"
                elif dealer_pattern.group(1) == "BAYERISCHE LANDESBAN":
                    trade_data['dealerCode'] = "Clearstream 51190"
                    trade_data['dealerFull'] = "Bayerische Landesbank"
                    trade_data['dealerShort'] = "BL"
                elif dealer_pattern.group(1) == "CREDIT AGRICOLE CIB":
                    trade_data['dealerCode'] = "Euroclear 913767"
                    trade_data['dealerFull'] = "Crédit Agricole Corporate and Investment Bank"
                    trade_data['dealerShort'] = "CA"         
                elif dealer_pattern.group(1) == "CITIGROUP GLOBAL MAR":
                    trade_data['dealerCode'] = "Euroclear 21159" 
                    trade_data['dealerFull'] = "Citigroup Global Markets Europe Limited"
                    trade_data['dealerShort'] = "CITI"
                elif dealer_pattern.group(1) == "ING":
                    trade_data['dealerCode'] = "Euroclear 22529" 
                    trade_data['dealerFull'] = "ING Bank N.V."
                    trade_data['dealerShort'] = "ING"        
                # ... add more elif blocks for other mappings ...
                else:
                    # Default value if no match 
                    trade_data['dealerCode'] = "Mapping not found" 
                    trade_data['dealerFull'] = "Mapping not found"
                    trade_data['dealerShort'] = "Mapping not found" 





                    
            #    trade_data['proceeds'] = re.search(proceeds_pattern, email_body).group(1) 

                for key, value in trade_data.items():
                    print(f"{key}: {value}")


                doc = DocxTemplate('CITI_Template.docx')

                # Prepare the data to replace the bookmark 
                context = {
                    'currency': trade_data['currency'],
                    'principal': trade_data['principal'],
                    'settle_date': trade_data['settle_date'],
                    'trade_date': trade_data['trade_date'],
                    'total': trade_data['total'],
                    'maturity_date': trade_data['maturity_date'],
                    'yield': trade_data['yield'],
                    'price': trade_data['price'],
                    'dealerCode': trade_data['dealerCode'],
                    'dealerFull': trade_data['dealerFull']                   
                }


                print(context)

                # Update the document content
                doc.render(context)

                # Save the changes
                doc.save('updated_prova.docx') 
               # doc.save(str(trade_data['issue_date']) + str(trade_data['dealerShort']) + str(trade_data['total']) + '.docx')

            
            # except AttributeError:
            #     print("Error: Email format may have changed. Could not extract data.")
            except re.error as e:
                print(f"Error: Regex pattern may be incorrect. Error message: {e}")
            
            
            break
        
            
check_for_new_vcon_emails()



#proceeds = TOTAL?
#anagrafica ENI 0 + maturity sempre?