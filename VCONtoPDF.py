import win32com.client
from datetime import datetime
import re
import shutil
from docxtpl import DocxTemplate
import docx2pdf

import os
import logging  

logging.basicConfig(filename='date_errors.log', level=logging.WARNING)

def check_for_new_vcon_emails():
    destination_directory = "R:/ECP Eni SpA"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items 


    def format_date(date_str, original_formats, target_format="%d/%m/%y"):
        """Handles multiple date formats with error logging."""
        for fmt in original_formats:
            try:
                dt = datetime.strptime(date_str, fmt)  
                return dt.strftime(target_format)
            except ValueError:
                pass  

        logging.warning(f"Error: Could not convert date '{date_str}' with formats {original_formats}")
        return None



    def convert_total(total_str):
        if total_str.endswith('M'):
            return "{:,.2f}".format(float(total_str[:-1]) * 1000000)  
        else:
            return "{:,.2f}".format(float(total_str.replace(',', '')) * 1000)  


    for message in reversed(messages):
        if re.search(r'VCON', message.Subject): 
            email_body = message.Body
            
            
            try:
                
                currency_pattern = r'\s*(EUR|USD)\s*'
                principal_pattern = r'\s*Principal\s*[:\-]*\s*(?:EUR|USD)?\s*(\d[\d,\.]*)\b'
                settle_date_pattern = r'(?:Settlement|Règlement)\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4}(\d{2,4})?)' 
                trade_date_pattern = r'(?:Trade\sDate|(?:As\sof\sDate)|(?:Transaction))\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4}(\d{2,4})?)' 
                total_pattern = r'(?:BUYS|ACHETE)\s*:\s*(\d+(?:,\d+)*M?)\b' 
                maturity_date_pattern = r'ENI\s+0\s+(\d{2}/\d{2}/\d{2})'  
                yield_pattern = r'\s*(?:Yield|Rdt)\s*:\s*([\d\.]+)'
                price_pattern = r'\s*(?:Price|Prix)\s*:\s*([\d\.]+)' 
                dealer_pattern = re.search(r'\((.*?)\)', email_body)


                # Extract and store data in a dictionary
                trade_data = {}
                possible_date_formats = ['%m/%d/%y','%m/%d/%Y', '%d/%m/%y', '%d/%m/%Y', '%B/%d/%Y']  
               
                
                trade_data['currency'] = re.search(currency_pattern, email_body).group(1)
                if trade_data['currency']=="EUR":
                    trade_data['currency_symbol']="€"
                elif trade_data['currency']=="USD":
                    trade_data['currency_symbol']="$"
                else: trade_data['currency_symbol']= "Mapping not found" 

                trade_data['principal'] = re.search(principal_pattern, email_body).group(1)
                trade_data['settle_date'] = format_date(re.search(settle_date_pattern, email_body).group(1), possible_date_formats)
                trade_data['trade_date'] = format_date(re.search(trade_date_pattern, email_body).group(1), possible_date_formats)
                trade_data['maturity_date'] = format_date(re.search(maturity_date_pattern, email_body).group(1), possible_date_formats)
                trade_data['tenor'] = round((datetime.strptime(trade_data['maturity_date'], "%d/%m/%y") - datetime.strptime(trade_data['trade_date'], "%d/%m/%y")).days / 30)
                month_text = "month" if trade_data['tenor'] == 1 else "months"
                trade_data['total'] = convert_total(re.search(total_pattern, email_body).group(1))
                trade_data['yield'] = f"{float(re.search(yield_pattern, email_body).group(1))}%"
                trade_data['price'] = re.search(price_pattern, email_body).group(1)


                # Logic to set value based on text_in_brackets
                if dealer_pattern.group(1) == "GOLDMAN SACHS INTL":
                    trade_data['dealerCode'] = "Euroclear 94589"
                    trade_data['dealerFull'] = "Goldman Sachs International"
                    trade_data['dealerShort'] = "GS"
                elif dealer_pattern.group(1)  in ("BNP PARIBAS FORTIS", "BNP PARIBAS"):
                    trade_data['dealerCode'] = "Euroclear 99290"
                    trade_data['dealerFull'] = "BNP Paribas"
                    trade_data['dealerShort'] = "BNP"
                elif dealer_pattern.group(1) == "BAYERISCHE LANDESBAN":
                    trade_data['dealerCode'] = "Clearstream 51190"
                    trade_data['dealerFull'] = "Bayerische Landesbank"
                    trade_data['dealerShort'] = "BL"
                elif dealer_pattern.group(1) == "CREDIT AGRICOLE CIB":
                    trade_data['dealerCode'] = "Euroclear 91376"
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
                elif dealer_pattern.group(1) == "BARCLAYS BANK PLC":
                    trade_data['dealerCode'] = "Clearstream 21864" 
                    trade_data['dealerFull'] = "Barclays Bank Ireland PLC"
                    trade_data['dealerShort'] = "BARCLAYS"     

                else:
                    trade_data['dealerCode'] = "Mapping not found" 
                    trade_data['dealerFull'] = "Mapping not found"
                    trade_data['dealerShort'] = "Mapping not found" 



                for key, value in trade_data.items():
                    print(f"{key}: {value}")


                doc = DocxTemplate('ECP_Template.docx')

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


                # Update the document content
                doc.render(context)

                # Save 
                doc.save('Update.docx') 
                def format_total(total_value):
                    return str(total_value).replace(",000,000.00", "M")
                filename = "{}_{}_{}.pdf".format(trade_data['dealerShort'], format_total(trade_data['total']), trade_data['yield'])
                doc.save(filename[:-4]+'.docx') 
                docx2pdf.convert("Update.docx", filename)
                docx2pdf.convert("Update.docx", destination_directory+"/"+filename)
                email_filename = f"{trade_data['dealerShort']}_{format_total(trade_data['total'])}_{trade_data['yield']}.msg"
                email_path = os.path.join(destination_directory, email_filename)
                message.SaveAs(email_path)
            
            except re.error as e:
                print(f"Error: Regex pattern may be incorrect. Error message: {e}")
            
            
            break
                
    # Email sending 
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  


    # mail.To = "dario.pozzi@eni.com;"
    mail.To = "TREASURY@eni.com; Fabio.Valerio@eni.com"
    mail.CC = "Paolo.Ferla@eni.com; Paolo.Barra@eni.com, Derivatives.Backoffice@eni.com"
    mail.Subject = f"Emissione Nuova ECP"
    

    mail.HTMLBody = f"""
    <html>
    <body>

    <p>
    Ciao Fabio,
    <br>
    <br>
    in allegato il Form of Notification per una CP emessa oggi:
    <br>
    <br>

    <table>
        <tbody>
            <tr>
            <td>Issue Size</td>
            <td>   </td> 
            <td>{trade_data['currency_symbol']}{format_total(trade_data['total'])[:-1]}mln</td>
            </tr>
            <tr>
            <td>Dealer</td>
            <td>   </td> 
            <td>{trade_data['dealerFull']}</td>
            </tr>
            <tr>
            <td>YTM</td>
            <td>   </td> 
            <td>{trade_data['yield']}</td>
            </tr>
            <tr>
            <td>Tenor</td>
            <td>   </td> 
            <td>{trade_data['tenor']} {month_text}</td>
            </tr>
            <td>Settlement</td>
            <td>   </td> 
            <td>{trade_data['settle_date']}</td>
            </tr>
            <td>Maturity</td>
            <td>   </td> 
            <td>{trade_data['maturity_date']}</td>
            </tr>
        </tbody>
    </table>

    Trovi il modulo da firmare e la relativa email di conferma nella cartella di scambio:
    Trovi il modulo da firmare (e la relativa mail di conferma) nella <a href="file:///\\Ennf1001\scambio\ECP Eni SpA">cartella di scambio</a>.
    <p>


    </body>
    </html>
    """
    
    # Attach the PDF
    attachment_path = os.path.abspath(filename)  
    mail.Attachments.Add(Source=attachment_path)

    mail.Send()
    print("Email sent successfully!")



    #Salvataggio su disco condiviso
    
    source_file = os.path.abspath(attachment_path)
    



check_for_new_vcon_emails()

