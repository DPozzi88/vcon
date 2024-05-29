import win32com.client

import logging  

import os

logging.basicConfig(filename='date_errors.log', level=logging.WARNING)

def check_for_new_vcon_emails():


    # Email sending 
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  


    mail.To = "dario.pozzi@eni.com;"
    # mail.To = "Teodoro.Digiulio@eni.com; eriberto.fraternale@eni.com"
    #mail.To = "TREASURY@eni.com; Fabio.Valerio@eni.com"
    #mail.CC = "Paolo.Ferla@eni.com; Paolo.Barra@eni.com, Derivatives.Backoffice@eni.com"
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
   
            </tr>
        </tbody>
    </table>


    Trovi il modulo da firmare (e la relativa mail di conferma) nella <a href="file:///\\Ennf1001\scambio\ECP Eni SpA">cartella di scambio</a>.
    <p>


    </body>
    </html>
    """
    
    filename = "Update.docx"


    # Attach the PDF
    attachment_path = os.path.abspath(filename)  
    mail.Attachments.Add(Source=attachment_path)


        
    # Attach the PDF


    mail.Send()
    print("Email sent successfully!")




check_for_new_vcon_emails()

