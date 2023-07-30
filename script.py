import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
import os
import time

# Load the xlsx file
file_path = 'recipients.xlsx'
wb = load_workbook(file_path)
ws = wb.active

# Read the email list from the xlsx file
email_list = pd.read_excel(file_path, sheet_name=ws.title)

# Define the email subject and body
subject = 'Rostlinolékařský pas k vaší faktuře č. '
body = '''
<html>
<head>
<style>
    body {
        font-family: Arial, sans-serif;
        font-size: 12px;
    }
    .larger-text {
        font-size: 18px;
    }
</style>
</head>
<body>
    Vážený zákazníku,
<br>
<br>
ode dne 14. prosince 2019 nabývá účinnosti nové nařízení Evropského parlamentu, a to nám ukládá jako prodejci povinnosti při prodeji rostlin dodat zákazníkovi Rostlinolékařský pas. Je to z důvodu, aby se evidoval pohyb prodávaných rostlin po území Evropské unie.
<br> 
<br>
<b>Více informací, proč došlo k této povinnosti se můžete dočíst na našem webu, v horním panelu (Nákupy, registrace). Vám, jakožto zákazníkovi, z toho neplyne žádná povinnost a na tuto automaticky generovanou zprávu neodpovídejte.</b> 
<br>
<br>
<b class="larger-text">Více informací o nařízení</b><br>
<br>
Nařízení Evropského parlamentu a Rady (EU) 2016/2031 o ochranných opatřeních proti škodlivým organismům rostlin (dále jen „nařízení“). Dle čl. 65 tohoto nařízení je pro internetové prodejce rostlin, rostlinných produktů a jiných předmětů, podléhajících fytosanitární regulaci (dále jen regulované komodity), stanovena povinnost registrace pro rostlinolékařské účely, a to bez výjimky. Dále je dle čl. 79 a čl. 81 nařízení stanovena povinnost opatřovat regulované komodity při internetovém obchodování (smlouvy uzavřené na dálku) rostlinolékařským pasem, a to i v případě dodávek těchto komodit přímo konečným uživatelům. 
<br>
<br>
Veškeré informace ohledně zákazu dovozu určitých rostlin, zvláštních a rovnocenných požadavcích, které musí při dovozu na území EU nebo při přemísťování na tomto území, vysoce rizikových rostlinách, rostlinných produktech či jiných předmětech, výjimkách z požadavku na rostlinolékařské osvědčení pro malá množství určitých rostlin naleznete na stránkách ÚKZÚZ http://eagri.cz/public/web/ukzuz/portal/ <br>
<br>
<br>
Rostlinolékařský pas naleznete v příloze. <br>

</body>
</html>
'''

# Specify the folder path containing the attachments
attachments_folder = r'C:\Users\Martin\Desktop\script\pdf'

def send_email(recipient, bcc_recipient, attachment_name):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    attachment_name_without_ext, _ = os.path.splitext(attachment_name)  # Remove the .pdf suffix
    mail.Subject = f'Rostlinolékařský pas k vaší faktuře č. {attachment_name_without_ext}'  # Format the subject with the attachment name without the .pdf suffix
    mail.HTMLBody = body
    mail.To = recipient
    mail.BCC = bcc_recipient

    if not pd.isna(attachment_name):
        attachment_path = os.path.join(attachments_folder, attachment_name)
        if os.path.isfile(attachment_path):
            mail.Attachments.Add(attachment_path)
        else:
            print(f"Attachment not found: {attachment_path}. Email not sent to {recipient}.")
            return

    mail.Send()
    time.sleep(5)  # Wait for 5 seconds

for _, row in email_list.iterrows():
    recipient = row['Email']
    bcc_recipient = 'martinpenkava1@gmail.com'
    attachment_name = row['Attachment']
    print(f"Připravuji mail pro {recipient}...")  # Print the recipient of the current email
    send_email(recipient, bcc_recipient, attachment_name)

print('Emaily uspěšně odeslány.')