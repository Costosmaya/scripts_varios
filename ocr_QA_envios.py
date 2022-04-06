from email.message import EmailMessage
from PIL import Image
import pytesseract
import imaplib
import numpy as np
import email
from email import policy
from io import BytesIO
import os
server = "imap.gmail.com"
imap = imaplib.IMAP4_SSL(server)

username = "costos@mayaprin.com"
password = "Mayaprin100%"

imap.login(username, password)


res,messages = imap.select('"Env&AO0-os Power BI/Producci&APM-n"')

messages = int(messages[0])

n = 36
counter = 0
for i in range(messages,0, -1):
    if counter < n:
        res, msg = imap.fetch(str(i),"(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1], policy=policy.default)
                for attachment in msg.iter_attachments():
                    if 'image/png' in attachment.get_content_type():
                        img = Image.open(BytesIO(attachment.get_content()))
                        img.save(f'C:/Users/User/Documents/Analisis  Desarrollo Costos/Scripts/Python/imgs_qa/{i}.png','PNG')
                        counter+=1
        
    else:
        break;
                   

path = 'C:\\Users\\User\\Documents\\Analisis  Desarrollo Costos\\Scripts\\Python\\imgs_qa'

list_dir = os.scandir(path)
blank_contador = 0
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
for item in list_dir:
    img = np.array(Image.open(item.path))
    text = pytesseract.image_to_string(img)
    if '(En bla' in text:
        blank_contador+=1
print(f'Cantidad Reportes en Blanco:{blank_contador}, ratio:{round(blank_contador/(n-1),2)}')