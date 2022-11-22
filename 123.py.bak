import imaplib
import email
import os
from io import StringIO
from io import BytesIO

mail = imaplib.IMAP4_SSL('imap-mail.outlook.com', 993)
mail.login('sujithdc@live.com', 'Suji!2best')
mail.list()
mail.select("inbox")

result, data = mail.search(None, "ALL")
 
ids = data[0] # data is a list.
id_list = ids.split() # ids is a space separated string
latest_email_id = id_list[-1] # get the latest
 
result, data = mail.fetch(latest_email_id, "(RFC822)") # fetch the email body (RFC822) for the given ID
 
raw_email = data[0][1] # here's the body, which is raw text of the whole email
# including headers and alternate payloads

raw_email = data[0][1]
email_message = email.message_from_bytes(raw_email)

for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            fileName = part.get_filename()

            if bool(fileName):
                print (fileName)
                filePath = os.path.join('d:/test1/', 'attachments', fileName)
                if not os.path.isfile(filePath) :
                    print (fileName)
                    fp = open(filePath, 'wb')
                    print (part.get_payload(decode=True))
                    fp.write(part.get_payload(decode=True))
                    fp.close()
##imapSession.close()
##imapSession.logout()

##def get_first_text_block(self, email_message_instance):
##    maintype = email_message_instance.get_content_maintype()
##    if maintype == 'multipart':
##        for part in email_message_instance.get_payload():
##            if part.get_content_maintype() == 'text':
##                return part.get_payload()
##    elif maintype == 'text':
##        return email_message_instance.get_payload()
