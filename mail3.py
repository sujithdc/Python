import imaplib
import base64
imap_user = input('Email: ')
imap_password = input('Password: ')

conn = imaplib.IMAP4_SSL('imap-mail.outlook.com', 993)

try:
    (retcode, capabilities) = conn.login(imap_user, imap_password)
except:
    print (sys.exc_info()[1])
    sys.exit(1)

conn.select(readonly=1) # Select inbox or default namespace
(retcode, messages) = conn.search(None, '(UNSEEN)')
if retcode == 'OK':
    for num in messages[0].split(b' '):
        print ('Processing :', num)
        typ, data = conn.fetch(num,'(RFC822)')
        msg = email.message_from_string(data[0][1])
        typ, data = conn.store(num,'-FLAGS','\\Seen')
        if ret == 'OK':
            print (data,'\n',30*'-')
            print (msg)

conn.close()