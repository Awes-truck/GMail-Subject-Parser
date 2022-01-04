import imaplib
import email
import csv

mail = imaplib.IMAP4_SSL("imap.gmail.com")

# User inputs their Gmail Credentials
while True:
    username = input("GMail Email: ")
    password = input("GMail Password: ")
    try:
        mail.login(username, password)
        print ("Logged in as %r!" %username)
        break
    except imaplib.IMAP4.error:
        print ("[!] ERROR: Log in failed [!]")

# Select the appropriate folder/label
mail.select("\"" + "Amazon Refunds" + "\"")
typ, data = mail.search(None,'(SUBJECT "Refund Initiated for Order ")')

i = len(data[0].split())
with open('gmail_data.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile, delimiter = ' ')
    for x in range(i):
        latest_email_uid = data[0].split()[x]
        result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = email_data[0][1]
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
        newSubject = subject.strip("Refund Initiated for Order ")

        # just some visual feedback in console that it's working
        print(newSubject + " added!")

        writer.writerow(newSubject.split(" "))
csvfile.close()      
mail.close()
mail.logout()
