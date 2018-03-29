#the contacts.txt file contains sample emails and names
import smtplib
from email.message import EmailMessage

def readTemplate(fileName):
    f = open(fileName, "r")
    template = f.read()
    f.close()
    return template

#modify this function to read in an ecel file, return the names and emails
#and any other relevant information to include in email body
def getContacts(fileName):
    f = open(fileName, "r")
    contactData = f.read().split("\n")
    names = []
    emails = []
    for line in contactData:
        names.append(line.split(" ")[0])
        emails.append(line.split(" ")[1])
    return names, emails

#progrsam is configured to use a queen's email address
YOUR_EMAIL = "insert sending email here"
PASS = "insert email password"
templateFileName = "template.txt" 
contactFileName = "contacts.txt" #change to exxcel file name 

s = smtplib.SMTP(host ="smtp.office365.com", port = 587) #change this if not using Queen's emails
s.starttls()
s.login(YOUR_EMAIL, PASS)

messageTemplate = readTemplate(templateFileName)

names, emails = getContacts(contactFileName)

# For each contact, send the email:
for name, email in zip(names, emails):
    msg = EmailMessage()       # create a message

    # add in the actual person name to the message template
    body = messageTemplate.replace("${PERSON_NAME}", name.title())

    #add message body
    msg.set_content(body)

    # setup the parameters of the message
    msg['From']=YOUR_EMAIL
    msg['To']=email
    msg['Subject']="This is TEST"

    # send the message via the server set up earlier.
    s.send_message(msg)
    
    del msg
    
s.quit()
