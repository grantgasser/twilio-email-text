import win32com.client
from twilio.rest import Client


#### Documentation on win32com.cleint reading emails ###
##  https://stackoverflow.com/questions/22813814/clearly-documented-reading-of-emails-functionality-with-python-win32com-outlook ##




#Namespace - The object itself provides methods for logging in and out, accessing storage objects directly by ID,
#accessing certain special default folders directly, and accessing data sources owned by other users.

#Use GetNameSpace ("MAPI") to return the Outlook NameSpace object from the Application object.
outlook = win32com.client.Dispatch("Outlook.Application")
outlook_ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


#email addr is root folder
'''
folder = outlook_ns.Folders['grant_gasser@baylor.edu'].Folders['Inbox'].Folders['Test Python Email']

messages = folder.Items

m_count = 0

for msg in messages:
    #print(msg.subject)
    #print(msg.sender)

    if str(msg.sender) == '201910 MIS 4V98 01 - Introduction to Python':
        print('You have an email from', msg.sender)

        msg2 = outlook.CreateItem(0)
        msg2.Importance = 2
        msg2.Subject = 'Got your ' + msg.subject + ' email'
        msg2.HTMLBody = 'Hi ' + str(msg.sender) + ' I am sending you an email through Python'

        msg2.To = msg.sender.address
        msg2.BCC = 'grant_gasser@baylor.edu'

        msg2.ReadReceiptRequested = True

        msg2.Send()

    m_count += 1

print('You have', m_count, 'messages in', folder)
'''

#Count unread messages
folder = outlook_ns.Folders['grant_gasser@baylor.edu'].Folders['Inbox']

messages = folder.Items

m_count = 0

for msg in messages:

    if msg.UnRead:
        m_count += 1

print('You have', m_count, 'unread messages')








##########  SEND TEXT MESSAGE ##############
########## CHECK TWILIO ACCT SET UP FIRST ##

accountSID = 'AC9fdc84024cdc0642f43ef58f98057255'

authToken = '3a03579680a3813e3803e6bda0190429'

client = Client(accountSID, authToken)

TwilioNumber = '+16235524120'

myCellPhone = "+16028267462"

msg = 'You have ' + str(m_count) + ' unread emails!'


#send text message
message = client.messages.create(
                     body=msg,
                     from_=TwilioNumber,
                     to=myCellPhone
                 )

print(message.sid)



#make a phone call
call = client.calls.create(url='http://demo.twilio.com/docs/voice.xml',
                    to=myCellPhone, from_=TwilioNumber)
