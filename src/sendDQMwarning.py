import win32com.client

profilename="Outlook2003"

outlook = win32com.client.Dispatch("Outlook.Application")

inboxEmails = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items

eIndex = 1
mail_content = ''
email = inboxEmails.GetLast()
# start to get the first mail

subject_content = email.subject
receivedTime_content = email.ReceivedTime
sender_content = email.SenderEmailAddress
body_content = email.body[0:1000]

mail_content = mail_content \
               + '\r\n================= start email ' + str(eIndex) + ' ====================='\
               + '\r\n   subject:  ' + subject_content \
               + '\r\n   rec time: ' + str(receivedTime_content) \
               + '\r\n   sender:   ' + str(sender_content) \
               + '\r\n   content:  ' + body_content \
               + '\r\n++++++++++++++++++ end email ' + str(eIndex) + ' +++++++++++++++++++++'
# get email number start with 'caokejie'
eNoIndex = body_content.find('caokejie')

if eNoIndex == -1:
    eNumber = 3
else:
    eNumber = int(body_content[eNoIndex+8:eNoIndex+10])
    if (eNumber < 1) or (eNumber >= 100):
        eNumber = 3

#print eNumber

while(eIndex < eNumber):
    eIndex += 1
    email = inboxEmails.GetPrevious()
    subject_content = email.subject
    receivedTime_content = email.ReceivedTime
    sender_content = email.SenderEmailAddress
    body_content = email.body[0:1000]

    mail_content = mail_content \
                   + '\r\n================= start email ' + str(eIndex) + ' ====================='\
                   + '\r\n   subject:  ' + subject_content \
                   + '\r\n   rec time: ' + str(receivedTime_content) \
                   + '\r\n   sender:   ' + str(sender_content) \
                   + '\r\n   content:  ' + body_content \
                   + '\r\n++++++++++++++++++ end email ' + str(eIndex) + ' +++++++++++++++++++++'

#print mail_content

Msg = outlook.CreateItem(0)
Msg.To = "82716176@qq.com"
Msg.CC = "barry.cao@outlook.com"
#Msg.BCC = "address"
Msg.Subject = "get DQM email"
Msg.Body = mail_content

#attachment1 = "Path to attachment no. 1"
#attachment2 = "Path to attachment no. 2"
#Msg.Attachments.Add(attachment1)
#Msg.Attachments.Add(attachment2)
Msg.Send()