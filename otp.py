import win32com.client
import time

def mail_monitor_outlook(mailbox_address, subject_line, max_attempts=5, interval=5):
    '''Please login to the email in Outlook before making a call to this function'''
    
    mailbox_email = mailbox_address
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    other_mailbox = outlook.Folders.Item(mailbox_email)
    
    inbox = other_mailbox.Folders("Inbox")
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    
    attempts = 0
    while attempts < max_attempts:
        for message in messages:
            if message.Subject == subject_line:
                print("Sender: " + message.SenderEmailAddress)
                print("Subject: " + message.Subject)
                return message.Body
        attempts += 1
        time.sleep(interval)
    
    print("OTP not found after {} attempts.".format(max_attempts))
    return None

# Call the function with a maximum of 5 attempts and a 5-second interval
# message_body = mail_monitor_outlook(mailbox_address='yourmailhere', subject_line='Yoursubjectcodehere', max_attempts=5, interval=5)
# print(message_body)