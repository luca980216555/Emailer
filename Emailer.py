import os
import win32com.client


ol = win32com.client.Dispatch("outlook.application")


olmailitem = 0x0


recipients = {
    'recipient1@mail.com': {'name': 'Mr. Luca', 'fileName': 'file1.txt'},
    'recipient2@mail.com': {'name': 'Ms. Alice', 'fileName': 'file2.txt'},
    'recipient3@mail.com': {'name': 'Mr. Bob', 'fileName': 'file3.txt'}
}


def get_signature():
    # Create a new mail item to grab the default signature
    temp_mail = ol.CreateItem(olmailitem)


    # Get the user's default signature (if any)
    temp_mail.Display()
    signature = temp_mail.HTMLBody
    return signature


# Get the user's signature
signature = get_signature()


for recipient, details in recipients.items():
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Test email'
    newmail.To = recipient
    newmail.CC = 'recipient4@mail.com; recipient5@mail.com'
    
    newmail.HTMLBody = '''<p>Dear {},</p>
    
    <p>This is a test email </p>
    
    <p>Kind regards,</p>
    '''.format(details['name']) + signature


    attach='C:\\User\\Desktop\\New folder\\'+ details['fileName']


    if os.path.isfile(attach):
        newmail.Attachments.Add(attach)
        
        newmail.Send()
        
    else:
        print(f'File does not exist: {attach}')  