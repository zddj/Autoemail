import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

context = ssl.create_default_context()

# this is the file that the script will reference
details = 'email_details.xlsx'

# edit this with your login details
# make sure that account access to less secure apps are enabled 
sender = 'JohnDoe@email.com'
sender_pass = '123456789'

# edit this with your message and its subject
email_subject = 'subject'
body = '''Dear Mx. {surname} 

Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec fringilla efficitur mi nec dapibus. Vivamus sagittis nunc non ligula accumsan, in bibendum arcu ultricies. Vestibulum nec ullamcorper ipsum. Donec auctor varius dolor, vel finibus metus venenatis vitae. In vitae ante ut lectus blandit rhoncus. Etiam ultricies justo id auctor rhoncus. Aliquam sed maximus odio. Donec in accumsan lectus. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Sed at massa nec elit pellentesque fringilla viverra in quam. Ut semper pulvinar mauris interdum ullamcorper. Phasellus tincidunt ornare sem at suscipit. 

Aliquam eget metus vitae metus hendrerit porta vel nec nisl. Maecenas laoreet luctus pretium. Nulla condimentum felis lectus, in molestie massa dignissim ut. Nulla eu fringilla ex, in cursus nulla. Ut eu quam vitae quam mollis consectetur. Cras ligula metus, tincidunt eu nisl a, tincidunt aliquet arcu. Vivamus congue leo tincidunt, mattis ante nec, lobortis purus. Nam aliquam neque vitae neque bibendum accumsan. Curabitur imperdiet dignissim erat in volutpat. Integer id placerat nunc. Pellentesque non nulla et quam elementum finibus. Ut tellus ante, eleifend eget massa pharetra, interdum placerat urna. Mauris condimentum a est non mollis. Maecenas arcu massa, vehicula id augue eu, faucibus porttitor nisl. Aliquam finibus sapien vel est gravida, ut fringilla mi placerat.
'''

def attach_attachment(attachment_list, payload):
    # opening excel file into dataframe
    attachments = pd.read_excel(attachment_list, sheet_name = 'attachments')

    for file in attachments['attachments']:
        with open(file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())

        encoders.encode_base64(part) # encoding attachment
        part.add_header(
            'Content-Disposition',
            f'attachment; filename = {file}'
        )

        # attaching attachment
        payload.attach(part)      

# IF YOU DONT WANT TO SEND ATTACHMENTS
# remove the attachment_list argument if you don't want to include attachments
def composing_email(e_sender, e_password, e_subject, e_recipient, attachment_list):
    # setup the MIME
    message = MIMEMultipart()
    message['From'] = e_sender
    message['To'] = e_recipient
    message['Subject'] = e_subject

    # attaching body and the attachments for the mail
    message.attach(MIMEText(body, 'plain'))

    # comment this if you don't want to include attachments
    attach_attachment(attachment_list, message)

    try:
        # create SMTP session for sending the mail
        session = smtplib.SMTP('smtp.gmail.com', 587) # gmail port
        session.starttls(context = context) # securing connection
        session.login(e_sender, e_password) # logging into account

        text = message.as_string()

        # actually sending the email
        session.sendmail(e_sender, 
                        e_recipient, 
                        text.format(surname = surname, 
                                    receiver_email = recipient, 
                                    sender = e_sender,
                                    )
                        )
        session.quit()

        # confirmation
        print(f'Mail successfully sent to {surname} with email {recipient}')

    except:
        session.quit()

        print(f'Failed to send email to {surname} with email {recipient}')

if __name__ == '__main__':
    surname_email = pd.read_excel(details, sheet_name = 'surname_email')
    for surname, recipient in surname_email.itertuples(index=False):

        # remove the details arguement if you don't want to include attachments
        composing_email(sender, sender_pass, email_subject, recipient, details)