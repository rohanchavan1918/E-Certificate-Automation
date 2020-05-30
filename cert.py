from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def send_mail(reciever_email,cert_file_name,name):
    mail_content = f'''Dear {name},

Greetings from SJCEM!

Thank you for attending the 5-day Webinar Series organized by Department of Computer Engineering, St. John College of Engineering and Management, Palghar.

Please find attached your E-certificate.

We hope that you enjoyed the sessions and would be interested in attending our future sessions too.

Please subscribe to our YouTube Channel (Link: https://www.youtube.com/channel/UCfPHjK7YAI77koFtmjhdd8Q/featured) for more updates.

Thank you!

Please feel free to contact the faculty coordinators for any queries. 

Best Regards,
Organizing Team
Department of Computer Engineering
St. John College of Engineering and Management, Palghar
Website: www.sjcem.edu.in
    '''
    #The mail addresses and password
    sender_address = 'rohanchavan@nimapinfotech.com'
    sender_pass = 'rohan@123'
    receiver_address = reciever_email
    file_header_name = name+'_Webinar_certificate'
    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = "SJCEM Webinar"
    message['To'] = reciever_email
    message['Subject'] = 'E-certificate of '+name+' for attending the Webinar Series at St. John College of Engineering and Management, Palghar from May 24 to 28, 2020'
    #The subject line
    #The body and the attachments for the mail
    message.attach(MIMEText(mail_content, 'plain'))
    attach_file_name = './iot/'+cert_file_name # CHANGE
    # attach_file_name = cert_file_name
    print("attaching",attach_file_name)
    attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
    payload = MIMEBase('application', 'octate-stream')
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload) #encode the attachment
    #add payload header with filename
    payload.add_header('Content-Disposition', 'attachment', filename=attach_file_name)
    message.attach(payload)
    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    print('Mail Sent to',reciever_email)


# Location of the xls file
loc = "pubg_squad.xlsx"

wb = xlrd.open_workbook (loc)
sheet = wb.sheet_by_index(0)


for i in range(0,sheet.nrows):
# for i in range(1,10):
    try:
        name = sheet.row_values(i)[0]
        print(name)
        # # designation = sheet.row_values(i)[1]
        # designation = "Bridge It"
        email = sheet.row_values(i)[1]
        # Image of certificate
        img = Image.open("cert.jpg") # CHANGE
        draw = ImageDraw.Draw(img)
        selectFont = ImageFont.truetype("font1.ttf", size = 100)
        draw.text( (1977,1151), name, (0,0,0), font=selectFont)
        # draw.text( (765,1595), designation, (0,0,0), font=selectFont)
        cert_file_name = name+'_certificate.pdf'
        img.save( './iot/'+cert_file_name, "PDF", resolution=100.0) #Change Folder
        # img.save( cert_file_name, "PDF", resolution=100.0)
        send_mail(email,cert_file_name,name)
    except Exception as e:
        print(e)
        print("Handled at ",1)
        pass

# 1977,1151

# Subject:  E-certificate of <<Participant Name>> for attending the Webinar Series at St. John College of Engineering and Management, Palghar from May 24 to 28, 2020
# Dear <<Participant Name>>
# Greetings from SJCEM!
# Thank you for attending the 5-day Webinar Series organized by Department of Computer Engineering, St. John College of Engineering and Management, Palghar.
# Please find attached your E-certificate.
# We hope that you enjoyed the sessions and would be interested in attending our future sessions too.
# Please subscribe to our YouTube Channel (Link: https://www.youtube.com/channel/UCfPHjK7YAI77koFtmjhdd8Q/featured) for more updates.
# Thank you!

# Please feel free to contact the faculty coordinators for any queries. 

# Best Regards,
# Organizing Team
# Department of Computer Engineering
# St. John College of Engineering and Management, Palghar
# Website: www.sjcem.edu.in
