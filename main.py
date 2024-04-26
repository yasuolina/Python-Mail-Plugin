import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from zipfile import ZipFile
import os


def send_email(zip_file, to, subject, body):
    msg = MIMEMultipart()
    msg['From'] = "example_email"
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(body))

    with open(zip_file, "rb") as f:
        part = MIMEApplication(f.read(), subtype="zip")
        part.add_header('content-disposition', 'attachment', filename=zip_file)
        msg.attach(part)

    smtp = smtplib.SMTP('smtp.office365.com', 587)
    smtp.starttls()
    sender_password = input("Please enter the password of sender : ")
    smtp.login("ysmn108@hotmail.com", sender_password)
    smtp.sendmail("ysmn108@hotmail.com", to, msg.as_string())
    smtp.quit()
    print('Mail sent successfully.')


def make_zip_file(path, file_name, file_type):
    print("File Zip Operation is processing..")
    # Create a ZipFile Object
    with ZipFile(path + file_name + '.zip', 'w') as zip_object:
        # Adding files that need to be zipped
        zip_object.write(path + file_name + file_type)
        print("ZIP file created")


def check_zip_file(path, file_name):
    # Check to see if the zip file is created
    if os.path.exists(path + file_name + '.zip'):
        flag = True
    else:
        flag = False
        print("ZIP file not created")
    return flag


def main():
    print("WELCOME TO PYTHON E-MAIL SENDING SERVICE\n")
    path = input("Please enter the path of file that you want to zip : ")
    file_name = input("Please enter the name of file that you want to zip : ")
    file_type = input("Please enter the type of file that you want to zip (.docx , .txt ..) : ")
    make_zip_file(path, file_name, file_type)
    flag = check_zip_file(path, file_name)
    if flag:
        mail_content = '''Hello,
                   This is a test mail.
                   In this mail we are sending some attachments.
                   The mail is sent using Python SMTP library.
                   Thank You
                   '''
        mail_subject = "TRIAL PYTHON E-MAIL"
        sender_mail = "your_email"
        send_email(path + file_name + '.zip', sender_mail, mail_subject,
                   mail_content)
    else:
        print("Unsuccessful Operation")


if __name__ == "__main__":
    main()

