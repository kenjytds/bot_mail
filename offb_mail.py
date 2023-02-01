import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
import xlrd

user = 'Your_mail'
# noinspection SpellCheckingInspection
pwd = 'Your_Password'


def read_template(filename):  # Function to read the mail template
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def main():
    message_template = read_template('test.txt')

    # SMTP Server Setup
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login(user, pwd)

    # Excel sheet reading
    workbook = xlrd.open_workbook("Book1.xlsx")
    worksheet = workbook.sheet_by_index(0)  # Read the first sheet no.0
    firstname = []  # Store data in lists
    lastname = []
    emails = []
    for i in range(1, 3):
        for j in range(2, 3):
            firstname.append(worksheet.cell_value(i, 0))
            lastname.append(worksheet.cell_value(i, 1))
            emails.append(worksheet.cell_value(i, 2))

    for name, email in zip(firstname, emails):
        message = message_template.substitute(nom_personne=name)  # Change variable of .txt file

        msg = MIMEMultipart()  # Create a message
        msg['FROM'] = 'your_sender_mail'
        msg['TO'] = email
        msg['SUBJECT'] = "Test"
        msg.attach(MIMEText(message, 'plain'))  # Body of the mail

        # Cmd to send message via SMTP Server set up earlier
        s.send_message(msg)
        del msg
        # Terminate the SMTP session and close the connection
        s.quit()


if __name__ == '__main__':
    main()
