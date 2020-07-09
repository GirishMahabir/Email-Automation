import smtplib, ssl, time, os
from openpyxl import load_workbook
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Email Configuration.
SMTP_SERVER = "smtp-mail.outlook.com"
PORT = 587 # for TLS on OUTLOOK.
SENDER_EMAIL = " "# Your email address will be updated.
PASSWORD = " " # Password for you email will be updated.

# Data Dictionary for further processing.
DATA = {'NAME':[],'SURNAME':[],'EMAIL':[],'MARK':[]}

def update_dict():
    """
    Generated Dictionary to work with.
    """
    row_len = SHEET_OBJ.max_row
    for rowIndex in range(2,row_len + 1):
        cell_obj = SHEET_OBJ.cell(row = rowIndex, column = 1)
        DATA['NAME'].append(cell_obj.value) 
    for rowIndex in range(2,row_len + 1):
        cell_obj = SHEET_OBJ.cell(row = rowIndex, column = 2)
        DATA['SURNAME'].append(cell_obj.value)
    for rowIndex in range(2,row_len + 1):
        cell_obj = SHEET_OBJ.cell(row = rowIndex, column = 3)
        DATA['EMAIL'].append(cell_obj.value)
    for rowIndex in range(2,row_len + 1):
        cell_obj = SHEET_OBJ.cell(row = rowIndex, column = 4)
        DATA['MARK'].append(cell_obj.value)

def msg_build(name,surname,mark,receiver_mail,header_message):
    message = MIMEMultipart("alternative")
    # Insert Your Email Subject here. 
    message["Subject"] = "Assignment Mark Submission"
    message["From"] = SENDER_EMAIL
    message["TO"] = receiver_mail
    # insert your html content below(The static message. Everything in curly brackets are variables that will be
    # completed by the program.)
    # Avoid using curly brackets part of your html, it will conflict here.
    html = (f"""\
<html>
  <body>
    <p>
      Hello {name},<br><br>
      {header_message}<br><br>
      &emsp;&emsp;&emsp;Name &emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;Total(100 Marks)<br>
      &emsp;&emsp;{surname} {name}&emsp;&emsp;&emsp;&emsp;&emsp;{mark}
    </p>

    <p>Thanks & Regards,<br><br>
       <strong>YOUR NAME here</strong><br>
       TITLE here!<br><br>
       <strong>Email:</strong> email address here!<br>
       <strong>Github:</strong>
       <a href="https://www.github.com/GirishMahabir">github.com/GirishMahabir</a><br>
       <strong>This was an automated python script.</strong>
    </p>
  </body>
</html>
""")

    #txt = MIMEText(text,"plain")
    htm = MIMEText(html, "html")
    #message.attach(txt)
    message.attach(htm)
    msg_ready = message.as_string()
    return msg_ready

def send_mail(receiver_mail, message):
    # Create a secure SSL context
    context = ssl.create_default_context()
    # Try to log in to server and send email
    try:
        server = smtplib.SMTP(SMTP_SERVER,PORT)
        # Identifying user on server.
        server.ehlo()
        server.starttls(context=context) # Secure the connection
        # Identifying user on server.
        server.ehlo()
        server.login(SENDER_EMAIL, PASSWORD)
        server.sendmail(SENDER_EMAIL, receiver_mail, message)
    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit()

def main():
    global SENDER_EMAIL, PASSWORD, FILE_PATH, SHEET_OBJ # Accessing GLOBAL VARIABLES.
    clear = lambda: os.system('cls')
    print("WELCOME TO OUR BATCH AUTOMATED EMAIL SCRIPT.\n")
    # Path of Excel File.
    print("WARNING: For the next Prompt:\nIF script and excel file are in the same folder you can just put the file name along with the correct extention.\n")
    FILE_PATH = input("Enter your .xlsx file PATH along with the correct extention: ")  # copy paste file name of your .xlsx file.
    # Workbook object is created to open.
    WB_OBJ = load_workbook(filename=FILE_PATH)
    # Sheet Object.
    SHEET_OBJ = WB_OBJ.active
    SENDER_EMAIL = input("Enter your email address: ")
    header_message = input("Enter Message you want to keep static above the mark table: ")
    PASSWORD = input(f"Please enter you password for {SENDER_EMAIL} HERE: ")
    clear()
    print("Cleared Terminal to hide sensitive information.\n[STARTING...]")
    update_dict()
    num = len(DATA['NAME'])
    for i in range(num):
        print(f"SENDING to {DATA['NAME'][i]}...")
        message_ready = msg_build(DATA['NAME'][i]\
            ,DATA['SURNAME'][i],DATA['MARK'][i],DATA["EMAIL"][i], header_message)
        send_mail(DATA['EMAIL'][i], message_ready)
        print(f"SENT TO {DATA['NAME'][i]}!")
        time.sleep(2)
    print("ALL Done, [EXITING]...")

if __name__=="__main__":
    main()
