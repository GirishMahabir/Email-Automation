# BULK MAIL SENDING.

- Normally for teachers who save their student's details and marks in an excel sheet.
- This script automates the sending of each student mark individually.
- Each mail is secured, using TLS and SSL.

## Not so mature yet, we have some constants: 
- The sender mail should be OUTLOOK.


## Excel Sheet template:
| Name | Surname | Email | Mark | \.\.\. |
|------|---------|-------|------|--------|
|      |         |       |      |        |
- The First 4 columns must be as above.
- After the column mark, you can put whatever you want.

## Usage:

- You can edit the script's msg_built function to customize your static message. (Play safe!)
- It is simple HTML, you can try to add pictures also(experimental feature!)
- Try send yourself an email to test the look.

### On Windows:
Step 1: Open Microsoft store and install python 3.7 or above.

Step 2: Open PowerShell and type these commands below.
```commandline
$ pip install openpyxl (Dependency)
```
Step 3: Run the script from PowerShell or cmd:
```commandline
$ python automail.py
```
- 1st prompt: name or path of the .xlsx file. 
(If having a problem with path try put the script in the same folder as the .xlsx file and instead of the path just put the name.)
- 2nd prompt: Your outlook email address.
- 3rd prompt: The custom message you want to put after the Hello <Student name,> and above the mark tab.
- 4th prompt: the password of your email.

### On Linux:
This script was first made for windows users so some features like terminal clear aren't functional yet.
Will be fixed in the next release.
```commandline
$ pip install openpyxl (Dependency)
$ python automail.py
```
- 1st prompt: Name of .xlsx(If script and excel file are in the same folder) else path of the .xlsx file. 
(If having a problem with path try put the script in the same folder as the .xlsx file and instead of the path just put the name.)
- 2nd prompt: Your outlook email address.
- 3rd prompt: The custom message you want to put after the Hello <Student name,> and above the mark tab.
- 4th prompt: the password of your email.
 