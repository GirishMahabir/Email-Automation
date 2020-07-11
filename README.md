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
#### Easy Way:
Step 1: Download executable file. (automail.exe)

Step 2: Run and give admin privilege. (For easy excel access install in same folder where your excel sheet is saved.)


Step 3: Browse to the folder and run the automail.exe script.

###### Now Follow on to the prompt section.

#### Challenge:
Step 1: Open Microsoft store and install python 3.7 or above.

Step 2: Open PowerShell and type these commands below.
```commandline
$ pip install openpyxl (Dependency)
```
Step 3: Run the script from PowerShell or cmd:
```commandline
$ python automail.py
```
###### Now Follow on to the prompt section.

### On Linux:
```commandline
$ pip install openpyxl
$ python automail.py
```
###### Now Follow on to the prompt section.
### Prompts Explained:
- 1st prompt: name or path of the .xlsx file. 
(If having a problem with path try put the script in the same folder as the .xlsx file and instead of the path just put the name.)
- 2nd prompt: Your outlook email address.
- 3rd prompt: Your Email subject.
- 4th prompt: The custom message you want to put after the Hello <Student name,> and above the mark tab.
- 5th prompt: The Assignment max mark.
- 6th prompt: the password of your email..