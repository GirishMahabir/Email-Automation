# BULK MAIL SENDING.
## Situation:
Teachers save student's details and assignment marks in an excel sheet. 
At the end of the day when the teacher wants to send each student's mark individually, it gets bulky and time-consuming.

Using Python along with its GUI and multi-thread Libraries I proposed an app that reads from the excel file 
(The student name, surname, email-address and mark.) It then requires the teacher to prepare a static message subject
header along with some personal details about the teacher.

After that, everything is handled by the script.
The script automatically builds a custom message one by one for each student and then sends it to the student.

In version 1 of the Bulk Mailing App, we had a template that the user had to follow so as the script can know
which column has which data.

In version 2 of the Bulk Mailing App, you can have additional columns in any order as long as the 
column name of the required fields is set right.

We should also note that each mail is secured, using TLS and SSL.

## Limitations:
- The sender mail should be OUTLOOK.

- The Excel Sheet: The script is case sensitive.
    - Name of the student column should be: Name
    - The surname of the student column should be: Surname
    - Email of the student column should be: Email
    - Mark of the student column should be: Mark
    
## Excel Example:

+-------+---------+--------------------------+--------------------------+---------------+--------------+-----------+
| Name  | Surname | Address                  | Email                    | Assignment 1  | Assignment 2 | Mark      |
+-------+---------+--------------------------+--------------------------+---------------+--------------+-----------+
| Aubin | Grimard | 21, Avenue Jean Portalis | aubaingrimard@dayrep.com | 60            | 80           | | 90      |
+-------+---------+--------------------------+--------------------------+---------------+--------------+-----------+

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
### On Linux:
```commandline
$ pip install openpyxl
$ python automail.py
```

### Directly Use the executable.

## Fill in the GUI and start.

## PYINSTALLER:

