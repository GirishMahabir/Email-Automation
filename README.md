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


| Name  | Surname | Address                  | Email                    | Assignment 1  | Assignment 2 | Mark      |
|-------|---------|--------------------------|--------------------------|---------------|--------------|-----------|
| Aubin | Grimard | 21, Avenue Jean Portalis | aubaingrimard@dayrep.com | 60            | 80           | 90        |


The program will use the column that's says Mark. Assignment 1 and 2 will be ignored.

Address will also be ignored by the program.

## Usage:
- You can edit the script's msg_built function to customize your static message. (Play safe!)
- It is simple HTML, you can try to add pictures also(experimental feature!)
- Try send yourself an email to test the look.

### If on windows 64 bit:
You can just use the executable in the Packaged Folder.

#### How we packaged the app(You can re-package it yourself for any other OS or for security purpose):
```commandline
$ pip install pyinstaller
```

Modified the python script to take the relative path of the background image that we've used in the app.
Where we had to put the path of our image in the GUI, we just pass this function
```python
def resource_path(relative_path):
    """
    Deals with external files that will be neaded after
    packaging.
    :param relative_path: path of file.
    :return: relative path.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


poly_logo_path = tk.PhotoImage(file=resource_path("some_back.png"))
```
In our spec file of pyinstaller we need to add some line and make some changes:
```specfile
a.datas += [('some_back.png', '.\\some_back.png', 'DATA')]

# Exe section:
console=false
```
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
The background image is needed to start the application.

## Fill in the GUI and start.
