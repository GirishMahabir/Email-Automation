# BULK MAIL SENDING.
## Situation:
Teachers save student's details and assignment marks in an excel sheet. 
At the end of the day when the teacher wants to send each student's mark individually, it gets bulky and time-consuming.

Using Python along with its GUI and multi-thread Libraries I proposed an app that reads from the excel file 
(The student name, surname, email-address and mark.) It then requires the teacher to prepare a static message subject
header along with some personal details about the teacher.

After that, everything is handled by the script.
The script automatically builds a custom message one by one for each student and then sends it to the student.

We should also note that each mail is secured, using TLS and SSL.

## Limitations:
- The sender mail should be OUTLOOK.

In version 1 of the Bulk Mailing App, we had a template that the user had to follow so as the script can know
which column has which data.

In version 2 of the Bulk Mailing App, you can have additional columns in any order as long as the 
column name of the required fields is set right and the v2 also comes with a full GUI.