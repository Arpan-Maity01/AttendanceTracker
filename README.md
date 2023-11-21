# AttendanceTracker
Libraries Used:

openpyxl: Used for working with Excel files.
smtplib: Used for sending emails using the Simple Mail Transfer Protocol (SMTP).
email.mime.multipart: Provides tools for creating MIME (Multipurpose Internet Mail Extensions) documents, a standard format for email messages.
email.mime.text: Used for creating the text part of the email.
os: Used for interacting with the operating system.
time: Used for introducing delays in the script.
Loading Excel Workbook:

The script begins by loading an Excel workbook named 'attendance.xlsx' using the openpyxl library.
Global Variables:

total_students: Keeps track of the total number of students.
absent_students: Keeps track of the number of absent students.
attendance_threshold: Represents the minimum attendance ratio required before sending an alert.
email_list: List to store email addresses.
Functions:

save_excel(): Saves the updated Excel sheet.
track_attendance(): Calculates the attendance ratio and sends email alerts if the ratio is below the specified threshold.
send_email(to_email): Sends an email alert to the specified email address.
Email Configuration:

The script uses a Gmail SMTP server (smtp.gmail.com) with port 587.
The sender's email address and password are hardcoded, which is not a recommended practice. It's better to use environment variables or a configuration file for sensitive information.
Main Loop:

The script enters an infinite loop using while True.
Inside the loop, it continuously tracks attendance and sends email alerts if necessary.
After each iteration, it waits for 1 hour (time.sleep(3600)) before checking attendance again.
Note: Make sure to replace placeholder values like 'sender_email@gmail.com' and 'password' with your actual email credentials. Also, consider using more secure methods for storing sensitive information, such as environment variables or a configuration file.






