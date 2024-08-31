
# Auto send compress outlook attachments with password
## This macro will compress Outlook attachments with passwords and send the email.
1. Check if there are any attachments, if not send email directly.

2. If there are any attachments, will replace the attachments as compress winzip
  with 8 digit password(at least one capital letter, one non-capital letter, one number, one punctutation)
  , and last sending out the email.


## To add this macro to Outlook follow these steps
1. Activate Developer Tab
  Open Outlook -> File -> Options -> Customize Ribbon -> Activate Developer Tab

2. Add macro to Outlook
  Developer Tab -> Visual Basic -> import ZipWithPasswordAndSendEmail.bas file
  or
  Developer Tab -> Visual Basic -> insert module -> copy paste the code to the module

3. You will find the macro at
  Devloper Tab -> Macros -> ZipWithPasswordAndSendEmail

Start sending compress Outlook attachments with passwords!!!
