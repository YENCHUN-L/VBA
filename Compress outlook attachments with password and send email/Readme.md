
# Compress Outlook Attachments With Password And Send Email

This project provides an Outlook VBA macro that sends the current draft email and automatically handles attachments as follows:

1. If there are no attachments, the email is sent immediately.
2. If attachments exist, they are saved to a temporary folder, zipped with a generated password, and replaced by one zip file.
3. A second email containing the password is sent to your own mailbox.
4. The original email is then sent.

## Important Security Note

This macro does not send the password to the recipient.

You must share the password with the recipient through a separate secure channel (for example: phone call, secure chat, or approved internal tool).

## Requirements

- Microsoft Outlook desktop (Windows) with VBA enabled.
- WinZip installed with command line executable available at:
  - C:\Program Files\WinZip\winzip64.exe
- Permission to create and delete files in:
  - C:\Temp\Attachments\

## Files In This Folder

- ZipWithPasswordAndSendEmail.bas: main VBA macro module.
- Readme.md: usage and setup documentation.
- Flowchart.drawio / Flowchart.drawio.pdf: process diagram.

## How It Works

When you run ZipWithPasswordAndSendEmail from a composed email window:

1. The macro checks the active email item.
2. All current attachments are saved to C:\Temp\Attachments\.
3. An 8-character password is generated (includes uppercase, lowercase, number, and punctuation).
4. WinZip creates C:\Temp\Attachments\Attachments.zip using that password.
5. Original attachments are removed from the email.
6. Attachments.zip is added to the email.
7. A new email is sent to your own address with the generated password.
8. The original email is sent.
9. Temporary files in C:\Temp\Attachments\ are deleted.

## Setup In Outlook

1. Enable the Developer tab:
   - Outlook -> File -> Options -> Customize Ribbon -> check Developer
2. Open VBA editor:
   - Developer -> Visual Basic
3. Import the module:
   - File -> Import File... -> select ZipWithPasswordAndSendEmail.bas
   - Or create a new module and paste the code manually
4. Save and restart Outlook if prompted by macro settings.
5. Run macro:
   - Open or compose an email
   - Developer -> Macros -> ZipWithPasswordAndSendEmail -> Run

## Macro Permissions

If the macro does not run, check Trust Center settings:

1. Outlook -> File -> Options -> Trust Center -> Trust Center Settings
2. Open Macro Settings
3. Choose your organization-approved option (for example, signed macros only)

Follow your company security policy before enabling macros.

## Configuration Points

If your environment differs, update these values in ZipWithPasswordAndSendEmail.bas:

- WinZip path:
  - "C:\Program Files\WinZip\winzip64.exe"
- Temporary folder path:
  - "C:\Temp\Attachments\"
- Zip output path:
  - "C:\Temp\Attachments\Attachments.zip"
- Password length:
  - GeneratePassword(8)

## Known Limitations

- The password email is sent to the current Outlook user only.
- Password delivery to recipients is manual.
- If WinZip is not installed at the expected path, zipping will fail.
- File names with duplicates may overwrite each other in the temp folder.
- The current password shuffle logic does not provide cryptographic-grade randomness.

## Troubleshooting

- Error: WinZip executable not found
  - Confirm WinZip is installed and update the path in the macro.
- Error: Permission denied for C:\Temp\Attachments\
  - Run Outlook with appropriate permissions or choose another writable folder.
- No active composed email
  - Run the macro from an open compose window.
- Macro blocked by policy
  - Check Trust Center and your IT policy for macro execution rules.

## Recommended Safe Usage

1. Draft email and attach files as usual.
2. Run the macro to replace attachments with password-protected zip.
3. Send password to recipient through a separate secure channel.
4. Keep password and recipient mapping in your approved audit method if required.

## Disclaimer

Use this macro according to your organization's security, retention, and compliance rules. Test in a non-production mailbox before broad rollout.
