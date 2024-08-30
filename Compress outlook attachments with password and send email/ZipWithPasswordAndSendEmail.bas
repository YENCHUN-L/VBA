Sub ZipWithPasswordAndSendEmail()
    On Error GoTo ErrorHandler
    
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim olNewMail As Outlook.MailItem
    Dim olAttachment As Outlook.Attachment
    Dim strFilePath As String
    Dim strZipPath As String
    Dim strCommand As String
    Dim objShell As Object
    Dim fso As Object
    Dim i As Integer
    Dim strPassword As String
    Dim myEmail As String

    ' Initialize Outlook application
    Set olApp = Outlook.Application
    ' Get the current email being composed
    Set olMail = olApp.ActiveInspector.CurrentItem
    
    ' Check if there are any attachments
    If olMail.Attachments.Count = 0 Then
        ' Send the email directly if there are no attachments
        olMail.Send
        MsgBox "Email sent directly as there are no attachments."
        Exit Sub
    End If
    
    ' Define the path for saving attachments and the zip file
    strFilePath = "C:\Temp\Attachments\"
    strZipPath = "C:\Temp\Attachments\Attachments.zip"
    
    ' Create the folder if it doesn't exist
    If Dir(strFilePath, vbDirectory) = "" Then
        MkDir strFilePath
    End If
    
    ' Save all attachments to the specified folder
    For i = 1 To olMail.Attachments.Count
        Set olAttachment = olMail.Attachments(i)
        olAttachment.SaveAsFile strFilePath & olAttachment.FileName
    Next i
    
    ' Generate a random 8-character password
    strPassword = GeneratePassword(8)
    
    ' Create the zip file using WinZip command line with password
    strCommand = """C:\Program Files\WinZip\winzip64.exe"" -a -s""" & strPassword & """ """ & strZipPath & """ """ & strFilePath & "*.*"""
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run strCommand, 1, True
    
    ' Remove all attachments from the email
    For i = olMail.Attachments.Count To 1 Step -1
        olMail.Attachments.Remove i
    Next i
    
    ' Attach the zip file to the current email
    olMail.Attachments.Add strZipPath
    
    ' Send an additional email with the password to yourself
    myEmail = olApp.Session.CurrentUser.Address
    Set olNewMail = olApp.CreateItem(olMailItem)
    With olNewMail
        .To = myEmail
        .Subject = "Password for the Zip File"
        .Body = "The password for the zip file is: " & strPassword
        .Send
    End With
    
    ' Send the original email
    olMail.Send

    ' Delete all files in the temp folder
    Kill strFilePath & "*.*"

    ' Clean up
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olNewMail = Nothing
    Set olApp = Nothing
    Set objShell = Nothing
    Set fso = Nothing
    
    MsgBox "Attachments have been zipped, replaced with the zip file, emails have been sent, and the temp folder has been cleaned up successfully!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Function GeneratePassword(length As Integer) As String
    Dim chars As String
    Dim i As Integer
    Dim password As String
    
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()"
    Randomize
    For i = 1 To length
        password = password & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next i
    GeneratePassword = password
End Function

