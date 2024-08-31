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
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'fso.DeleteFile strFilePath & "*.*", True
    
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
    Dim randIndex As Integer
    Dim requiredChars As String
    
    ' Character sets
    Dim capital As String
    Dim nonCapital As String
    Dim number As String
    Dim punctuation As String
    
    capital = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    nonCapital = "abcdefghijklmnopqrstuvwxyz"
    number = "0123456789"
    punctuation = "!@#$%^&*()"
    
    ' Ensure the password contains at least one of each required character type
    requiredChars = Mid(capital, Int((Len(capital) * Rnd) + 1), 1)
    requiredChars = requiredChars & Mid(nonCapital, Int((Len(nonCapital) * Rnd) + 1), 1)
    requiredChars = requiredChars & Mid(number, Int((Len(number) * Rnd) + 1), 1)
    requiredChars = requiredChars & Mid(punctuation, Int((Len(punctuation) * Rnd) + 1), 1)
    
    ' Combine all character sets
    chars = capital & nonCapital & number & punctuation
    
    ' Randomly fill the rest of the password
    For i = 1 To (length - 4)
        randIndex = Int((Len(chars) * Rnd) + 1)
        password = password & Mid(chars, randIndex, 1)
    Next i
    
    ' Add the required characters to the password
    password = password & requiredChars
    
    ' Shuffle the password to ensure randomness
    password = StrConv(password, vbUnicode)
    password = StrConv(password, vbFromUnicode)
    
    GeneratePassword = password
End Function

