Sub makePDFs()

    ' Declare variables for worksheets
    Dim dataTab As Worksheet
    Dim templateTab As Worksheet
    Dim fieldsTab As Worksheet
    Dim activeTab As Worksheet
    
    ' Set references to the worksheets
    Set dataTab = Worksheets("Data")
    Set templateTab = Worksheets("Template")
    Set fieldsTab = Worksheets("Fields")
    Set activeTab = Worksheets("Active")
    
    ' Define the range to replace in the template
    Dim targetRangeToReplace As String
    targetRangeToReplace = "G14:G17"
    Dim replaceRange As Range
    Set replaceRange = activeTab.Range(targetRangeToReplace)
    Dim templateRange As Range
    Set templateRange = templateTab.Range(targetRangeToReplace)
    
    ' Define column numbers for file name and file path
    Dim colFileName As Integer
    colFileName = 35
    Dim colFilePath As Integer
    colFilePath = 36
    
    ' Variable for path to certificate folder
    Dim fullFilePath As String
    
    ' Get folder path and Python script path from Fields worksheet
    Dim folderPath As String
    Dim pythonScriptPath As String
    
    ' Ask user to enter the row number
    Dim r As Long
    r = InputBox("What row?", "Enter row")

    ' For each row of data in the data tab
    ' For r = 2 To dataTab.Cells(Rows.Count, 1).End(xlUp).Row
        ' If IsEmpty(ActiveSheet.Cells(r, 36).Value) = False Then
            ' Copy template and paste it on the active tab
            templateRange.Copy replaceRange
            ' For each variable in the fields tab
            For Each field In fieldsTab.Range(fieldsTab.Cells(2, 1), fieldsTab.Cells(fieldsTab.Cells(Rows.Count, 1).End(xlUp).Row, 1))
        
                ' If the field is {FolderPath}, get the folder path
                If (StrComp(field.Value, "{FolderPath}", vbTextCompare) = 0) Then
                    folderPath = field.Offset(0, 1).Value
                ' If the field is {PythonScriptPath}, get the Python script path
                ElseIf (StrComp(field.Value, "{PythonScriptPath}", vbTextCompare) = 0) Then
                    pythonScriptPath = field.Offset(0, 1).Value
                Else
                    ' For each {}, replace with corresponding data
                    For Each cell In replaceRange
                        cell.Value = Replace(cell.Value, field.Value, dataTab.Cells(r, field.Offset(0, 1).Value).Value)
                    Next cell
                End If
        
            Next field
            
            ' Reassign the fullFilePath to the certificate folder
            fullFilePath = folderPath & dataTab.Cells(r, colFileName).Value
            ' Encrypt the file
            activeTab.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            fullFilePath _
            , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=False
            
            ' Call subroutine to execute Python script to protect with password
            GeneratePDF fullFilePath, encryptedPath, pythonScriptPath
            
            ' Update file path in the data tab
            dataTab.Cells(r, colFilePath).Value = encryptedPath
            
            ' Create Outlook email for the learner
            Dim EmailApp As Outlook.Application
            Dim EmailItem As Outlook.MailItem
            Set EmailApp = New Outlook.Application
            Set EmailItem = EmailApp.CreateItem(olMailItem)
            ' Set email properties
            EmailItem.To = dataTab.Cells(r, 4).Value ' Learner email
            EmailItem.CC = dataTab.Cells(r, 27).Value & ";" & dataTab.Cells(r, 28).Value & ";" & dataTab.Cells(r, 32).Value ' CEA email, CEA CC, and DR email(s)
            EmailItem.Subject = "Certificate for the Homeowner ATU Online Program"
            ' Set email body
            EmailItem.HTMLBody = "Dear " & dataTab.Cells(r, 24).Value & "," & "<br>" & "<br>" & _
            "Congratulations for completing the training ""Homeowner Maintenance of Aerobic Treatment Units"". " & _
            "We have processed your certificate and sent it to your County Extension Office. Please allow 2-3 business " & _
            "days from this email and then schedule an appointment with your County Extension Agent (" & dataTab.Cells(r, 30).Value & _
            ", Agriculture and Natural Resources Program Area) to pick up your certificate. Remember to take with you a valid " & _
            "form of photo identification." & "<br>" & "<br>" & "Feel free to contact me if you have any questions." & "<br>" & "<br>" & _
            "Best Regards," & "<br>" & _
            "Gabriele Bonaiti" & "<br>" & _
            "p: +1 (979) 862-2593 | c: +1 (979) 922-4991 | f: +1 (979) 862-3442"
            EmailItem.Display
            Set EmailItem = Nothing
            
            ' Create Outlook email for CEA
            Dim Email2 As Outlook.MailItem
            Set Email2 = EmailApp.CreateItem(olMailItem)
            ' Set email properties
            Email2.To = dataTab.Cells(r, 27).Value ' CEA email
            Email2.CC = dataTab.Cells(r, 28).Value ' CEA email CC
            Email2.Subject = "Certificate for the Homeowner ATU Online Program"
            ' Set email body
            Email2.HTMLBody = "Dear " & dataTab.Cells(r, 26).Value & "," & "<br>" & "<br>" & _
            "Please find attached certificate for the class ""Homeowner Maintenance of Aerobic Treatment Units"", " & _
            "which is not signed. Please sign it and have it ready for pick up by the learner at your office. " & _
            "We have also informed the learner and the local regulator via email about issuance of this certificate. " & _
            "<br>" & "<br>" & "Let me know if you have any questions." & "<br>" & "<br>" & _
            "Best Regards," & "<br>" & _
            "Gabriele Bonaiti" & "<br>" & _
            "p: +1 (979) 862-2593 | c: +1 (979) 922-4991 | f: +1 (979) 862-3442"
            ' Attach encrypted PDF
            Email2.Attachments.Add encryptedPath
            Email2.Display
            Set Email2 = Nothing
            
        ' End If
        
    ' Next r ' Uncomment to generate all the rows

End Sub

Sub GeneratePDF(inputPath, outputPath, pythonScriptPath)
    ' Call the Python script to password protect the PDF
    
    ' Declare variables
    Dim pythonPath As String
    Dim scriptPath As String
    Dim password As String
    
    ' Set paths and password
    pythonPath = "python.exe" ' Path to Python executable
    scriptPath = pythonScriptPath & "password_protect_pdf.py" 
    'inputPath = "last_name_first_name.pdf" ' Update the path to  input PDF file
    'outputPath = "encrypted_last_name_first_name.pdf" ' Update the path to  output PDF file
    password = "HO" ' Update the password to use for the PDF
    
    Dim shellArgs As String
    shellArgs = pythonPath & " " & scriptPath & " " & inputPath & " " & outputPath & " " & password
    
    pid = Shell(pythonPath & " " & scriptPath & " " & inputPath & " " & outputPath & " " & password, vbNormalFocus)
    
    'pid = Shell(pythonPath & " " & scriptPath, vbNormalFocus)
End Sub

