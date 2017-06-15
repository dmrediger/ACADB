'Function to export Excel Page as .txt file for further processing. Includes error checking and custom message boxes.'

Sub Upload()

'Setting Variables'
Dim txtSht As Worksheet
Dim entrySht As Worksheet
Dim SavePath As String
Dim FileName As String
Dim errorCnt As Integer

Set txtSht = ActiveWorkbook.Worksheets("Text Render")
Set entrySht = ActiveWorkbook.Worksheets("Entry Input")
errorCnt = entrySht.Cells(2, 10)

SavePath = "\savepath"
FileName = "\filename.csv"
txtRngCnt = txtSht.UsedRange.Rows.Count

'An alternative path for saving directly to one's desktop'
'SavePath = "C:\Users\" & Environ$("Username") & "\Desktop"'

'Checks the Error Count and stops the process if an error is detected'
If errorCnt > 0 Then
If MsgBox("There are Format Errors with the entry." & Chr(10) & Chr(10) & "Make sure Error Count is ZERO before continuing.", vbOKOnly + vbCritical, "Format Error") = vbOK Then
Exit Sub
End If
End If

If MsgBox("This will save the entry data as filename.csv and upload it to the FTP Server for GL Input." & Chr(10) & Chr(10) & "This will overwrite any existing file of the same name on the server." & Chr(10) & Chr(10) & "Would you like to Proceed?", vbYesNo + vbExclamation + vbDefaultButton2, "Entry Upload Confirmation") = vbYes Then

'Primary loop that will delete all "Empty" cells'
Application.ScreenUpdating = False 'Disable Updating for efficiency"
While txtRngCnt > 1
    With txtSht
        If (.Cells(txtRngCnt, 1).Value = " ") Then
        .Cells(txtRngCnt, 1).EntireRow.Delete
        End If
    End With
    txtRngCnt = txtRngCnt - 1
Wend
Application.ScreenUpdating = True

'Saves the file to the specified file path and name'
Application.DisplayAlerts = False
txtSht.Activate
txtSht.SaveAs FileName:=SavePath & FileName, FileFormat:=xlCSV
ActiveWorkbook.Close
Application.DisplayAlerts = True
End If 'Ends the IF statement executed by choosing "Yes" in the opening dialog box'
End Sub
