'Auto populate sheet from data using Excel dropdown control format object as the key. Update data and print when print_update button is pressed.' 

Private ws1 As Worksheet
Private ws2 As Worksheet
Private ws3 As Worksheet
Private wrkBk As Workbook
Private myArray(30) As String
Private locationArray(30) As String
Private keyVal As Integer


Sub ddKey_Change()

Set wrkBk = ThisWorkbook
Set ws1 = wrkBk.Worksheets("Wire Sheet")
Set ws2 = wrkBk.Worksheets("Wire Table")
Set ws3 = wrkBk.Worksheets("Location Key")
keyVal = ws1.Shapes("ddKey").ControlFormat.Value

Application.ScreenUpdating = False

For i = 0 To 30
    myArray(i) = ws2.Cells(keyVal + 1, i + 1)
    locationArray(i) = ws3.Cells(i + 2, 2)
Next i

For i = 1 To 30
    ws1.Range(locationArray(i)).Value = myArray(i)
Next i

Application.ScreenUpdating = True

End Sub

Sub print_update()

Set wrkBk = ThisWorkbook
Set ws1 = wrkBk.Worksheets("Wire Sheet")
Set ws2 = wrkBk.Worksheets("Wire Table")
Set ws3 = wrkBk.Worksheets("Location Key")
keyVal = ws1.Shapes("ddKey").ControlFormat.Value

ws1.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False

Application.ScreenUpdating = False

For i = 1 To 30
    locationArray(i) = ws3.Cells(i + 2, 2)
    myArray(i) = ws1.Range(locationArray(i))
    ws2.Cells(keyVal + 1, i + 1).Value = myArray(i)
Next i
     
Application.ScreenUpdating = True
    
End Sub
