'VBA for Excel tool designed to parse, review, and summarize data for input into Access Database'

Sub Import_Text()

'Declare iVars'
Dim myFile
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim wsRecycle As Worksheet
Dim wsProcess As Worksheet
Dim wsSameDayData As Worksheet
Dim wsSnapshot As Worksheet
Dim arr() As Variant
Dim arrPos As Long
Dim payerTag As String
Set ws1 = ActiveWorkbook.Worksheets("Data Import")
Set ws2 = ActiveWorkbook.Worksheets("Processor")
Set wsRecycle = ActiveWorkbook.Worksheets("RecycledPayments")
Set wsProcess = ActiveWorkbook.Worksheets("ProcessedPayments")
Set wsSameDayData = ActiveWorkbook.Worksheets("SameDayPayments")
Set wsSnapshot = ActiveWorkbook.Worksheets("Snapshot")

ws1.Cells.ClearContents 'Clears the Data Input Sheet at the Start'
wsRecycle.Range("A3:F" & wsRecycle.UsedRange.Count).ClearContents 'Clear the RecyclePayments Sheet at Start'
wsProcess.Range("A3:F" & wsProcess.UsedRange.Count).ClearContents 'Clear the ProcessPayments Sheet at Start'
wsSameDayData.Range("A4:BZ" & wsProcess.UsedRange.Count).ClearContents 'Clear the SameDayPayments Sheet at Start'
wsSameDayData.Range("F3:BZ3").ClearContents 'Clear the SameDayPayments Sheet at PAYERNAME Headers at start'

myFile = Application.GetOpenFilename() 'Opens the system dialog box to select a file'
Application.ScreenUpdating = False 'Disables Screen Updating while Macro runs'

startTime = TimeValue(Now) 'Time value reporting for processing optimization'

'This loop is responsible for pulling the entire text file into one text String and then pulling out the Same Day processed data'
Dim text As String, textline As String, posPSD As Double 'PSD stands for Processed Same Day'
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    text = text & textline
Loop
Close #1
posPSD = InStr(text, "820 FILES PROCESSED")
wsSameDayData.Range("B4").Value = Mid(text, posPSD + 50, 16)
wsSameDayData.Range("C4").Value = Mid(text, posPSD + 66, 14)
wsSameDayData.Range("D4").Value = Mid(text, posPSD + 80, 10)
wsSameDayData.Range("E4").Value = Mid(text, posPSD + 90, 14)


'Programically adds the .txt file into excel using a specific fixed width spacing that allows for optimal data recognition'
    With ws2.QueryTables.Add(Connection:= _
        "TEXT;" & myFile, Destination:=ws2.Range("$A$2") _
        )
        .Name = "DocDir_Export"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 2, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = Array(18, 15, 27, 24, 11, 17)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    'Copy and Paste values from hidden Processor sheet to the Data Import sheet for data parsing - this allows manipulation of data without .txt connection'
    ws2.UsedRange.Copy
    ws1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    ws2.Cells.ClearContents 'Clear Processor Sheet after Copy'
    
    'Declare Variables for data processing'
    rwCntD = ws1.UsedRange.Rows.Count 'Data Input Used Row Count'
    rwCntR = wsRecycle.UsedRange.Rows.Count 'Recycled Sheet Used Row Count'
    rwCntP = wsProcess.UsedRange.Rows.Count 'Processed Sheet Used Row Count'
    TypeTag = "" 'PROCESSED or RECYCLED'
    payerTag = "" 'Tag of the latest Payer Name'
    recycleCount = 0 'Pull value from Recycle Count Field to populate this variable'
    posD = 1 'Row Position in Data Import Worksheet'
    posR = 3 'Row Position in Recycled Payments Worksheet'
    posP = 3 'Row Position in Processed & Default Payments Worksheet'
    pAmount1 = 0 'Processed Amount Con 1'
    pAmount2 = 0 'Processed Amount Con 2'
    rAmount1 = 0 'Recycled Amount Con 1'
    rAmount2 = 0 'Recycled Amount Con 2'
    accName = "Con1" 'Variable for Concentration Account Identifier'
    ReDim Preserve arr(2, 0) As Variant 'Initiate dynamic array for SameDay data'
    
    '<<BEGINNING OF PRIMARY LOOP>>'
    While posD <= rwCntD

        'If Column 1 is blank, view the value in Column 4 to determine if it's an indicator or TypeTag (PROCESSED, RECYCLED, DEFAULTED)'
        If ws1.Cells(posD, 1) = "" Then
            If ws1.Cells(posD, 4) = "ROCESSED CURRENT DAY" Then
                TypeTag = "PROCESSED"
            ElseIf ws1.Cells(posD, 4) = "ECKS RECYCLED" Then
                TypeTag = "RECYCLED"
            ElseIf ws1.Cells(posD, 4) = "ECKS DEFAULTED" Then
                TypeTag = "DEFAULTED"
            End If

        'If this cell has value "11702 CORPORATE" then subsequent cell in column 4 will have an account number which will be used to determine accName Value (Con1 OR Con2)'
        ElseIf ws1.Cells(posD, 1) = "11702 CORPORATE" Then
            accName = Switch(ws1.Cells(posD, 4) = "4128175346", "Con1", ws1.Cells(posD, 4) = "4128185220", "Con2")
        
        'Primary IF with internal looping to handle RECYCLE Day 0 and PROCESSED Day > 0 - These rows will be copied and pasted to corresponding sheet _
        RecycledPayments OR ProcessedPayments where they will be stored for upload to Access DB. The internal loops look for matched payerTag and recycleCount _
        and will copy all subsequent matching rows to substantially increase processing speed and efficiency'
        ElseIf ws1.Cells(posD, 1) = "820" Then
            If payerTag <> ws1.Cells(posD, 2) Then payerTag = ws1.Cells(posD, 2)
            If recycleCount <> ws1.Cells(posD, 6) Then recycleCount = ws1.Cells(posD, 6)
            
            nextRow = "MATCH"
            matchRowCnt = 0
            
            If TypeTag = "RECYCLED" And recycleCount = 0 Then
                While nextRow = "MATCH"
                    If ws1.Cells(posD + matchRowCnt + 1, 1) = "820" _
                    And ws1.Cells(posD + matchRowCnt + 1, 2) = payerTag _
                    And ws1.Cells(posD + matchRowCnt + 1, 6) = recycleCount Then
                        matchRowCnt = matchRowCnt + 1
                    Else
                        ws1.Range(ws1.Cells(posD, 1), ws1.Cells(posD + matchRowCnt, 6)).Copy
                        wsRecycle.Cells(posR, 1).PasteSpecial xlPasteValues
                        posR = posR + matchRowCnt + 1
                        nextRow = "NO MATCH"
                        posD = posD + matchRowCnt + 1
                    End If
                Wend
                
            ElseIf TypeTag = "PROCESSED" And recycleCount > 0 Then
                While nextRow = "MATCH"
                    If ws1.Cells(posD + matchRowCnt + 1, 1) = "820" _
                    And ws1.Cells(posD + matchRowCnt + 1, 2) = payerTag _
                    And ws1.Cells(posD + matchRowCnt + 1, 6) = recycleCount Then
                        matchRowCnt = matchRowCnt + 1
                    Else
                        ws1.Range(ws1.Cells(posD, 1), ws1.Cells(posD + matchRowCnt, 6)).Copy
                        wsProcess.Cells(posP, 1).PasteSpecial xlPasteValues
                        posP = posP + matchRowCnt + 1
                        nextRow = "NO MATCH"
                        posD = posD + matchRowCnt + 1
                    End If
                Wend
            End If
            
        'If Column 1 value "TOTAL CHE" and Column 2 value "CK AMOUNT:" then subsequent cell in column 3 has a total amount that either represents _
        Sameday RECYCLE/PROCESS or >0 day RECYCLE/PROCESS. This IF Statement block handles the summation of these TOTAL Amounts for post process reporting'
        ElseIf ws1.Cells(posD, 1) = "TOTAL CHE" And ws1.Cells(posD, 2) = "CK AMOUNT:" Then
            If TypeTag = "PROCESSED" Then
                If recycleCount = 0 Then
                    If IsInArray(payerTag, arr) = -1 Then 'This is the If function where the SameDay Array "arr" is built to store SameDay PayerTag and SUM AMNT & CNT'
                        ReDim Preserve arr(2, LBound(arr, 2) To UBound(arr, 2) + 1) As Variant
                        arr(0, UBound(arr, 2)) = payerTag
                        arr(1, UBound(arr, 2)) = arr(1, UBound(arr, 2)) + ws1.Cells(posD, 3)
                        arr(2, UBound(arr, 2)) = arr(2, UBound(arr, 2)) + ws1.Cells(posD - 1, 3)
                    ElseIf IsInArray(payerTag, arr) >= 0 Then
                        arrPos = IsInArray(payerTag, arr)
                        amntValCnvrsn = CDbl(arr(1, arrPos)) + ws1.Cells(posD, 3)
                        cntValCnvrsn = CDbl(arr(2, arrPos)) + ws1.Cells(posD - 1, 3)
                        arr(1, arrPos) = CStr(amntValCnvrsn)
                        arr(2, arrPos) = CStr(cntValCnvrsn)
                    End If
                        
                'If typeTag = "PROCESSED" and recycle count is greater than 0 and less than or equal to 3 then add it to pAmount for corresponding account'
                ElseIf recycleCount > 0 And recycleCount <= 3 Then
                    If accName = "Con1" Then
                        pAmount1 = pAmount1 + ws1.Cells(posD, 3)
                    ElseIf accName = "Con2" Then
                        pAmount2 = pAmount2 + ws1.Cells(posD, 3)
                    End If
                End If
            'If typeTag = "RECYCLED" and recycle count is equal to 0 then add values to rAmount for corresponding account'
            ElseIf TypeTag = "RECYCLED" And recycleCount = 0 Then
                If accName = "Con1" Then
                    rAmount1 = rAmount1 + ws1.Cells(posD, 3)
                ElseIf accName = "Con2" Then
                    rAmount2 = rAmount2 + ws1.Cells(posD, 3)
                End If
            End If
        End If
            
        posD = posD + 1
    Wend 'End of WHILE LOOP from Line 111'
    
    'For Loop to populate the SameDayPayments with SUM AMNT and CNT for Processed Day 0 payments'
    sdColPos = 6 'Hard code for first col of data'
    sdNameRow = 3 'Hard code for row for PayerTags'
    sdValRow = 4 'Hard code for row for AMNT and CNT values'
    wsSameDayData.Cells(sdValRow, 1) = Right(ws1.Cells(2, 7), 8) 'Populate the data DATE'
    For i = 1 To UBound(arr, 2) 'Start with i=1 instead of LBound, because array initated with no values in arr(dimension 0)'
        wsSameDayData.Cells(sdNameRow, sdColPos) = arr(0, i) & " AMNT"
        wsSameDayData.Cells(sdNameRow, sdColPos + 1) = arr(0, i) & " CNT"
        wsSameDayData.Cells(sdValRow, sdColPos) = arr(1, i)
        wsSameDayData.Cells(sdValRow, sdColPos + 1) = arr(2, i)
        sdColPos = sdColPos + 2
    Next i
    
    'Add processed data to Snapshot sheet'
    With wsSnapshot
    .Activate
    .Range("B6").Value = rAmount1
    .Range("B10").Value = rAmount2
    .Range("B7").Value = pAmount1 * -1
    .Range("B11").Value = pAmount2 * -1
    End With
    
    'This removes the connection to the .txt file that is initially established in this Macro to import data'
    Dim xConnect As Object
    For Each xConnect In ActiveWorkbook.Connections
    If xConnect.Name = "" Then xConnect.Delete
    Next xConnect
    
    'This imports all newly added RECYCLED, PROCESSED, and SAMEDAY data to the corresponding Table in the ACADB.accdb Access Database'
    Dim acc As New Access.Application
    acc.OpenCurrentDatabase "S:\Cash Management\ACA Database Files\ACADB.accdb"
    acc.DoCmd.TransferSpreadsheet _
            TransferType:=acImport, _
            SpreadSheetType:=acSpreadsheetTypeExcel12, _
            TableName:="SameDaySum", _
            Filename:=ActiveWorkbook.FullName, _
            HasFieldNames:=True, _
            Range:="SameDayPayments$A3:" & wsSameDayData.Cells(4, (UBound(arr, 2) * 2 + 5)).Address(False, False) 'Dynamic end range cell based on Array "arr" size - multiply by 2 for AMNT and CNT, and add 5 for date and total amnts/cnts cells'
    
    acc.DoCmd.TransferSpreadsheet _
            TransferType:=acImport, _
            SpreadSheetType:=acSpreadsheetTypeExcel12, _
            TableName:="Recycled", _
            Filename:=ActiveWorkbook.FullName, _
            HasFieldNames:=True, _
            Range:="RecycledPayments$A2:F" & (posR - 1)

    acc.DoCmd.TransferSpreadsheet _
            TransferType:=acImport, _
            SpreadSheetType:=acSpreadsheetTypeExcel12, _
            TableName:="Processed", _
            Filename:=ActiveWorkbook.FullName, _
            HasFieldNames:=True, _
            Range:="ProcessedPayments$A2:F" & (posP - 1)
            
    acc.CloseCurrentDatabase
    acc.Quit
    Set acc = Nothing
    
    'Retrieves the ending time value and processes a calculation to provide the user with time to run process information'
    endTime = TimeValue(Now)
    timeElapse = endTime - startTime
    wsSnapshot.Range("B13") = Left(timeElapse, 5) & " Seconds!"
    
    Application.ScreenUpdating = True 'After this portion of the macro is completed - turn screen updating back on'
    
    'Sends a message to the user confirming the data input and amounts'
    MsgBox "Data Import Complete!" & vbNewLine & "Total Recycled Con 1: " & Format(rAmount1, "Currency") & vbNewLine & "Total Recycled Con 2: " & Format(rAmount2, "Currency") _
    & vbNewLine & vbNewLine & "Total Processed Con 1: " & Format(pAmount1, "Currency") & vbNewLine & "Total Processed Con 2: " & Format(pAmount2, "Currency")
     
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Long
  Dim i As Long
  ' default return value if value not found in array
  IsInArray = -1

  For i = LBound(arr, 2) To UBound(arr, 2)
    If StrComp(stringToBeFound, arr(0, i), vbTextCompare) = 0 Then
      IsInArray = i
      Exit For
    End If
  Next i
End Function

