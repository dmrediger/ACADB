'Add worksheet cells into array, sort/format them and paste into separate worksheet'

Sub Sort_Debit()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim arr() As Double
    Dim sumAmount As Double
    Dim amountLength As Integer
    Set ws1 = ActiveWorkbook.Worksheets("Data Report")
    Set ws2 = ActiveWorkbook.Worksheets("Processor")
    Set ws3 = ActiveWorkbook.Worksheets("Sorted Transactions")
    
    Application.ScreenUpdating = False 'Disable Screen Update During Processing'
    
    ws2.Cells.Clear
    ws3.Select
    ws3.Range("A3:A500").Clear
    ws1.UsedRange.Copy
    ws2.Range("A1").PasteSpecial
    Selection.MergeCells = False
    Selection.WrapText = False
    
    ReDim Preserve arr(0) 'Initiate dynamic array for SameDay data'
    
    'Go through the amount column (Col J) in the "Processor" Sheet - inserting values into an array and adding values to sumAmount variable'
    Dim j As Integer
    For j = 5 To ws2.UsedRange.Rows.Count 'Start at ROW 5 because this is the first amount value'
        If ws2.Range("J" & j) <> "" Then
                arr(UBound(arr)) = ws2.Range("J" & j)
                sumAmount = sumAmount + arr(UBound(arr))
                ReDim Preserve arr(UBound(arr) + 1)
        End If
    Next j
    
    'Loop to insert array values into Amount Column (COL A) in "Sorted Transactions" Sheet'
    Dim i As Integer
    For i = 1 To UBound(arr) 'Start with i = 1 because final array value will be empty based on previous loop'
        ws3.Cells(i + 2, 1) = arr(UBound(arr) - i) 'Insert Upper Bound less i variable to insert values in reverse (Descending) order'
    Next i
    
    amountLength = UBound(arr) + 2 'Add 2 to account for top two Label rows'
    
    'Convert all of the cells in the range to a Number Format'
    With ws3.Range("A3:A" & amountLength)
        .Value = Evaluate(.Address & "*1")
        .Style = "Comma"
        With .Font
            .Name = ("Calibri")
            .Size = 14
        End With
    End With

    ws3.Cells(amountLength + 2, 1) = sumAmount 'Insert the sumAmount value into the Sheet'
    
    'Convert sumAmount value to Number Format and add double underline styling'
    With ws3.Cells(amountLength, 1)
        With .Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
        End With
    End With
    With ws3.Cells(amountLength + 2, 1)
        .Value = Evaluate(.Address & "*1")
        .Style = "Comma"
        With .Font
            .Name = ("Calibri")
            .Size = 14
            .Bold = True
        End With
    End With
     
    Application.ScreenUpdating = True 'Re-enable Screen Update After Processing'
    
End Sub
