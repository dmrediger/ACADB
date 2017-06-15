'Using the Replace method to automatically update Workbook references to specific file names/dates'

Sub Update_Projections()

    Application.ScreenUpdating = False
    ActiveWindow.DisplayFormulas = True
    
    'Updates Forecast Cells to current day information'
    Cells.Replace What:="\CYEAR\[CDAY.xlsx]", Replacement:="\" + Cells.Range("C3") + "\[" + Cells.Range("E3") + ".xlsx]", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Updates the Accum. ECR to the previous days data - The previous day Closing Ledger must be added to this'
    Cells.Replace What:="\CYEAR\[PDAY AM CASH TOTALS.xlsx]", Replacement:="\" + Cells.Range("C3") + "\[" + Cells.Range("G3") + " AM CASH TOTALS.xlsx]", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    ActiveWorkbook.RefreshAll
    ActiveWindow.DisplayFormulas = False
    Application.ScreenUpdating = True
    
End Sub
