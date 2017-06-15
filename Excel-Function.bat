'Example of the creation of a custom Public Function to be used in Excel'

Public Function SUCCESSRATE(Payer As String, Year As Integer, Month As Integer)

Dim processed As Integer
Dim allReceived As Integer
Dim epConnection As Worksheet
Dim acaReporting As Workbook

Set acaReporting = Workbooks("ACA Process Reporting.xlsm")
Set epConnection = acaReporting.Sheets("epConnection")

If Payer = "CUMULATIVE" Then 'Pull Cumulative Success Rate for ALL Payers'
processed = Application.WorksheetFunction.SumIfs(epConnection.Range("epConnection[PROCESSED_COUNT]") _
            , epConnection.Range("epConnection[YEAR_RECEIVED]"), "=" & Year _
            , epConnection.Range("epConnection[MONTH_RECEIVED]"), "=" & Month _
            , epConnection.Range("epConnection[DEFAULT_STATUS]"), "=0")

allReceived = Application.WorksheetFunction.SumIfs(epConnection.Range("epConnection[PROCESSED_COUNT]") _
            , epConnection.Range("epConnection[YEAR_RECEIVED]"), "=" & Year _
            , epConnection.Range("epConnection[MONTH_RECEIVED]"), "=" & Month)
Else 'Pull Success Rate by Payer Name'
processed = Application.WorksheetFunction.SumIfs(epConnection.Range("epConnection[PROCESSED_COUNT]") _
            , epConnection.Range("epConnection[PAYER_NAME]"), "=" & Payer _
            , epConnection.Range("epConnection[YEAR_RECEIVED]"), "=" & Year _
            , epConnection.Range("epConnection[MONTH_RECEIVED]"), "=" & Month _
            , epConnection.Range("epConnection[DEFAULT_STATUS]"), "=0")

allReceived = Application.WorksheetFunction.SumIfs(epConnection.Range("epConnection[PROCESSED_COUNT]") _
            , epConnection.Range("epConnection[PAYER_NAME]"), "=" & Payer _
            , epConnection.Range("epConnection[YEAR_RECEIVED]"), "=" & Year _
            , epConnection.Range("epConnection[MONTH_RECEIVED]"), "=" & Month)
End If

SUCCESSRATE = processed / allReceived

End Function
