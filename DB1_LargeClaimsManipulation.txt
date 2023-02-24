Sub LargeClaimsManipulation()

'varialbe declrations
Dim ICDCol As Integer
Dim claimCol As Integer
Dim testerRow As Integer
Dim endRow As Integer
Dim headerRow As Integer
Dim rng As Range
Dim LastCell As Long
Dim Add2 As Range

'prelim page set up
Application.DisplayAlerts = False
On Error Resume Next
Worksheets("Data").Delete
On Error GoTo 0
Application.DisplayAlerts = True
Worksheets(1).Name = "Raw"
Worksheets(1).Copy after:=Worksheets("Raw")
Worksheets(2).Name = "Data"
Worksheets("Data").Activate

'locate rows and cols needed for calculations
ICDCol = Worksheets("Data").UsedRange.Find("ICD DESCRIPTION").Column
claimCol = Worksheets("Data").UsedRange.Find("CLAIMANT TOTAL").Column
headerRow = Worksheets("Data").UsedRange.Find("CLAIMANT TOTAL").Row
testerRow = headerRow + 1
endRow = Worksheets("Data").Cells(Rows.count, 1).End(xlUp).Row

While testerRow <= endRow
    If Worksheets("Data").Cells(testerRow, claimCol) = 0 Then
        Worksheets("Data").Rows(testerRow).Delete
        endRow = endRow - 1
    ElseIf Worksheets("Data").Rows(testerRow).Font.Bold = False Then
        'if not bold, ctrl+down to the first empty row, which should be a bold total row
        boldRow = Worksheets("Data").Cells(testerRow, claimCol).End(xlDown).Row
        'populate the cols of the bold row with the cols of the row just above. THis is to track ICD Desc, and other info. CLaims should be at the end of the report, and this will not overwrite med/rx claims
        For fillCol = 1 To claimCol - 3
            Worksheets("data").Cells(boldRow, fillCol).Value = Worksheets("Data").Cells(boldRow - 1, fillCol)
        Next fillCol
        'odd definition here, need to ne sure we only delete the rows we need
        For deleteRows = 1 To boldRow - testerRow
        Worksheets("Data").Rows(testerRow).Delete
        endRow = endRow - 1
        Next deleteRows
    Else
        testerRow = testerRow + 1
    End If
Wend

'sort/filter/bold
Set rng = Range(Cells(headerRow, 1), Cells(endRow, claimCol))
Set Add2 = Cells.Find(What:="CLAIMANT TOTAL", LookAt:=xlWhole)
LastCell = Cells(Rows.count, Add2.Column).End(xlUp).Row
rng.Resize(LastCell).Sort key1:=Add2, order1:=xlDescending, header:=xlYes
Sheets("Data").UsedRange.Font.Bold = False

End Sub
