Sub DB2_Claims()
'Modifies the report from carrier, this macro alters the original file, eventually outputting
'a list of claims per type per month, and claimants who exceed a specifed threshold, to be entered by the user
'on the first worksheet in cell B1
Dim lastDataCol As Long
Dim lastDataRow As Long
Dim rangeHPC As Range
Dim rangeLC  As Range
Dim cacheHPC As PivotCache
Dim cacheLC As PivotCache
Dim pivotTableHPC As PivotTable
Dim pivoTableLC As PivotTable
Dim PivTbl As PivotTable
Dim rng As Range
Dim cell As Range

'This IF statement prevents the macro from being run on an previously modified sheet
If Sheets(1).Name = "Criteria" Then
    Exit Sub
    
Else

'Disables screen updating and alerts to expedite macro speed, then renames/deletes sheets as needed
Application.ScreenUpdating = False
On Error Resume Next
Application.DisplayAlerts = False
Sheets(1).Name = "Criteria"
Sheets(2).Name = "Raw"
Sheets(3).Delete
Sheets("Raw").Copy after:=Sheets("Raw")
Sheets("Raw (2)").Name = "DataHPC"
Application.DisplayAlerts = True


'Determine last row/column used in data set, then removes the "totals" information in the row below the last row
lastDataCol = Sheets("DataHPC").Cells(1, 1).End(xlToRight).Column
lastDataRow = Sheets("DataHPC").Cells(1, 2).End(xlDown).Row
Sheets("DataHPC").Rows(lastDataRow + 1).ClearContents

'Choose two cells (m2/n2) to act as "filler", then selects a range to be filled.
'Input header information, then formulas. these formulas will go in cells m2/n2
'Finally, use the filler to fill the range of cells defined above
'This method is faster than looping through individual cells
Set sourceRange = Sheets("DataHPC").Range("M2:N2")
Set fillRange = Sheets("DataHPC").Range("M2:N" & lastDataRow)
Sheets("DataHPC").Cells(1, lastDataCol + 1).Value = "Paid Year"
Sheets("DataHPC").Cells(1, lastDataCol + 2).Value = "Paid Month"
Sheets("DataHPC").Cells(2, lastDataCol + 1).Value = "=Year(I2)"
Sheets("DataHPC").Cells(2, lastDataCol + 2).Value = "=Month(I2)"
sourceRange.autofill Destination:=fillRange



'Find and replace keywords (in plan description column) to make sheet readable in pivot table
'Use this definition for the following| ~ = NOT operator
'd = find(word), if ~d = nothing, ==> ~nothing = d, ==> d = something
With Sheets("DataHPC").Range(Sheets("DataHPC").Cells(2, 3), Sheets("DataHPC").Cells(lastDataRow, 3))
     Set d = .Find("Pharmacy", LookIn:=xlValues)
        If Not d Is Nothing Then
        Do
            d.Value = "Rx"
            Set d = .FindNext(d)
            Loop While Not d Is Nothing
        End If
End With
With Sheets("DataHPC").Range(Sheets("DataHPC").Cells(2, 3), Sheets("DataHPC").Cells(lastDataRow, 3))
     Set c = .Find("Dental ", LookIn:=xlValues)
        If Not c Is Nothing Then
        Do
            c.Value = "Dental"
            Set c = .FindNext(c)
            Loop While Not c Is Nothing
        End If
End With
With Sheets("DataHPC").Range(Sheets("DataHPC").Cells(2, 3), Sheets("DataHPC").Cells(lastDataRow, 3))
     Set e = .Find("Vision ", LookIn:=xlValues)
        If Not e Is Nothing Then
        Do
            e.Value = "Vision"
            Set e = .FindNext(e)
            Loop While Not e Is Nothing
        End If
End With

'Fill remaining cells in Plan Description Column, not account for in the loops above
Set rng = Range(Sheets("DataHPC").Cells(2, 3), Sheets("DataHPC").Cells(lastDataRow, 3))
For Each cell In rng.Cells
    If cell.Find("Dental") Is Nothing Then
        If cell.Find("Vision") Is Nothing Then
            If cell.Find("Rx") Is Nothing Then
                cell.Value = "Med"
            End If
        End If
    End If
Next cell

'Create LC sheet from HPC sheet
Sheets("DataHPC").Copy after:=Sheets("DataHPC")
Sheets("DataHPC (2)").Name = "DataLC"

'Create header, populates values, autofills the values down the rows until the last row
Sheets("DataLC").Range("O1").Value = "Unique ID"
Sheets("DataLC").Range("O2").Value = "=D2&F2&G2"
Set sourceRange = Sheets("DataLC").Range("O2")
Set fillRange = Sheets("DataLC").Range("O2:O" & lastDataRow)
sourceRange.autofill Destination:=fillRange


'For the LC table, we do not need vision or dental claims
For i = 2 To lastDataRow
   If Sheets("DataLC").Cells(i, 3) = "Vision" Then
        Sheets("DataLC").Cells(i, 11).Value = 0
       End If
    If Sheets("DataLC").Cells(i, 3) = "Dental" Then
        Sheets("DataLC").Cells(i, 11).Value = 0
        End If
Next i


'---------------------------------Creation of pivot tables worksheet and health paid claims (HPC) pivot table-----------------------------------------
'Create pivot tables worksheet
Sheets.Add after:=Sheets("DataLC")
ActiveSheet.Name = "PivotTables"

'Create HPC pivot Cache
Set rangeHPC = Sheets("DataHPC").Range("A1").CurrentRegion
Set cacheHPC = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=rangeHPC)

'Create HPC pivot table
Set PivotTables = cacheHPC.CreatePivotTable _
(TableDestination:=Sheets("PivotTables").Cells(1, 1), _
TableName:="pivotTableHPC")

'Insert columns (paid year/paid month)-HPC
With ActiveSheet.PivotTables("PivotTableHPC").PivotFields("Paid Year")
.Orientation = xlColumnField
.Position = 1
End With
With ActiveSheet.PivotTables("PivotTableHPC").PivotFields("Paid Month")
.Orientation = xlColumnField
.Position = 2
End With

'Insert rows (plan description)-HPC
With ActiveSheet.PivotTables("PivotTableHPC").PivotFields("Plan Description")
.Orientation = xlRowField
.Position = 1
End With

'Insert data filed (paid amount)-HPC
With ActiveSheet.PivotTables("PivotTableHPC").PivotFields("Paid Amount")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "$#,##.##"
.Name = "Paid Amount "
End With

Set PivTbl = Worksheets("PivotTables").PivotTables("PivotTableHPC")
With PivTbl
   'Loop through all fields in PivotTable, set subtotals to false
   For Each PivFld In .PivotFields
      PivFld.Subtotals(1) = False
   Next PivFld
   'Set grand total to False
     .ColumnGrand = False
     .RowGrand = False
End With


'---------------------------------Large claimant (LC) pivot table creation------------------------------------
Set rangeLC = Sheets("DataLC").Range("A1").CurrentRegion
Set cacheLC = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=rangeLC)
 
'Create LC pivot table
Set PivotTables = cacheLC.CreatePivotTable _
(TableDestination:=Sheets("PivotTables").Cells(15, 1), _
TableName:="PivotTableLC")

'Insert rows (Unique ID)-LC
With ActiveSheet.PivotTables("PivotTableLC").PivotFields("Unique ID")
.Orientation = xlRowField
.Position = 1
End With

'Insert (paid year/paid month)-LC
With ActiveSheet.PivotTables("PivotTableLC").PivotFields("Paid Year")
.Orientation = xlColumnField
.Position = 1
End With
With ActiveSheet.PivotTables("PivotTableLC").PivotFields("Paid Month")
.Orientation = xlColumnField
.Position = 2
End With

Set PivTbl = Worksheets("PivotTables").PivotTables("PivotTableLC")
With PivTbl
   'Loop through all fields in PivotTable, set subtotals to false
   For Each PivFld In .PivotFields
      PivFld.Subtotals(1) = False
   Next PivFld
   'Set grand total to False/True
     .ColumnGrand = False
     .RowGrand = True
End With

'Insert data field(paid amount) -LC
With ActiveSheet.PivotTables("PivotTableLC").PivotFields("Paid Amount")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "$#,##.##"
End With

'Sort from largest to smallest
With Worksheets("PivotTables").PivotTables("PivotTableLC").PivotFields("Unique ID")
    .AutoSort xlDescending, "Sum of Paid Amount"
End With

'Creates final worksheet, this sheet will contain all relevant financial input
Sheets.Add(after:=Sheets(Sheets.count)).Name = "InputsFinancial"
num = Sheets("PivotTables").Range("A3").End(xlToRight).Column

'The placement of the LC pivot table on the pivot table worksheet ensures that the 1st unique ID starts on row 18
'The user is expected to place the large claimant threshold on the Criteria worksheet in cell B1. This tracks all claimants over half that amount...
    'and adjusts the endRow variable to match that row
i = 18
endRow = 0
Do While Sheets("PivotTables").Cells(i, num + 1) > Worksheets("Criteria").Cells(1, 2) / 2
    endRow = endRow + 1
    i = i + 1
Loop

'Copy/paste both pivot tables
Sheets("PivotTables").Range(Sheets("PivotTables").Cells(17, 1), Sheets("PivotTables").Cells(17 + endRow, 14)).Copy
Sheets("InputsFinancial").Cells(1, 1).PasteSpecial
Sheets("PivotTables").Range("A1").CurrentRegion.Copy
Sheets("InputsFinancial").Cells(endRow + 1 + 2, 1).PasteSpecial Transpose:=True


'Loop to calculate overages on LC per member per month (remember, months are cols in this)
endCol = Sheets("InputsFinancial").Cells(1, Columns.count).End(xlToLeft).Column
Sheets("InputsFinancial").Activate
 For rowCounter = 0 To endRow - 1
    For colCounter = 0 To endCol - 3
        zeroStep = Cells(rowCounter + 2, 2) - Worksheets("Criteria").Cells(1, 2)
        firstStep = Application.Sum(Range(Cells(rowCounter + 2, 2), Cells(rowCounter + 2, colCounter + 2)))
        secondStep = Application.Sum(Range(Cells(rowCounter + 2, 2), Cells(rowCounter + 2, colCounter + 3)))
        maxZero = Application.WorksheetFunction.Max(0, zeroStep)
        maxFirst = Application.WorksheetFunction.Max(0, firstStep - Worksheets("Criteria").Cells(1, 2))
        maxSecond = Application.WorksheetFunction.Max(0, secondStep - Worksheets("Criteria").Cells(1, 2))
                    
        If colCounter = 0 Then
            Worksheets("InputsFinancial").Cells(rowCounter + 2, endCol + 3).Value = maxZero
        End If
        
        finalCalc = Application.WorksheetFunction.Max(maxSecond - maxFirst)
        Worksheets("InputsFinancial").Cells(rowCounter + 2, endCol + 4 + colCounter).Value = finalCalc
        Next colCounter
    Next rowCounter

'Place the LC totals alongside the pasted HPC data and autofit the cells to show all information
Worksheets("InputsFinancial").Cells(endRow + 2 + 1 + k, 8).Value = "Over ISL"
For k = 0 To endCol - 2
    Worksheets("InputsFinancial").Cells(endRow + 2 + 2 + k, 8).Value = Application.Sum(Range(Cells(2, endCol + 3 + k), Cells(endRow + 1, endCol + 3 + k)))
Next k
Sheets("InputsFinancial").UsedRange.Columns.AutoFit

'Pouplates header information, and activates the inputs financial worksheet
endRow = Sheets("InputsFinancial").Cells(1, 1).End(xlDown).Row + 2
Worksheets("InputsFinancial").Cells(endRow, 9).Value = "Members"
Worksheets("Criteria").Activate
Sheets("Criteria").Range(Cells(4, 2), Cells(4, 13)).Copy
Sheets("InputsFinancial").Cells(endRow + 1, 9).PasteSpecial Transpose:=True
Sheets("InputsFinancial").Activate

End If
End Sub
