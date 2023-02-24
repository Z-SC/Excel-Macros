Sub LocationMatch()
'This macro manipulates a specific worksheet into a format that makes the copy/paste process (into another workbook) easier.

Application.ScreenUpdating = False
'Specific names redacted
Dim reportNames(1 To 8) As String
reportNames(1) = 1
reportNames(2) = 2
reportNames(3) = 3
reportNames(4) = 4
reportNames(5) = 5
reportNames(6) = 6
reportNames(7) = 7
reportNames(8) = 8

Dim financialNames(1 To 8) As String
financialNames(1) = a
financialNames(2) = b
financialNames(3) = c
financialNames(4) = d
financialNames(5) = e
financialNames(6) = f
financialNames(7) = g
financialNames(8) = h

'if we don't unmerge no 'find' feature will work
numsheets = Worksheets.count
For i = 1 + 1 To numsheets
    Worksheets(i).Cells.UnMerge
Next i

'disable error notifications, delete then create "Data" sheet
Application.DisplayAlerts = False
On Error Resume Next
Sheets("Data").Delete
On Error GoTo 0
Application.DisplayAlerts = True
Sheets.Add(after:=Sheets(Sheets.count)).Name = "Data"

'watch the name of counters, i j k are not specific and might lead to counfusion

colPlacement = 1
For reportNamesCounter = 1 To 8
worksheetCounter = 1
    For Each ws In Worksheets
        With ws.UsedRange
            Set rFound = .Find(reportNames(reportNamesCounter), after:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole)
            If Not rFound Is Nothing Then
                On Error GoTo ErrorHandle
                'this switches out names, prevents double results on the 'find' portion
                Worksheets("Data").Cells(1, colPlacement + 2).Value = financialNames(colPlacement)
                'uses 'find' to look for keywords to determine specific rows
                medRow = Worksheets(worksheetCounter).Range("A:A").Find("Medical").Row
                visionRow = Worksheets(worksheetCounter).Range("A:A").Find("Vision").Row
                eeCol = Worksheets(worksheetCounter).Cells.Find("Employee").Column
                depCol = Worksheets(worksheetCounter).Cells.Find("Dependent").Column
                eeTotalRow = Worksheets(worksheetCounter).Cells(Rows.count, eeCol).End(xlUp).Row
                depTotalRow = eeTotalRow
                'filling out employee/dependent med/vision totals
                Sheets("Data").Cells(2, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(medRow, eeCol)
                Sheets("Data").Cells(3, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(medRow, depCol)
                Sheets("Data").Cells(8, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(visionRow, eeCol)
                Sheets("Data").Cells(9, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(visionRow, depCol)
                'dental totals are total - (med+vision)
                Sheets("Data").Cells(5, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(eeTotalRow, eeCol) - (Sheets(worksheetCounter).Cells(medRow, eeCol) + Sheets(worksheetCounter).Cells(visionRow, eeCol))
                Sheets("Data").Cells(6, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(depTotalRow, depCol) - (Sheets(worksheetCounter).Cells(medRow, depCol) + Sheets(worksheetCounter).Cells(visionRow, depCol))
                colPlacement = colPlacement + 1
            End If
        End With
       worksheetCounter = worksheetCounter + 1
    Next ws
Next reportNamesCounter

For i = 1 To 3
    Worksheets("Data").Cells((i * 2) + (i - 1), 2).Value = "Employee"
    Worksheets("Data").Cells((i * 3), 2).Value = "Dependent"
Next i
Worksheets("Data").Cells(2, 1).Value = "Medical"
Worksheets("Data").Cells(5, 1).Value = "Dental"
Worksheets("Data").Cells(8, 1).Value = "Vision"
Worksheets("Data").UsedRange.Columns.AutoFit
Worksheets("Data").Cells.NumberFormat = "$#,##"

done:
Exit Sub

ErrorHandle:
eeTotalRow = eeTotalRow - 1
depTotalRow = depTotalRow - 1
Sheets("Data").Cells(5, colPlacement + 2).Value = Sheets(worksheetCounter).Cells(eeTotalRow, eeCol) - (Sheets(worksheetCounter).Cells(medRow, eeCol) + Sheets(worksheetCounter).Cells(visionRow, eeCol))
Resume
End Sub
