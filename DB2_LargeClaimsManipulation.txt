Sub DB2_LargeClaims()

Dim LastRow As Integer
Dim yyyymm(1 To 12) As String
Dim mmm(1 To 12) As String
Dim benefitPlan(1 To 4) As String
Dim financialPlan(1 To 4) As String
Dim LCTable As PivotTable
Dim CTable As PivotTable


Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
Sheets("LCACP").Delete
Sheets("PivotTable").Delete
Sheets(1).Name = "Raw"
Sheets("Raw").Copy after:=Sheets("Raw")
Sheets("Raw (2)").Name = "LCACP"
On Error GoTo 0

'these are the actual string values originally created by the excel sheet
yyyymm(1) = "Paid Year & Month:   202001"
yyyymm(2) = "Paid Year & Month:   202002"
yyyymm(3) = "Paid Year & Month:   202003"
yyyymm(4) = "Paid Year & Month:   202004"
yyyymm(5) = "Paid Year & Month:   202005"
yyyymm(6) = "Paid Year & Month:   202006"
yyyymm(7) = "Paid Year & Month:   202007"
yyyymm(8) = "Paid Year & Month:   202008"
yyyymm(9) = "Paid Year & Month:   202009"
yyyymm(10) = "Paid Year & Month:   202010"
yyyymm(11) = "Paid Year & Month:   202011"
yyyymm(12) = "Paid Year & Month:   202012"
'these are the names we're going to use to replace the above list
mmm(1) = "Jan"
mmm(2) = "Feb"
mmm(3) = "Mar"
mmm(4) = "Apr"
mmm(5) = "May"
mmm(6) = "Jun"
mmm(7) = "Jul"
mmm(8) = "Aug"
mmm(9) = "Sep"
mmm(10) = "Oct"
mmm(11) = "Nov"
mmm(12) = "Dec"

'These are the actual string values originally created by the excel sheet. Redacted
benefitPlan(1) = 1
benefitPlan(2) = 2
benefitPlan(3) = 3
benefitPlan(4) = 4
'These are the plan names we're going to use in the financial
financialPlan(1) = a
financialPlan(2) = b
financialPlan(3) = c
financialPlan(4) = d

'Create headers
Worksheets("LCACP").Cells(1, 13).Value = "Type"
Worksheets("LCACP").Cells(1, 14).Value = "Month"
Worksheets("LCACP").Cells(1, 15).Value = "Plan"

MonthEntry.Show
endMonth = Worksheets("Raw").Cells(1, 16)
'Find the last row, then move six rows up, to exclude footer information
LastRow = Worksheets("LCACP").Cells(Rows.count, 1).End(xlUp).Row

With Sheets("LCACP").Range(Sheets("LCACP").Cells(1, 1), Sheets("LCACP").Cells(LastRow, 1))
    Set deleteRow = .Find("Subtotals", LookIn:=xlValues, LookAt:=xlPart)
        If Not deleteRow Is Nothing Then
            Do
                Row = deleteRow.Row
                Sheets("LCACP").Rows(Row).Delete Shift:=xlShiftUp
                Set deleteRow = .Find("Subtotals", LookIn:=xlValues, LookAt:=xlPart)
            Loop While Not deleteRow Is Nothing
        End If
End With

LastRow = Worksheets("LCACP").Cells(Rows.count, 1).End(xlUp).Row

'Find end rows for medical and Rx, and save those row values
With Sheets("LCACP").Range(Sheets("LCACP").Cells(1, 1), Sheets("LCACP").Cells(LastRow, 1))
        Set endRx = .Find("Totals for Prescription Drug:", LookIn:=xlValues, LookAt:=xlWhole)
        endRxRow = endRx.Row - 1
        Set endMed = .Find("Totals for Medical:", LookIn:=xlValues, LookAt:=xlWhole)
        endMedRow = endMed.Row - 1
End With
'Fill the empty cells with the respective values
Sheets("LCACP").Range(Sheets("LCACP").Cells(2, 13), Sheets("LCACP").Cells(endMedRow, 13)).Value = "Med"
Sheets("LCACP").Range(Sheets("LCACP").Cells(endMedRow + 2, 13), Sheets("LCACP").Cells(endRxRow, 13)).Value = "Rx"

'Find and replace values for Months
For i = 1 To endMonth
    With Sheets("LCACP").Range(Sheets("LCACP").Cells(1, 1), Sheets("LCACP").Cells(endRxRow, 1))
        Set monthYear = .Find(yyyymm(i), LookIn:=xlValues, LookAt:=xlWhole)
            If Not monthYear Is Nothing Then
            Do
                monthYearRow = monthYear.Row
                monthYear.Value = mmm(i)
                Sheets("LCACP").Cells(monthYearRow, 14).Value = mmm(i)
                Set monthYear = .FindNext(monthYear)
                Loop While Not monthYear Is Nothing
            End If
    End With
Next i
'populate blank cells with approporate month values
top = 2
While top < endRxRow
    bot = Cells(top, 14).End(xlDown).Row - 1
    If bot > endRxRow Then
        bot = endRxRow
    End If
        Range(Cells(top, 14), Cells(bot, 14)) = Cells(top, 14)
    top = bot + 1
Wend

'find replace and offset text for plan options
For i = 1 To 4
    With Sheets("LCACP").Range(Sheets("LCACP").Cells(1, 1), Sheets("LCACP").Cells(endRxRow, 1))
        Set longPlan = .Find(benefitPlan(i), LookIn:=xlValues, LookAt:=xlWhole)
            If Not longPlan Is Nothing Then
            Do
                longPlanRow = longPlan.Row
                longPlan.Value = financialPlan(i)
                Sheets("LCACP").Cells(longPlanRow, 15).Value = financialPlan(i)
                Set longPlan = .FindNext(longPlan)
                Loop While Not longPlan Is Nothing
            End If
    End With
Next i
'populate blank cells with approporate plan vlaues
top = 3
While top < endRxRow
    bot = Cells(top, 15).End(xlDown).Row - 1
    If bot > endRxRow Then
        bot = endRxRow
    End If
        Range(Cells(top, 15), Cells(bot, 15)) = Cells(top, 15)
    top = bot + 1
Wend


Sheets.Add after:=Sheets("LCACP")
ActiveSheet.Name = "PivotTable"

'Declare Variables

'create pivot cache
Set claimsRange = Sheets("LCACP").Range("A1").CurrentRegion
Set claimsCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=claimsRange)

'create pivot table
Set LCTable = claimsCache.CreatePivotTable _
(TableDestination:=Sheets("PivotTable").Cells(2, 2), _
TableName:="LCPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("LCPivotTable").PivotFields("Member ID")
.Orientation = xlRowField
.Position = 1
End With
With ActiveSheet.PivotTables("LCPivotTable").PivotFields("Member Name")
.Orientation = xlRowField
.Position = 2
End With
With ActiveSheet.PivotTables("LCPivotTable").PivotFields("Relationship")
.Orientation = xlRowField
.Position = 3
End With
With ActiveSheet.PivotTables("LCPivotTable").PivotFields("Plan")
.Orientation = xlRowField
.Position = 4
End With

'Insert Data Field
With ActiveSheet.PivotTables("LCPivotTable").PivotFields("Paid Amt")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
.Name = "Sum"
End With

'remove subtotals and grand totals
With LCTable
   'Loop through all fields in PivotTable, set subtotals to false
   For Each PivFld In .PivotFields
      PivFld.Subtotals(1) = False
   Next PivFld
   'Set grand total to False
     .ColumnGrand = False
     .RowGrand = False
End With

Worksheets("PivotTable").PivotTables("LCPivotTable").RowAxisLayout xlTabularRow
With Worksheets("PivotTable").PivotTables("LCPivotTable").PivotFields("Member ID")
    .AutoSort xlDescending, "Sum"
End With

Set lcCache = claimsCache

'create pivot table
Set CTable = lcCache.CreatePivotTable _
(TableDestination:=Sheets("PivotTable").Cells(2, 8), _
TableName:="ClaimsPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Month")
.Orientation = xlRowField
.Position = 1
End With

'Insert Column Fields
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Type")
.Orientation = xlColumnField
.Position = 1
End With
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Plan")
.Orientation = xlColumnField
.Position = 2
End With

'Insert Data Field
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Paid Amt")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
.Name = "Sum"
End With

'remove subtotals and grand totals
With CTable
   'Loop through all fields in PivotTable, set subtotals to false
   For Each PivFld In .PivotFields
      PivFld.Subtotals(1) = False
   Next PivFld
   'Set grand total to False
     .ColumnGrand = False
     .RowGrand = False
End With

End Sub
