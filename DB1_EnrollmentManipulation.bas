Sub DB1_EnrollmentCount()
'This macro manipulates a specific worksheet into a format that makes the copy/paste process (into another workbook) easier.

Dim sht As Worksheet
Dim fndList As Variant
Dim rplcList As Variant
Dim x As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
'Disable errors to prevent pop-ups when checking for double run
On Error Resume Next

'Redacted due to sensitive information
fndList = Array()
rplcList = Array()

'Prevent double-run
Sheets("Scrubbed").Delete
Sheets("PivotTable").Delete

Sheets(1).Name = "Raw"
Sheets("Raw").Copy after:=Sheets("Raw")
Sheets("Raw (2)").Name = "Scrubbed"
Sheets.Add after:=Sheets("Scrubbed")
ActiveSheet.Name = "PivotTable"
'Resume error catching
On Error GoTo 0

'Loop through each item in Array lists
  For x = LBound(fndList) To UBound(fndList)
    'Loop through worksheet to find and replace items in respective (redacted) lists
        Worksheets("Scrubbed").Cells.Replace What:=fndList(x), Replacement:=rplcList(x), _
          LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
          SearchFormat:=False, ReplaceFormat:=False
  Next x

'Create elements of pivottable
'Create pivot cache
Set enrollmentRange = Sheets("Scrubbed").Range("A1").CurrentRegion
Set enrollmentCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=enrollmentRange)

'Create pivot table
Set EnrollmentTable = enrollmentCache.CreatePivotTable _
(TableDestination:=Sheets("PivotTable").Cells(2, 2), _
TableName:="EnrollmentPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Month")
.Orientation = xlRowField
.Position = 1
End With

'Insert Colulmn Fields
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Benefit Plan")
.Orientation = xlColumnField
.Position = 1
End With
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Coverage Tier")
.Orientation = xlColumnField
.Position = 2
End With

'Insert Data Field
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Subscribers")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "##"
.Name = "Sum"
End With

'Remove subtotals and grand totals
With EnrollmentTable
   'Loop through all fields in PivotTable, set subtotals to false
   For Each PivFld In .PivotFields
      PivFld.Subtotals(1) = False
   Next PivFld
   'Set grand total to False
     .ColumnGrand = False
     .RowGrand = False
End With

'Ensure Order of outputs is as desired. Redacted due to sensitive information
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Coverage Tier").PivotItems(1).Position = 1
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Coverage Tier").PivotItems(2).Position = 2
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Coverage Tier").PivotItems(3).Position = 3
        
End Sub
