Sub DB3_Enrollment()
Dim endRow As Integer
Dim endCol As Integer
Dim sourceRange As Range
Dim fillRange As Range
Dim sht As Worksheet
Dim fndList As Integer
Dim rplcList As Integer
Dim tbl As ListObject
Dim myArray As Variant
Dim PTable As PivotTable

'Prevents double running
If Workbooks("BulkClientEnrollment").Worksheets.count < 2 Then
    Workbooks("BulkClientEnrollment").Worksheets(1).Copy after:=Worksheets(1)
    Workbooks("BulkClientEnrollment").Worksheets(1).Name = "Raw"
    Workbooks("BulkClientEnrollment").Worksheets(2).Name = "Scrubbed"
    Else
        Exit Sub
End If

Application.ScreenUpdating = False

endRow = Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Cells(1, 1).End(xlDown).Row
endCol = Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Cells(1, 1).End(xlToRight).Column
ClientStructureFilePath = "P:\Docs\Work\Projects\Client\ClientFacetsClientStructure.xlsx" 'file location and name
ClientStrucutreFile = "ClientFacetsClientStructure" 'string name of file
fndList = 1
rplcList = 2


'sorts enrollment, makes it easier to delete all 0's rather than comb through the entire file
With Workbooks("BulkClientEnrollment").Worksheets("Scrubbed")
.Range(.Cells(1, 1), .Cells(endRow, endCol)).Sort key1:=.Cells(1, endCol), order1:=xlAscending, header:=xlYes
End With

'finds the first row with value greater than zero, and saves that row - 1 for deleting
For rowCounter = 2 To endRow
    If Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Cells(rowCounter, endCol).Value > 0 Then
    deleteRows = rowCounter - 1
    Exit For
    End If
Next rowCounter

'selects all rows with values of 0, and deltes them.
With Workbooks("BulkClientEnrollment").Worksheets("Scrubbed")
    .Range(.Cells(2, endCol), .Cells(deleteRows, endCol)).EntireRow.Delete
End With
    
'creates header and fills first value. This creates the key which will be replaced by the short plan name
Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Cells(1, endCol + 1).Value = "Plan"
Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Cells(2, endCol + 1).Value = "=A2&C2"

'sets range, and autofills the keys through the file
Set sourceRange = Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Range("G2")
Set fillRange = Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Range(Cells(2, endCol + 1), Cells(endRow, endCol + 1))
sourceRange.autofill Destination:=fillRange






'convert to values before we find/replace
With Range("A1").CurrentRegion
    .Value = .Value
End With

'Create variable to point to table
Set tbl = Workbooks.Open(ClientStructureFilePath).Worksheets("Structure").ListObjects("ClientPlanKey")

'Create an Array out of the Table's Data
Set TempArray = tbl.DataBodyRange
myArray = Application.Transpose(TempArray)
Workbooks(ClientStrucutreFile).Close

'Loop through each item in Array lists, find and replace key with short plan name
For x = LBound(myArray, 1) To UBound(myArray, 2)
    Workbooks("BulkClientEnrollment").Worksheets("Scrubbed").Cells.Replace What:=myArray(fndList, x), Replacement:=myArray(rplcList, x), _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next x

'adds a new sheet for a pivot table
ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count)
Workbooks("BulkClientEnrollment").Worksheets(3).Name = "PivotTable"

'create pivot cache
Set rangeEnrollment = Sheets("Scrubbed").Range("A1").CurrentRegion
Set cacheEnrollment = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=rangeEnrollment)

'create pivot table
Set PTable = cacheEnrollment.CreatePivotTable _
(TableDestination:=Sheets("PivotTable").Cells(2, 2), _
TableName:="EnrollmentPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("ACCOUNT")
.Orientation = xlRowField
.Position = 1
End With
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("Plan")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Data Field
With ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("SumOfENROLLMENT")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
.Name = "Sum"
End With

'Format Pivot Table, ensure cols are in correct order
ActiveSheet.PivotTables("EnrollmentPivotTable").ShowTableStyleRowStripes = False
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPONLY").Position = 1
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPSPOUSE").Position = 2
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPCHILDREN").Position = 3
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPFAMILY").Position = 4
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPLOYEE                 ").Position = 5
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPLOYEE + SPOUSE        ").Position = 6
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPLOYEE + CHILD(REN)    ").Position = 7
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPLOYEE + FAMILY        ").Position = 8
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPLOYEE +1              ").Position = 9
ActiveSheet.PivotTables("EnrollmentPivotTable").PivotFields("TIER").PivotItems( _
        "EMPLOYEE + 2 OR MORE DEPE").Position = 10
        
'remove subtotals and grand totals
With PTable
   'Loop through all fields in PivotTable, set subtotals to false
   For Each PivFld In .PivotFields
      PivFld.Subtotals(1) = False
   Next PivFld
   'Set grand total to False
     .ColumnGrand = False
     .RowGrand = False
End With
End Sub
