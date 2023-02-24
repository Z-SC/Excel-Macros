Sub Claims_Database()
'This macro manipulates data from a specific worksheet into a format that makes the copy/paste process (into another workbook) easier.

'Prevents double running
If Workbooks("BulkClientClaims").Worksheets.count < 2 Then
    Workbooks("BulkClientClaims").Worksheets(1).Copy after:=Worksheets(1)
    Workbooks("BulkClientClaims").Worksheets(1).Name = "Raw"
    Workbooks("BulkClientClaims").Worksheets(2).Name = "Scrubbed"
    Else
        Exit Sub
End If

endRow = Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells(1, 1).End(xlDown).Row
endCol = Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells(1, 1).End(xlToRight).Column
ClientStructureFilePath = "P:\Docs\Work\Projects\Client\ClientFacetsClientStructure.xlsx" 'file location and name
ClientStrucutreFile = "ClientFacetsClientStructure" 'string name of file
fndList = 1
rplcList = 2

'creates header and fills first value. This creates the key which will be replaced by the short plan name
Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells(1, endCol + 1).Value = "Plan"
Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells(2, endCol + 1).Value = "=A2&D2"
Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells(1, endCol + 2).Value = "Med Claims"
Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells(2, endCol + 2).Value = "=SUM(F2:H2)"


'sets range, and autofills the keys through the file
Set sourceRange = Workbooks("BulkClientClaims").Worksheets("Scrubbed").Range("R2:S2")
Set fillRange = Workbooks("BulkClientClaims").Worksheets("Scrubbed").Range(Cells(2, endCol + 1), Cells(endRow, endCol + 2))
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
    Workbooks("BulkClientClaims").Worksheets("Scrubbed").Cells.Replace What:=myArray(fndList, x), Replacement:=myArray(rplcList, x), _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next x

'adds a new sheet for a pivot table
ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count)
Workbooks("BulkClientClaims").Worksheets(3).Name = "PivotTable"

'create pivot cache (Claims)
Set rangeEnrollment = Sheets("Scrubbed").Range("A1").CurrentRegion
Set cacheClaims = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=rangeEnrollment)

'create pivot table (Claims)
Set PTable = cacheClaims.CreatePivotTable _
(TableDestination:=Sheets("PivotTable").Cells(2, 2), _
TableName:="ClaimsPivotTable")

'Insert Row Fields (Claims)
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("ACCOUNT")
.Orientation = xlRowField
.Position = 1
End With
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Plan")
.Orientation = xlRowField
.Position = 2
End With

'Insert Data Field (Claims)
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Med Claims")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "$#,##0"
.Name = "Med_Claims"
End With
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("DRUG")
.Orientation = xlDataField
.Position = 2
.Function = xlSum
.NumberFormat = "$#,##0"
.Name = "Drug_Claims"
End With

'create cache and pivot table (membership)
Set cacheMembership = cacheClaims
Set QTable = cacheMembership.CreatePivotTable _
(TableDestination:=Sheets("PivotTable").Cells(2, 7), _
TableName:="MembershipPivotTable")

'Insert Row Fields (memberhsip)
With ActiveSheet.PivotTables("MembershipPivotTable").PivotFields("ACCOUNT")
.Orientation = xlRowField
.Position = 1
End With

'Insert Data Field (membership)
With ActiveSheet.PivotTables("MembershipPivotTable").PivotFields("TOTAL MBRS")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
.Name = "Members"
End With
End Sub
