Sub DB4_EnrollmentManipulation()

Dim headerRow As Integer
Dim LastRow As Integer
Dim deleteRows As New Collection
Dim element As Integer
Dim fndList As Integer
Dim rplcList As Integer
Dim tbl As ListObject
Dim myArray As Variant
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim lastCol As Long
Dim headerRowLength As Integer
Dim headerTitles As New Collection
Dim tiers As String

'prelim work, copy and paste sheets, format data, determine values
Worksheets("Scrubbed").Copy after:=Worksheets("Scrubbed")
Sheets("Data (2)").Name = "Raw"
Sheets("Scrubbed").Activate
Range("A:A").NumberFormat = "mm-dd-yyyy;@"
headerRow = Sheets("Scrubbed").Range("A:A").Find("YTD/Month").Row
LastRow = Sheets("Scrubbed").Range("A:A").Find("BENEFIT OPTION").Row
count = headerRow
i = 0

While count < (LastRow - i * 2)
element = Sheets("Scrubbed").Range("A:A").Find("Total", after:=Cells(count, 1)).Row
Sheets("Scrubbed").Rows(element).Delete
Sheets("Scrubbed").Rows(element).Delete
i = i + 1
count = element
Wend

top = headerRow + 1
While bot < (LastRow - 2 * i)
    bot = Cells(top, 1).End(xlDown).Row - 1
    Range(Cells(top, 1), Cells(bot, 1)) = Cells(top, 1)
    top = bot + 1
Wend

'Designate Columns for Find/Replace data
  fndList = 1
  rplcList = 2

'Loop through each item in Array lists
  For x = LBound(myArray, 1) To UBound(myArray, 2)
    'Loop through worksheet
          Sheets("Scrubbed").Cells.Replace What:=myArray(fndList, x), Replacement:=myArray(rplcList, x), _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
  Next x

headerRowLength = Sheets("Scrubbed").Cells(headerRow, 1).End(xlToRight).Column
For j = 3 To headerRowLength - j
    tiers = Sheets("Scrubbed").Cells(5, j)
    headerTitles.Add tiers
Next j

'Insert a New Blank Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add before:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")
Set DSheet = Worksheets("Scrubbed")

'Define Data Range
LastRow = (LastRow - 2 * i)
lastCol = DSheet.Cells(headerRow, 1).End(xlToRight).Column
Set PRange = DSheet.Cells(headerRow, 1).CurrentRegion

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="EnrollmentPivotTabls")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="EnrollmentPivotTabls")

'Insert Row Fields
With ActiveSheet.PivotTables("EnrollmentPivotTabls").PivotFields("YTD/MONTH")
.Orientation = xlRowField
.Position = 1
End With

'Insert Column Fields
With ActiveSheet.PivotTables("EnrollmentPivotTabls").PivotFields("BENEFIT OPTION")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Data Field
For i = 1 To (headerRowLength - 2)
With ActiveSheet.PivotTables("EnrollmentPivotTabls").PivotFields(headerTitles(i))
.Orientation = xlDataField
.Function = xlSum
.Position = i
End With
Next i
End Sub
