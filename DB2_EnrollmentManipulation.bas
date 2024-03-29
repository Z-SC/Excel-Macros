Sub DB2_EnrollmentManipulation()
'This macro manipulates a specific worksheet into a format that makes the copy/paste process (into another workbook) easier.

'declare variables
Dim benefitOption(1 To 5) As String
Dim financialPlan(1 To 5) As String
Dim headerRow As Integer
Dim LastRow As Integer
Dim count As Integer
Dim element As Integer
Dim top As Integer
Dim bot As Integer
Dim headerTitles As New Collection

'This section looks defines arrays with specific naming conventions. The exact names have been redacted
benefitOption(1) = a
benefitOption(2) = b
benefitOption(3) = c
benefitOption(4) = d
benefitOption(5) = e

financialPlan(1) = 1
financialPlan(2) = 2
financialPlan(3) = 3
financialPlan(4) = 4
financialPlan(5) = 5

'Prelim work, copy and paste sheets, format data, determine values
Worksheets(1).Copy after:=Worksheets(1)
Sheets(1).Name = "Raw"
Sheets(2).Name = "Scrubbed"
Sheets("Scrubbed").Activate
Range("A:A").NumberFormat = "mm-dd-yyyy;@"
headerRow = Sheets("Scrubbed").Range("A:A").Find("YTD/MONTH").Row
LastRow = Sheets("Scrubbed").Range("A:A").Find("BENEFIT OPTION").Row

count = headerRow
i = 0

'Looks for the specific key phrase Total, in order to determine where the counter should stop.
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


'Loop through each item in Array lists
  For x = 1 To 5
    'Loop through worksheet
          Sheets("Scrubbed").Cells.Replace What:=benefitOption(x), Replacement:=financialPlan(x), _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
  Next x


For j = 3 To 6
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
For i = 1 To 4
With ActiveSheet.PivotTables("EnrollmentPivotTabls").PivotFields(headerTitles(i))
.Orientation = xlDataField
.Function = xlSum
.Position = i
End With
Next i
End Sub
