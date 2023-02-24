Sub DB1_ClaimsManipulation()
'This macro manipulates a specific worksheet into a format that makes the copy/paste process (into another workbook) easier.

'Prelimanary work, ensuring the worksheet is formatted as needed, and preventing pop-ups when deleting sheets
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
Sheets("Scrubbed").Delete
Sheets("PivotTables").Delete
Sheets(1).Name = "Raw"
Sheets("Raw").Copy after:=Sheets("Raw")
Sheets("Raw (2)").Name = "Scrubbed"
Sheets.Add before:=ActiveSheet
ActiveSheet.Name = "PivotTables"
Sheets("Scrubbed").Activate
Worksheets("Scrubbed").Range("A:A").Replace What:=" ", Replacement:=""
Range("A:A").NumberFormat = "mm-dd-yyyy;@"
On Error GoTo 0

'declare variables
Dim headerRow As Integer
Dim LastRow As Integer
Dim top As Integer
Dim bot As Integer
Dim lastCol As Integer
Dim startRange As Range
Dim fillRange As Range


'This is an odd bit, but basically it looks for the 1st and last row, the last row is actually 2 up from the lastRow variable (just the way the report is set up)
'then it looks at the top part of the row, and skips down to the bottom using end(xldown)
'the bot(tom) portion is set to the top (the top is the date, the bot is the 'Month-Year Total' string, we're replacing the string with the date
'finally, the top (date) is removed, as we have moved that to the bot
LastRow = Sheets("Scrubbed").Range("A:A").Find("GrandTotal").Row - 2
headerRow = Sheets("Scrubbed").Range("A:A").Find("YTD/MONTH").Row
top = headerRow + 1
bot = 0
While bot < LastRow
    bot = Cells(top, 1).End(xlDown).Row
    Cells(bot, 1).Value = Cells(top, 1)
    Cells(top, 1).Value = ""
    top = bot + 2
Wend

'Places header information, and begins the autofill of the formulas to calculate med costs. This is faster than a loop
lastCol = Sheets("Scrubbed").Cells(headerRow, 1).End(xlToRight).Column
Sheets("Scrubbed").Cells(headerRow, lastCol + 1).Value = "Medical"

'Id like to get rid of this system here, too stagnet. No room for dynamic allocation of cell values.
Sheets("Scrubbed").Range("K7").Value = "=E7+F7"
Set sourceRange = Sheets("Scrubbed").Range("k7")
Set fillRange = Sheets("Scrubbed").Range("k7:k" & LastRow)
sourceRange.autofill Destination:=fillRange
 
'Create pivot table
'Define Data Range
Set PSheet = Worksheets("PivotTables")
Set DSheet = Worksheets("Scrubbed")
With Sheets("Scrubbed")
    Set PRange = Range(Cells(headerRow, 1), Cells(LastRow, lastCol + 1))
End With
Sheets("pivotTables").Activate

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="ClaimsPivotTable")


'Insert Row Fields
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("YTD/MONTH")
.Orientation = xlRowField
.Position = 1
End With


'Insert Data Field
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("Medical")
.Orientation = xlDataField
.Function = xlSum
.Position = 1
.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
End With
With ActiveSheet.PivotTables("ClaimsPivotTable").PivotFields("DRUG")
.Orientation = xlDataField
.Function = xlSum
.Position = 2
.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
End With

End Sub
