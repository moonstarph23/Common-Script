Sub CopyRAWSheet()
    On Error GoTo ErrorHandler
    Dim mainWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim filesSheet As Worksheet
    Dim sourceFilePath As String
    Dim lastRow As Long
    Dim currentRow As Long
    Dim lastSourceRow As Long
    Dim deptRevRow As Long
    Dim deptRevTotalRow As Long
    Dim originalDisplayAlerts As Boolean

    ' Save the current state of DisplayAlerts
    originalDisplayAlerts = Application.DisplayAlerts

    ' Turn off display alerts to avoid large clipboard warning
    Application.DisplayAlerts = False

    ' Set the main workbook and target sheet
    Set mainWorkbook = ThisWorkbook
    Set targetSheet = mainWorkbook.Sheets("RAW") ' Change Cover to your target sheet name
    Set filesSheet = mainWorkbook.Sheets("Files") ' Set Files sheet
    Set wsCriteria = mainWorkbook.Sheets("FILTER")

    ' Get the source file path from Files sheet cell B5
    sourceFilePath = filesSheet.Range("B4").Value

    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceSheet = sourceWorkbook.Sheets(1) ' Change if the data is not on the first sheet
    Set NewSheet = sourceWorkbook.Sheets.Add
    NewSheet.Name = "NewSheetName" ' You can set your desired sheet name here
    Set NewSheet = sourceWorkbook.Sheets("NewSheetName")

    'Unmerge cells and unwrap text in columns A to AK
    With sourceSheet.Range("A:AN")
        .UnMerge

        '.WrapText = False
    End With

    ' Determine the last row with data in column A of the source sheet
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "V").End(xlUp).Row

    ' Copy the data from columns A to A, including values and formats
    sourceSheet.Range("A6:AN" & lastSourceRow).Copy
    NewSheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    NewSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
    With NewSheet.Range("A:AN")

    '.UnMerge
    .WrapText = False
    End With
    NewSheet.Columns("G:G").Delete
    NewSheet.Columns("F:F").Delete
    NewSheet.Columns("E:E").Delete
    NewSheet.Columns("D:D").Delete
    NewSheet.Columns("C:C").Delete
 lastSourceRowN = NewSheet.Cells(NewSheet.Rows.Count, "A").End(xlUp).Row
    NewSheet.Range("B1:B" & lastSourceRowN).Value = NewSheet.Range("B1:B" & lastSourceRowN).Value
    Dim topRows As Range
      Set ws = sourceWorkbook.Sheets("NewSheetName")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    '-------
Dim PercHolder As Double
Dim PerProg As Integer
Application.ScreenUpdating = False
Dim F As Object
   Dim fso As Object

'Dim varfile As Variant
Dim Rng As Range
    critLastRow = wsCriteria.Cells(wsCriteria.Rows.Count, "A").End(xlUp).Row

    ' Initialize the criteria array based on non-blank cells in Sheet2 Column A
    ReDim criteria(1 To critLastRow)
    critIndex = 1
    For i = 1 To critLastRow
        If wsCriteria.Cells(i, 1).Value <> "" Then
            criteria(critIndex) = wsCriteria.Cells(i, 1).Value
            critIndex = critIndex + 1
        End If
    Next i
NewSheet.Range("A1:AI1").AutoFilter Field:=17, Criteria1:=criteria, Operator:=xlFilterValues
Set Rng = NewSheet.Range("A2:AI" & lastSourceRowN).CurrentRegion
    NewSheet.Range("A2:AI" & lastSourceRowN).Copy
    targetSheet.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
    Application.DisplayAlerts = False
    Application.CutCopyMode = False

    ' Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False

    ' Restore the original state of DisplayAlerts
    Application.DisplayAlerts = originalDisplayAlerts

    ' Set alignment for columns A and B to left
    With targetSheet
        .Columns("A:AI").HorizontalAlignment = xlLeft
    End With
    targetSheet.Activate
    Range("A1").Select

    ' If no errors, write "successful" in Files sheet cell B3
    filesSheet.Range("C4").Value = "Successful"
    Exit Sub
ErrorHandler:

    ' If there is an error, write the error message in Files sheet cell B3
    filesSheet.Range("C4").Value = "Error: " & Err.Description
    Application.DisplayAlerts = originalDisplayAlerts

    'MsgBox "Error: " & Err.Description
End Sub
