Sub CopyMISheet()
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
    Set targetSheet = mainWorkbook.Sheets("MENU ITEM 2") ' Change Cover to your target sheet name
    Set filesSheet = mainWorkbook.Sheets("Files") ' Set Files sheet

    ' Get the source file path from Files sheet cell B5
    sourceFilePath = filesSheet.Range("B3").Value

    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceSheet = sourceWorkbook.Sheets(1) ' Change if the data is not on the first sheet
    Set NewSheet = sourceWorkbook.Sheets.Add
    NewSheet.Name = "NewSheetName" ' You can set your desired sheet name here
    Set NewSheet = sourceWorkbook.Sheets("NewSheetName")

    'Unmerge cells and unwrap text in columns A to AK
    With sourceSheet.Range("A:AC")
        .UnMerge
        .WrapText = False
    End With

    ' Determine the last row with data in column A of the source sheet
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "Q").End(xlUp).Row

    ' Copy the data from columns A to A, including values and formats
    sourceSheet.Range("A11:AC" & lastSourceRow).Copy
    NewSheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    NewSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
    NewSheet.Columns("AB:AB").Delete
    NewSheet.Columns("X:X").Delete
    NewSheet.Columns("U:U").Delete
    NewSheet.Columns("S:S").Delete
    NewSheet.Columns("O:O").Delete
    NewSheet.Columns("N:N").Delete
    NewSheet.Columns("L:L").Delete
    NewSheet.Columns("J:J").Delete
    NewSheet.Columns("G:G").Delete
    NewSheet.Columns("F:F").Delete
    NewSheet.Columns("E:E").Delete
    NewSheet.Columns("C:C").Delete

    ' Determine the last row with data in column A of the source sheet
    lastSourceRowN = NewSheet.Cells(NewSheet.Rows.Count, "A").End(xlUp).Row
    NewSheet.Range("B1:B" & lastSourceRowN).Value = NewSheet.Range("B1:B" & lastSourceRowN).Value
    NewSheet.Range("D1:D" & lastSourceRowN).Value = NewSheet.Range("D1:D" & lastSourceRowN).Value
    NewSheet.Range("E1:E" & lastSourceRowN).Value = NewSheet.Range("E1:E" & lastSourceRowN).Value
    Dim Rng As Range
    Dim topRows As Range
      Set ws = sourceWorkbook.Sheets("NewSheetName")

        ' Set the range to apply the filter
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'MsgBox lastRow
    Set Rng = ws.Range("K1:K" & lastRow)
    Set dta = ws.Range("A2:Q" & lastRow)

    ' Apply filter to delete MODIFIER
    With ws
        .AutoFilterMode = False
       Rng.AutoFilter Field:=1, Criteria1:="<>MODIFIER"

        ' Delete visible rows
        On Error Resume Next ' In case there are no visible rows
    dta.SpecialCells(xlCellTypeVisible).Copy
        On Error GoTo 0

        ' Remove filter
        .AutoFilterMode = False
    End With

    'NewSheet.Range("A2:Q" & lastSourceRowN).Copy
    targetSheet.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    targetSheet.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    targetSheet.Range("A2").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    lastSourceRowT = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    targetSheet.Range("S1").Copy
    targetSheet.Range("S2:S" & lastSourceRowT).PasteSpecial

    'targetSheet.Range("B1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False

    ' Restore the original state of DisplayAlerts
    Application.DisplayAlerts = originalDisplayAlerts

    ' Set alignment for columns A and B to left
    With targetSheet
        .Columns("A:B").HorizontalAlignment = xlLeft
    End With
    targetSheet.Activate
    Range("A1").Select

    ' If no errors, write "successful" in Files sheet cell B3
    filesSheet.Range("C3").Value = "Successful"
    Exit Sub
ErrorHandler:

    ' If there is an error, write the error message in Files sheet cell B3
    filesSheet.Range("C3").Value = "Error: " & Err.Description
    Application.DisplayAlerts = originalDisplayAlerts

    'MsgBox "Error: " & Err.Description
End Sub
