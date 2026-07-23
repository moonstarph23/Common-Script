Sub CopyFilterSheet()
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
    Set targetSheet = mainWorkbook.Sheets("FILTER") ' Change Cover to your target sheet name
    Set filesSheet = mainWorkbook.Sheets("Files") ' Set Files sheet

    ' Get the source file path from Files sheet cell B5
    sourceFilePath = filesSheet.Range("B5").Value

    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceSheet = sourceWorkbook.Sheets(1) ' Change if the data is not on the first sheet

    ' Determine the last row with data in column A of the source sheet
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

    ' Copy the data from columns A to A, including values and formats
    sourceSheet.Range("A1:C" & lastSourceRow).Copy
    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
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
    filesSheet.Range("C5").Value = "Successful"
    Exit Sub
ErrorHandler:

    ' If there is an error, write the error message in Files sheet cell B3
    filesSheet.Range("C5").Value = "Error: " & Err.Description
    Application.DisplayAlerts = originalDisplayAlerts

    'MsgBox "Error: " & Err.Description
End Sub
