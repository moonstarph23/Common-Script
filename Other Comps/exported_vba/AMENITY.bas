Sub CopyHASheet()
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
    Set targetSheet = mainWorkbook.Sheets("Htl Amenity") ' Change Cover to your target sheet name
    Set sourceSheet = mainWorkbook.Sheets("EMP CLOSED CHECK") ' Change if the data is not on the first sheet
    Set filesSheet = mainWorkbook.Sheets("Files") ' Set Files sheet

    ' Determine the last row with data in column A of the source sheet'
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row

    ' Set the range to filter'
    Set Rng = sourceSheet.Range("A2:L" & lastSourceRow) ' Adjust the range as necessary

    ' Clear any existing filters'
    If sourceSheet.AutoFilterMode Then sourceSheet.AutoFilterMode = False

    ' Apply the filter'
    Rng.AutoFilter Field:=4, Criteria1:="*COMP Htl Amnty*"

    ' Apply the filter'

    'Rng.AutoFilter Field:=44, Criteria1:="*HH*", Operator:=xlOr, Criteria2:="*PF*"

    ' Copy the data from columns AP to AV, including values and formats'
    sourceSheet.Range("B3:K" & lastSourceRow).SpecialCells(xlCellTypeVisible).Copy

    ' Find the first empty row starting from A9
    If targetSheet.Range("B11").Value = "" Then
        emptyRow = 10 ' If A9 is empty, the first empty row is 9
    Else
        emptyRow = targetSheet.Cells(targetSheet.Rows.Count, 2).End(xlUp).Row + 1

        ' Ensure that the first empty row is at or below A9
        If emptyRow < 10 Then emptyRow = 10
    End If
    targetSheet.Range("B" & emptyRow + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    'targetSheet.Range("A" & emptyRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Restore the original state of DisplayAlerts'
    Application.DisplayAlerts = originalDisplayAlerts
    targetSheet.Activate
    Range("A1").Select

    ' Find the first empty row starting from A9
    If targetSheet.Range("B11").Value = "" Then
        emptyRowF = 11 ' If A9 is empty, the first empty row is 9
    Else
        emptyRowF = targetSheet.Cells(targetSheet.Rows.Count, 2).End(xlUp).Row + 1

        ' Ensure that the first empty row is at or below A9
        If emptyRowF < 11 Then emptyRowF = 11
    End If

    ' Insert the SUM formula from A10 to the last filled row
    targetSheet.Range("A" & emptyRow + 1 & ":A" & emptyRowF - 1).Value = Format(Date - 1, "mm/dd/yyyy")
    targetSheet.Range("L8:P8").Copy
    targetSheet.Range("L" & emptyRow + 1 & ":L" & emptyRowF - 1).PasteSpecial
    targetSheet.Range("L" & emptyRow + 1 & ":L" & emptyRowF - 1).Value = targetSheet.Range("L" & emptyRow + 1 & ":L" & emptyRowF - 1).Value
    Application.CutCopyMode = False
    On Error Resume Next

                    ' Clear any existing filters'
    If sourceSheet.AutoFilterMode Then sourceSheet.AutoFilterMode = False
            On Error GoTo 0

    ' Optional: Clear the clipboard
    Application.CutCopyMode = False

    ' If no errors, write "successful" in Files sheet cell B3
    filesSheet.Range("C9").Value = "Successful"
    Exit Sub
ErrorHandler:

    ' If there is an error, write the error message in Files sheet cell B3
    filesSheet.Range("C9").Value = "Error: " & Err.Description
    Application.DisplayAlerts = originalDisplayAlerts

    'MsgBox "Error: " & Err.Description
End Sub
