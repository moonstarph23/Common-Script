Sub CopyOTHERSheet()
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
    Set targetSheet = mainWorkbook.Sheets("OTHERS") ' Change Cover to your target sheet name
    Set sourceSheet = mainWorkbook.Sheets("EMP CLOSED CHECK") ' Change if the data is not on the first sheet
    Set filesSheet = mainWorkbook.Sheets("Files") ' Set Files sheet
    Set wsCriteria = mainWorkbook.Sheets("FILTER")

    ' Determine the last row with data in column A of the source sheet'
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row

    '-------
Dim PercHolder As Double
Dim PerProg As Integer
Application.ScreenUpdating = False
Dim F As Object
   Dim fso As Object

'Dim varfile As Variant
Dim Rng As Range
    critLastRow = wsCriteria.Cells(wsCriteria.Rows.Count, "C").End(xlUp).Row

    ' Initialize the criteria array based on non-blank cells in Sheet2 Column A
    ReDim criteria(1 To critLastRow)
    critIndex = 1
    For i = 1 To critLastRow
        If wsCriteria.Cells(i, 3).Value <> "" Then
            criteria(critIndex) = wsCriteria.Cells(i, 3).Value
            critIndex = critIndex + 1
        End If
    Next i
sourceSheet.Range("A2:L" & lastSourceRow).AutoFilter Field:=4, Criteria1:=criteria, Operator:=xlFilterValues

 ' Copy the data from columns AP to AV, including values and formats'
    sourceSheet.Range("B3:K" & lastSourceRow).SpecialCells(xlCellTypeVisible).Copy

    ' Find the first empty row starting from A9
    If targetSheet.Range("B9").Value = "" Then
        emptyRow = 8 ' If A9 is empty, the first empty row is 9
    Else
        emptyRow = targetSheet.Cells(targetSheet.Rows.Count, 2).End(xlUp).Row + 1

        ' Ensure that the first empty row is at or below A9
        If emptyRow < 8 Then emptyRow = 8
    End If
    targetSheet.Range("B" & emptyRow + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    'targetSheet.Range("A" & emptyRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Restore the original state of DisplayAlerts'
    Application.DisplayAlerts = originalDisplayAlerts
    targetSheet.Activate
    Range("A1").Select

    ' Find the first empty row starting from A9
    If targetSheet.Range("B9").Value = "" Then
        emptyRowF = 9 ' If A9 is empty, the first empty row is 9
    Else
        emptyRowF = targetSheet.Cells(targetSheet.Rows.Count, 2).End(xlUp).Row + 1

        ' Ensure that the first empty row is at or below A9
        If emptyRowF < 9 Then emptyRowF = 9
    End If

    ' Insert the SUM formula from A10 to the last filled row
    targetSheet.Range("A" & emptyRow + 1 & ":A" & emptyRowF - 1).Value = Format(Date - 1, "mm/dd/yyyy")
    targetSheet.Range("L6:M6").Copy
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
    filesSheet.Range("C11").Value = "Successful"
    Exit Sub
ErrorHandler:

    ' If there is an error, write the error message in Files sheet cell B3
    filesSheet.Range("C11").Value = "Error: " & Err.Description
    Application.DisplayAlerts = originalDisplayAlerts

    'MsgBox "Error: " & Err.Description
End Sub
