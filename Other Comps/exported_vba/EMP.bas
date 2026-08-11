Attribute VB_Name = "EMP"
Sub CopyECCSheet()
    On Error GoTo ErrorHandler
    Dim mainWorkbook As Workbook
    Dim rawSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim filesSheet As Worksheet
    Dim rowCount As Long
    Dim totalRows As Long
    Dim rawFirstRow As Long
    Dim rawLastRow As Long
    Dim firstRawRowToAppend As Long
    Dim batchRows As Long
    Dim firstTargetRow As Long
    Dim lastTargetRow As Long
    Dim expectedTargetRow As Long
    Dim errorMessage As String

    Set mainWorkbook = ThisWorkbook
    Set filesSheet = mainWorkbook.Sheets("Files")
    Set rawSheet = mainWorkbook.Sheets("RAW")
    Set targetSheet = mainWorkbook.Sheets("EMP CLOSED CHECK")

    rowCount = rawSheet.Cells(rawSheet.Rows.Count, "A").End(xlUp).Row
    totalRows = rowCount - 1
    If totalRows < 0 Then totalRows = 0
    rawSheet.Range("AP1").Value = totalRows

    If Not IsNumeric(rawSheet.Range("AP2").Value) Or _
            Not IsNumeric(rawSheet.Range("AP3").Value) Then
        Err.Raise vbObjectError + 1001, "CopyECCSheet", _
            "Latest RAW batch boundaries are missing."
    End If

    rawFirstRow = CLng(rawSheet.Range("AP2").Value)
    rawLastRow = CLng(rawSheet.Range("AP3").Value)
    If rawFirstRow < 2 Or rawLastRow < rawFirstRow Or rawLastRow > rowCount Then
        Err.Raise vbObjectError + 1002, "CopyECCSheet", _
            "Latest RAW batch boundaries are invalid."
    End If

    ' Row 3 is the template and already represents RAW row 2 on the first batch.
    firstRawRowToAppend = rawFirstRow
    If firstRawRowToAppend = 2 Then firstRawRowToAppend = 3
    batchRows = rawLastRow - firstRawRowToAppend + 1

    If batchRows > 0 Then
        firstTargetRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row + 1
        If firstTargetRow < 4 Then firstTargetRow = 4
        expectedTargetRow = firstRawRowToAppend + 1
        If firstTargetRow <> expectedTargetRow Then
            Err.Raise vbObjectError + 1003, "CopyECCSheet", _
                "RAW and EMP CLOSED CHECK rows are out of sync."
        End If
        lastTargetRow = firstTargetRow + batchRows - 1

        Application.ScreenUpdating = False
        targetSheet.Range("A3:L3").Copy
        targetSheet.Range("A" & firstTargetRow & ":L" & lastTargetRow).PasteSpecial
    End If

    Application.CutCopyMode = False
    filesSheet.Range("C6").Value = "Successful"
    Exit Sub
ErrorHandler:
    errorMessage = Err.Description
    On Error Resume Next
    filesSheet.Range("C6").Value = "Error: " & errorMessage
    Application.CutCopyMode = False
End Sub
