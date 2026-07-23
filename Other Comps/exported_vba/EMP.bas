Sub CopyECCSheet()
    On Error GoTo ErrorHandler

    ' Define the first row to start counting (row 6)
    firstRow = 2

   'Get the last row with data in column A
    RowCount = Sheets("RAW").Cells(Sheets("RAW").Rows.Count, "A").End(xlUp).Row

    ' Calculate the number of rows from row 6 onwards
    totalRows = RowCount - firstRow + 1

    ' Ensure the row count is valid (non-negative)
    If totalRows < 0 Then totalRows = 0

    ' Store the total row count for further use

    ' Example: Storing it in cell B1
    Sheets("RAW").Range("AP1").Value = totalRows

    ' Disable screen updating and set calculation to manual
    Application.ScreenUpdating = False

    'Application.Calculation = xlCalculationManual
    lastRow = 3 + totalRows - 1
    Sheets("EMP CLOSED CHECK").Range("A3:L3").Copy
    Sheets("EMP CLOSED CHECK").Range("A4:L" & lastRow).PasteSpecial

    ' Restore screen updating and calculation
    Application.CutCopyMode = False

    ' If no errors, write "successful" in Files sheet cell C4
     Sheets("Files").Range("C6").Value = "Successful"
    Exit Sub
ErrorHandler:

    ' If there is an error, write the error message in Files sheet cell C4
    Sheets("Files").Range("C6").Value = "Error: " & Err.Description
    Application.DisplayAlerts = originalDisplayAlerts

    'MsgBox "Error: " & Err.Description
End Sub
