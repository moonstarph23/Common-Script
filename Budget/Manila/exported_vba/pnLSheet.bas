Attribute VB_Name = "pnLSheet"
Private Function GetLastUsedRow(ByVal ws As Worksheet, Optional ByVal minimumRow As Long = 8) As Long

    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        GetLastUsedRow = minimumRow
    ElseIf lastCell.row < minimumRow Then
        GetLastUsedRow = minimumRow
    Else
        GetLastUsedRow = lastCell.row
    End If

End Function

Private Function GetLastDataRow(ByVal ws As Worksheet, ByVal keyColumn As Variant, Optional ByVal minimumRow As Long = 12) As Long

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).row
    If lastRow < minimumRow Then lastRow = minimumRow
    GetLastDataRow = lastRow

End Function

Private Function ValidatePnLTargetRows(ByVal sourceData As Variant, ByVal sourceSheetName As String, _
                                       ByVal deptCol As Long, ByVal department As String, _
                                       ByVal targetSheet As Worksheet, ByVal targetLastRow As Long, _
                                       ByRef validationError As String) As Boolean

    Const TARGET_ROW_COLUMN As Long = 31 ' AE
    Dim i As Long
    Dim mappingValue As Variant
    Dim targetRow As Long
    Dim includeRow As Boolean

    ValidatePnLTargetRows = False

    For i = 1 To UBound(sourceData, 1)
        If deptCol = 0 Then
            includeRow = True
        Else
            includeRow = (CStr(sourceData(i, deptCol)) = department)
        End If

        If includeRow Then
            mappingValue = sourceData(i, TARGET_ROW_COLUMN)

            If IsError(mappingValue) Then
                validationError = sourceSheetName & " row " & (i + 11) & _
                                  " has an Excel error in the P&L target column AE."
                Exit Function
            ElseIf Not IsEmpty(mappingValue) And CStr(mappingValue) <> "" Then
                If Not IsNumeric(mappingValue) Then
                    validationError = sourceSheetName & " row " & (i + 11) & _
                                      " has a nonnumeric P&L target in AE: '" & CStr(mappingValue) & "'."
                    Exit Function
                End If

                If CDbl(mappingValue) = 0 Then GoTo NextMappingRow

                If CDbl(mappingValue) <> Fix(CDbl(mappingValue)) Then
                    validationError = sourceSheetName & " row " & (i + 11) & _
                                      " has a non-integer P&L target in AE: '" & CStr(mappingValue) & "'."
                    Exit Function
                End If

                targetRow = CLng(mappingValue)
                If targetRow < 8 Or targetRow > targetLastRow Then
                    validationError = sourceSheetName & " row " & (i + 11) & _
                                      " maps AE to P&L row " & targetRow & _
                                      ", outside the valid range 8:" & targetLastRow & "."
                    Exit Function
                End If

                If Trim(CStr(targetSheet.Cells(targetRow, "F").value)) = "" Then
                    validationError = sourceSheetName & " row " & (i + 11) & _
                                      " maps AE to P&L row " & targetRow & _
                                      ", but column F on that row is blank."
                    Exit Function
                End If
            End If
        End If
NextMappingRow:
    Next i

    ValidatePnLTargetRows = True

End Function

Sub ClearPNL()

    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents

    Set wsTarget = ThisWorkbook.ActiveSheet
    lastRow = GetLastUsedRow(wsTarget, 8)
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Loop through rows 8 to lastRow in the target sheet
    For i = 8 To lastRow

        ' Check Column AO (41) for "Do not clear"
        If wsTarget.Cells(i, 41).value <> "Do not clear" Then
            wsTarget.Range(wsTarget.Cells(i, "G"), wsTarget.Cells(i, "R")).ClearContents
            wsTarget.Range(wsTarget.Cells(i, "W"), wsTarget.Cells(i, "AH")).ClearContents
        End If
    Next i

    wsTarget.Range(wsTarget.Cells(8, "AV"), wsTarget.Cells(lastRow, "BG")).ClearContents
    wsTarget.Range(wsTarget.Cells(8, "V"), wsTarget.Cells(lastRow, "V")).ClearContents
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents

End Sub

Sub ClearDeptPNL()

    Dim wsTarget As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Set wsTarget = ThisWorkbook.Sheets("P&L by Dept")
    lastRow = GetLastUsedRow(wsTarget, 8)

    ' Find last populated column in row 7 (from G to ZZZ)
    lastCol = wsTarget.Cells(7, wsTarget.Columns.Count).End(xlToLeft).Column
    If lastCol < 7 Then lastCol = 7 ' Ensure at least column G
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Loop through rows 8 to lastRow in the target sheet
    For i = 8 To lastRow

        ' Check Column D for "Do not clear"
        If wsTarget.Cells(i, 4).value <> "Do not clear" Then
            For j = 7 To lastCol ' Columns G to lastCol

                ' Do not clear if row 6 in this column is "Formula"
                If wsTarget.Cells(6, j).value <> "Formula" Then
                    wsTarget.Cells(i, j).ClearContents
                End If
            Next j
        End If
    Next i
    wsTarget.Calculate

    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents

End Sub

Sub updatePnL(deptCol As Long)

    '-----------------------------

    ' updatePnL: Optimized P&L update using single dictionary and bulk writes

    ' Flow: 1) Read data arrays 2) Build unified dictionary 3) Bulk write to master sheet

    '-----------------------------
    Dim department As String
    Dim lastRow As Long
    Dim i As Long, k As Long
    Dim startTime As Double, elapsedTime As Double
    Dim arrBudget As Variant, arrActual As Variant
    Dim AEcol As Long, StatCol As Long
    Dim wsTarget As Worksheet
    Dim pnlLastRow As Long, budgetLastRow As Long, actualLastRow As Long
    Dim validationError As String
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    ' Initialize
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    startTime = Timer
    Call declareGlobal
    Set wsTarget = ThisWorkbook.ActiveSheet

    department = CStr(wsTarget.Range("I1").value)
    pnlLastRow = GetLastUsedRow(wsTarget, 8)
    budgetLastRow = GetLastDataRow(globalBudgetFYSheet, "F", 12)
    actualLastRow = GetLastDataRow(globalActualFYSheet, "F", 12)

    ' Recalculate AE target-row formulas after rows are inserted in the P&L layout.
    globalBudgetFYSheet.Calculate
    globalActualFYSheet.Calculate

    arrBudget = globalBudgetFYSheet.Range("A12:AJ" & budgetLastRow).value
    arrActual = globalActualFYSheet.Range("A12:AJ" & actualLastRow).value

    If Not ValidatePnLTargetRows(arrBudget, globalBudgetFYSheet.Name, deptCol, department, _
                                 wsTarget, pnlLastRow, validationError) Then GoTo ValidationFailed
    If Not ValidatePnLTargetRows(arrActual, globalActualFYSheet.Name, deptCol, department, _
                                 wsTarget, pnlLastRow, validationError) Then GoTo ValidationFailed

    Call ClearPNL
    lastRow = pnlLastRow

    'DeptCol = 25 ' department column
    AEcol = 31 ' AE column (target row mapping)
    StatCol = 24 ' Statistics column
    Debug.Print "=== P&L Update Started for Department: " & department & " ==="

    '--- STEP 1: FY data was read and validated before clearing the target sheet ---

    '--- STEP 2: Build two separate dictionaries for G-R and W-AH columns ---
    Dim dictPnlRowsColGtoR As Object, dictPnlRowsColWtoAH As Object
    Set dictPnlRowsColGtoR = CreateObject("Scripting.Dictionary")
    Set dictPnlRowsColWtoAH = CreateObject("Scripting.Dictionary")
    Debug.Print "Building dictPnlRowsColGtoR (Budget data + G-R formulas)..."

    ' 2a) Budget data from globalBudgetFYSheet (filtered by department) -> G-R columns
    Dim budgetCount As Long: budgetCount = 0
    For i = 1 To UBound(arrBudget, 1)
        If arrBudget(i, deptCol) = department And arrBudget(i, AEcol) <> 0 Then
            Dim targetRow As Long: targetRow = arrBudget(i, AEcol)

            ' Check if column F (ledger account) starts with 9 to determine sign
            Dim ledgerAccount As String: ledgerAccount = CStr(arrBudget(i, 6)) ' Column F
            Dim signMultiplier As Double: signMultiplier = 1
            If Left(ledgerAccount, 1) <> "9" Then
                signMultiplier = -1
                If budgetCount < 5 Then ' Debug first few entries
                    Debug.Print "Budget Row " & i & ": Ledger '" & ledgerAccount & "' doesn't start with 9, inverting sign"
                End If
            End If

            ' Check if this target row already exists in dictionary
            If dictPnlRowsColGtoR.Exists(targetRow) Then

                ' Accumulate values to existing entry
                Dim existingRowGtoR As Variant: existingRowGtoR = dictPnlRowsColGtoR(targetRow)
                For k = 10 To 21
                    existingRowGtoR(k - 8) = existingRowGtoR(k - 8) + (arrBudget(i, k) * signMultiplier) ' Add values with sign correction
                Next k
                dictPnlRowsColGtoR(targetRow) = existingRowGtoR
            Else

                ' Create new entry
                Dim arrBudgetRowGtoR() As Variant
                ReDim arrBudgetRowGtoR(1 To 13) ' Row + 12 columns (G-R)
                arrBudgetRowGtoR(1) = targetRow ' Target row number
                For k = 10 To 21
                    arrBudgetRowGtoR(k - 8) = arrBudget(i, k) * signMultiplier ' J->2, K->3, ..., U->13 with sign correction
                Next k
                dictPnlRowsColGtoR(targetRow) = arrBudgetRowGtoR
            End If
            budgetCount = budgetCount + 1
        End If
    Next i
    Debug.Print "Budget rows added to G-R dictionary: " & budgetCount

    ' 2b) Budget formulas from MasterSheet (rows with AO = "Do not clear") -> G-R columns
    Dim formulaBudgetCount As Long: formulaBudgetCount = 0
    Debug.Print "Scanning formula rows from row 8 to " & lastRow & " for 'Do not clear' in column AO..."
    For i = 8 To lastRow
        If wsTarget.Cells(i, 41).value = "Do not clear" Then
            Dim arrFormulaBudgetGtoR() As Variant
            ReDim arrFormulaBudgetGtoR(1 To 13) ' Row + 12 columns (G-R)
            arrFormulaBudgetGtoR(1) = i ' Target row number
            For k = 1 To 12
                Dim formulaTextGR As String
                formulaTextGR = wsTarget.Cells(i, 6 + k).Formula ' G to R

                ' Write formula directly, no prefix needed
                If formulaTextGR <> "" Then
                    arrFormulaBudgetGtoR(k + 1) = formulaTextGR
                Else
                    arrFormulaBudgetGtoR(k + 1) = ""
                End If
            Next k
            dictPnlRowsColGtoR(i) = arrFormulaBudgetGtoR
            formulaBudgetCount = formulaBudgetCount + 1
        End If
    Next i
    Debug.Print "Formula Budget rows added to G-R dictionary: " & formulaBudgetCount
    Debug.Print "Building dictPnlRowsColWtoAH (Actual data + W-AH formulas)..."

    ' 2c) Actual data from globalActualFYSheet (filtered by department) -> W-AH columns
    Dim actualCount As Long: actualCount = 0
    For i = 1 To UBound(arrActual, 1)
        If arrActual(i, deptCol) = department And arrActual(i, AEcol) <> 0 Then
            Dim targetRowActual As Long: targetRowActual = arrActual(i, AEcol)

            ' Check if column F (ledger account) starts with 9 to determine sign
            Dim ledgerAccountActual As String: ledgerAccountActual = CStr(arrActual(i, 6)) ' Column F
            Dim signMultiplierActual As Double: signMultiplierActual = 1
            If Left(ledgerAccountActual, 1) <> "9" Then
                signMultiplierActual = -1
                If actualCount < 5 Then ' Debug first few entries
                    Debug.Print "Actual Row " & i & ": Ledger '" & ledgerAccountActual & "' doesn't start with 9, inverting sign (multiplier: " & signMultiplierActual & ")"
                End If
            Else
                If actualCount < 5 Then ' Debug first few entries
                    Debug.Print "Actual Row " & i & ": Ledger '" & ledgerAccountActual & "' starts with 9, keeping original sign (multiplier: " & signMultiplierActual & ")"
                End If
            End If

            ' Check if this target row already exists in dictionary
            If dictPnlRowsColWtoAH.Exists(targetRowActual) Then

                ' Accumulate values to existing entry
                Dim existingRowWtoAH As Variant: existingRowWtoAH = dictPnlRowsColWtoAH(targetRowActual)
                For k = 10 To 21
                    existingRowWtoAH(k - 8) = existingRowWtoAH(k - 8) + (arrActual(i, k) * signMultiplierActual) ' Add values with sign correction
                Next k
                dictPnlRowsColWtoAH(targetRowActual) = existingRowWtoAH
            Else

                ' Create new entry
                Dim arrActualRowWtoAH() As Variant
                ReDim arrActualRowWtoAH(1 To 13) ' Row + 12 columns (W-AH)
                arrActualRowWtoAH(1) = targetRowActual ' Target row number
                For k = 10 To 21
                    arrActualRowWtoAH(k - 8) = arrActual(i, k) * signMultiplierActual ' J->2, K->3, ..., U->13 with sign correction
                Next k
                dictPnlRowsColWtoAH(targetRowActual) = arrActualRowWtoAH
            End If
            actualCount = actualCount + 1
        End If
    Next i
    Debug.Print "Actual rows added to W-AH dictionary: " & actualCount

    ' 2d) Actual formulas from MasterSheet (rows with AO = "Do not clear") -> W-AH columns
    Dim formulaActualCount As Long: formulaActualCount = 0
    Debug.Print "Scanning same formula rows for W-AH columns..."
    For i = 8 To lastRow
        If wsTarget.Cells(i, 41).value = "Do not clear" Then
            Dim arrFormulaActualWtoAH() As Variant
            ReDim arrFormulaActualWtoAH(1 To 13) ' Row + 12 columns (W-AH)
            arrFormulaActualWtoAH(1) = i ' Target row number
            For k = 1 To 12
                Dim formulaTextWAH As String
                formulaTextWAH = wsTarget.Cells(i, 22 + k).Formula ' W to AH

                ' Write formula directly, no prefix needed
                If formulaTextWAH <> "" Then
                    arrFormulaActualWtoAH(k + 1) = formulaTextWAH
                Else
                    arrFormulaActualWtoAH(k + 1) = ""
                End If
            Next k
            dictPnlRowsColWtoAH(i) = arrFormulaActualWtoAH
            formulaActualCount = formulaActualCount + 1
        End If
    Next i
    Debug.Print "Formula Actual rows added to W-AH dictionary: " & formulaActualCount

    '--- STEP 2e: Process Special Win Rows (VIP Win, Mass Win, Total Slot Win) ---
    Call ProcessSpecialWinRows(department, deptCol, dictPnlRowsColGtoR, dictPnlRowsColWtoAH, arrBudget, arrActual)

    '--- STEP 3: Debug output for the first 10 entries of each dictionary (sorted by row) ---
    Debug.Print "=== 1) G-R Dictionary Data (first 10, sorted by row) ==="

    ' Sort G-R keys
    Dim keysGR As Variant, sortedKeysGR() As Long
    Dim countGR As Long: countGR = 0
    keysGR = dictPnlRowsColGtoR.keys
    ReDim sortedKeysGR(0 To UBound(keysGR))

    ' Copy and sort
    For i = 0 To UBound(keysGR)
        sortedKeysGR(i) = CLng(keysGR(i))
    Next i

    ' Simple bubble sort
    For i = 0 To UBound(sortedKeysGR) - 1
        For k = i + 1 To UBound(sortedKeysGR)
            If sortedKeysGR(i) > sortedKeysGR(k) Then
                Dim temp As Long
                temp = sortedKeysGR(i)
                sortedKeysGR(i) = sortedKeysGR(k)
                sortedKeysGR(k) = temp
            End If
        Next k
    Next i

    ' Print first 10 sorted entries
    For i = 0 To UBound(sortedKeysGR)
        If countGR >= 10 Then Exit For
        Dim rowDataGR As Variant: rowDataGR = dictPnlRowsColGtoR(sortedKeysGR(i))
        Dim debugStrGR As String: debugStrGR = "Row " & sortedKeysGR(i) & " -> Columns G-R: "
        For k = 2 To 13 ' Skip row number, show G-R data
            debugStrGR = debugStrGR & "[" & rowDataGR(k) & "] "
        Next k
        Debug.Print debugStrGR
        countGR = countGR + 1
    Next i
    Debug.Print "=== 2) W-AH Dictionary Data (first 10, sorted by row) ==="

    ' Sort W-AH keys
    Dim keysWAH As Variant, sortedKeysWAH() As Long
    Dim countWAH As Long: countWAH = 0
    keysWAH = dictPnlRowsColWtoAH.keys
    ReDim sortedKeysWAH(0 To UBound(keysWAH))

    ' Copy and sort
    For i = 0 To UBound(keysWAH)
        sortedKeysWAH(i) = CLng(keysWAH(i))
    Next i

    ' Simple bubble sort
    For i = 0 To UBound(sortedKeysWAH) - 1
        For k = i + 1 To UBound(sortedKeysWAH)
            If sortedKeysWAH(i) > sortedKeysWAH(k) Then
                temp = sortedKeysWAH(i)
                sortedKeysWAH(i) = sortedKeysWAH(k)
                sortedKeysWAH(k) = temp
            End If
        Next k
    Next i

    ' Print first 10 sorted entries
    For i = 0 To UBound(sortedKeysWAH)
        If countWAH >= 10 Then Exit For
        Dim rowDataWAH As Variant: rowDataWAH = dictPnlRowsColWtoAH(sortedKeysWAH(i))
        Dim debugStrWAH As String: debugStrWAH = "Row " & sortedKeysWAH(i) & " -> Columns W-AH: "
        For k = 2 To 13 ' Skip row number, show W-AH data
            debugStrWAH = debugStrWAH & "[" & rowDataWAH(k) & "] "
        Next k
        Debug.Print debugStrWAH
        countWAH = countWAH + 1
    Next i

    '--- STEP 4: Bulk write both dictionaries to master sheet ---
    Call BulkWriteGtoR(dictPnlRowsColGtoR, wsTarget)
    Call BulkWriteWtoAH(dictPnlRowsColWtoAH, wsTarget)
    Call BulkWriteAVtoBH(dictPnlRowsColGtoR, wsTarget)

    ' Finalize
    wsTarget.Calculate
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    elapsedTime = Timer - startTime
    Debug.Print "=== P&L Update Completed ==="
    Debug.Print "Total G-R dictionary entries: " & dictPnlRowsColGtoR.Count
    Debug.Print "Total W-AH dictionary entries: " & dictPnlRowsColWtoAH.Count

    'MsgBox "P&L Update completed!" & vbCrLf & "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf & "G-R entries: " & dictPnlRowsColGtoR.Count & ", W-AH entries: " & dictPnlRowsColWtoAH.Count, vbInformation
    Exit Sub

ValidationFailed:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    MsgBox "P&L refresh stopped before clearing any values." & vbCrLf & vbCrLf & validationError, _
           vbCritical, "Invalid P&L row mapping"

End Sub

Private Sub BulkWriteGtoR(dictPnlRowsColGtoR As Object, wsTarget As Worksheet)

    '--- Write G-R data (budget values + budget formulas) in bulk ---
    Debug.Print "=== Writing G-R Dictionary (Budget + Formulas) ==="
    Debug.Print "G-R dictionary entries: " & dictPnlRowsColGtoR.Count
    If dictPnlRowsColGtoR.Count = 0 Then Exit Sub

    ' Get sorted keys (row numbers)
    Dim keys As Variant, sortedKeys() As Long
    Dim i As Long, j As Long, temp As Long
    keys = dictPnlRowsColGtoR.keys
    ReDim sortedKeys(0 To UBound(keys))

    ' Copy keys to sortedKeys array
    For i = 0 To UBound(keys)
        sortedKeys(i) = CLng(keys(i))
    Next i

    ' Simple bubble sort by row number
    For i = 0 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Find min and max rows for bulk range
    Dim minRow As Long, maxRow As Long
    minRow = sortedKeys(0)
    maxRow = sortedKeys(UBound(sortedKeys))

    ' Create bulk arrays for the entire range
    Dim rangeRows As Long, rangeCols As Long
    rangeRows = maxRow - minRow + 1
    rangeCols = 12 ' G to R (12 columns)
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To rangeRows, 1 To rangeCols)

    ' Read existing values first
    Dim existingRange As Variant
    existingRange = wsTarget.Range(wsTarget.Cells(minRow, 7), wsTarget.Cells(maxRow, 18)).value

    ' Copy existing values to our array
    For i = 1 To rangeRows
        For j = 1 To rangeCols
            If IsArray(existingRange) Then
                bulkArray(i, j) = existingRange(i, j)
            Else
                bulkArray(i, j) = existingRange ' Single cell case
            End If
        Next j
    Next i

    ' Apply our dictionary data to the array
    Dim key As Variant, rowData As Variant
    Dim rowNum As Long, arrayRow As Long
    Dim cellValue As String
    For Each key In dictPnlRowsColGtoR.keys
        rowData = dictPnlRowsColGtoR(key)
        rowNum = rowData(1)
        arrayRow = rowNum - minRow + 1

        ' Apply each column value
        For j = 1 To 12 ' G to R columns
            cellValue = CStr(rowData(j + 1)) ' Skip row number, get data

            ' Apply values directly - formulas start with =, numbers are numeric
            If IsNumeric(cellValue) And cellValue <> "" Then
                bulkArray(arrayRow, j) = CDbl(cellValue)
            ElseIf cellValue <> "" Then
                bulkArray(arrayRow, j) = cellValue ' This includes formulas starting with =
            End If
        Next j
    Next key

    ' Write entire array to sheet in one operation
    wsTarget.Range(wsTarget.Cells(minRow, 7), wsTarget.Cells(maxRow, 18)).value = bulkArray
    Debug.Print "G-R bulk write completed for range " & minRow & ":" & maxRow & " (" & dictPnlRowsColGtoR.Count & " entries)"

End Sub

Private Sub BulkWriteWtoAH(dictPnlRowsColWtoAH As Object, wsTarget As Worksheet)

    '--- Write W-AH data (actual values + actual formulas) in bulk ---
    Debug.Print "=== Writing W-AH Dictionary (Actual + Formulas) ==="
    Debug.Print "W-AH dictionary entries: " & dictPnlRowsColWtoAH.Count
    If dictPnlRowsColWtoAH.Count = 0 Then Exit Sub

    ' Get sorted keys (row numbers)
    Dim keys As Variant, sortedKeys() As Long
    Dim i As Long, j As Long, temp As Long
    keys = dictPnlRowsColWtoAH.keys
    ReDim sortedKeys(0 To UBound(keys))

    ' Copy keys to sortedKeys array
    For i = 0 To UBound(keys)
        sortedKeys(i) = CLng(keys(i))
    Next i

    ' Simple bubble sort by row number
    For i = 0 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Find min and max rows for bulk range
    Dim minRow As Long, maxRow As Long
    minRow = sortedKeys(0)
    maxRow = sortedKeys(UBound(sortedKeys))

    ' Create bulk arrays for the entire range
    Dim rangeRows As Long, rangeCols As Long
    rangeRows = maxRow - minRow + 1
    rangeCols = 12 ' W to AH (12 columns)
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To rangeRows, 1 To rangeCols)

    ' Read existing values first
    Dim existingRange As Variant
    existingRange = wsTarget.Range(wsTarget.Cells(minRow, 23), wsTarget.Cells(maxRow, 34)).value

    ' Copy existing values to our array
    For i = 1 To rangeRows
        For j = 1 To rangeCols
            If IsArray(existingRange) Then
                bulkArray(i, j) = existingRange(i, j)
            Else
                bulkArray(i, j) = existingRange ' Single cell case
            End If
        Next j
    Next i

    ' Apply our dictionary data to the array
    Dim key As Variant, rowData As Variant
    Dim rowNum As Long, arrayRow As Long
    Dim cellValue As String
    For Each key In dictPnlRowsColWtoAH.keys
        rowData = dictPnlRowsColWtoAH(key)
        rowNum = rowData(1)
        arrayRow = rowNum - minRow + 1

        ' Apply each column value
        For j = 1 To 12 ' W to AH columns
            cellValue = CStr(rowData(j + 1)) ' Skip row number, get data

            ' Apply values directly - formulas start with =, numbers are numeric
            If IsNumeric(cellValue) And cellValue <> "" Then
                bulkArray(arrayRow, j) = CDbl(cellValue)
            ElseIf cellValue <> "" Then
                bulkArray(arrayRow, j) = cellValue ' This includes formulas starting with =
            End If
        Next j
    Next key

    ' Write entire array to sheet in one operation
    wsTarget.Range(wsTarget.Cells(minRow, 23), wsTarget.Cells(maxRow, 34)).value = bulkArray
    Debug.Print "W-AH bulk write completed for range " & minRow & ":" & maxRow & " (" & dictPnlRowsColWtoAH.Count & " entries)"

End Sub

Private Sub BulkWriteAVtoBH(dictPnlRowsColGtoR As Object, wsTarget As Worksheet)

    '--- Write AV-BG data (budget values + budget formulas) in bulk ---
    Debug.Print "=== Writing AV-BG Dictionary (Budget + Formulas) ==="
    Debug.Print "AV-BG dictionary entries: " & dictPnlRowsColGtoR.Count
    If dictPnlRowsColGtoR.Count = 0 Then Exit Sub

    ' Get sorted keys (row numbers)
    Dim keys As Variant, sortedKeys() As Long
    Dim i As Long, j As Long, temp As Long
    keys = dictPnlRowsColGtoR.keys
    ReDim sortedKeys(0 To UBound(keys))

    ' Copy keys to sortedKeys array
    For i = 0 To UBound(keys)
        sortedKeys(i) = CLng(keys(i))
    Next i

    ' Simple bubble sort by row number
    For i = 0 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Find min and max rows for bulk range
    Dim minRow As Long, maxRow As Long
    minRow = sortedKeys(0)
    maxRow = sortedKeys(UBound(sortedKeys))

    ' Create bulk arrays for the entire range
    Dim rangeRows As Long, rangeCols As Long
    rangeRows = maxRow - minRow + 1
    rangeCols = 12 ' AV to BG (12 columns)
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To rangeRows, 1 To rangeCols)

    ' Read existing values first
    Dim existingRange As Variant
    existingRange = wsTarget.Range(wsTarget.Cells(minRow, 48), wsTarget.Cells(maxRow, 59)).value

    ' Copy existing values to our array
    For i = 1 To rangeRows
        For j = 1 To rangeCols
            If IsArray(existingRange) Then
                bulkArray(i, j) = existingRange(i, j)
            Else
                bulkArray(i, j) = existingRange ' Single cell case
            End If
        Next j
    Next i

    ' Apply our dictionary data to the array
    Dim key As Variant, rowData As Variant
    Dim rowNum As Long, arrayRow As Long
    Dim cellValue As String
    For Each key In dictPnlRowsColGtoR.keys
        rowData = dictPnlRowsColGtoR(key)
        rowNum = rowData(1)
        arrayRow = rowNum - minRow + 1

        ' Apply each column value
        For j = 1 To 12 ' AV to BG columns
            cellValue = CStr(rowData(j + 1)) ' Skip row number, get data

            ' Apply values directly - formulas start with =, numbers are numeric
            If IsNumeric(cellValue) And cellValue <> "" Then
                bulkArray(arrayRow, j) = CDbl(cellValue)
            ElseIf cellValue <> "" Then
                bulkArray(arrayRow, j) = cellValue ' This includes formulas starting with =
            End If
        Next j
    Next key

    ' Write entire array to sheet in one operation
    wsTarget.Range(wsTarget.Cells(minRow, 48), wsTarget.Cells(maxRow, 59)).value = bulkArray
    Debug.Print "AV-BG bulk write completed for range " & minRow & ":" & maxRow & " (" & dictPnlRowsColGtoR.Count & " entries)"

End Sub



Public Sub exportPnLs()

    '-----------------------------

    ' exportPnLs: Export P&L data to separate workbooks based on Setup sheet Column Q values

    ' Each workbook contains sheets for each unique Department (Column D)

    '-----------------------------
    Dim saveDialog As FileDialog
    Dim savePath As String
    Dim uniqueFilenames As Object
    Dim uniqueDepts As Object
    Dim setupLastRow As Long
    Dim i As Long
    Dim filename As String
    Dim CleanFilename As String
    Dim dept As String
    Dim newWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim masterDict As Object

    ' Initialize
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call declareGlobal

    ' File dialog for save location
    Set saveDialog = Application.FileDialog(msoFileDialogFolderPicker)
    saveDialog.Title = "Select folder to save exported P&L files"
    If saveDialog.Show = -1 Then
        savePath = saveDialog.SelectedItems(1)
    Else
        MsgBox "Export cancelled."
        Exit Sub
    End If

    ' Get unique values from Setup sheet Column Q (Q9 to lastrow)
    setupLastRow = globalsetupSheet.Cells(globalsetupSheet.Rows.Count, "Q").End(xlUp).row
    Set uniqueFilenames = CreateObject("Scripting.Dictionary")
    Set uniqueDepts = CreateObject("Scripting.Dictionary")
    Debug.Print "Reading Setup sheet from Q9 to Q" & setupLastRow

    ' Collect unique filenames (Column Q) and departments (Column D)
    For i = 9 To setupLastRow
        filename = Trim(CStr(globalsetupSheet.Cells(i, "Q").value))
        dept = Trim(CStr(globalsetupSheet.Cells(i, "D").value))

        ' Skip blank rows in Column Q
        If filename <> "" And filename <> "0" And Not uniqueFilenames.Exists(filename) Then
            uniqueFilenames.Add filename, Nothing
        End If

        ' Skip blank rows in Column D
        If dept <> "" And dept <> "0" And Not uniqueDepts.Exists(dept) Then
            uniqueDepts.Add dept, Nothing
        End If
    Next i
    Debug.Print "Found " & uniqueFilenames.Count & " unique filenames and " & uniqueDepts.Count & " unique departments"

    ' Debug: Show what's in the dictionaries
    Debug.Print "=== DICTIONARY CONTENTS DEBUG ==="
    Debug.Print "uniqueFilenames.Keys Type: " & TypeName(uniqueFilenames.keys)
    Debug.Print "uniqueFilenames.Keys IsArray: " & IsArray(uniqueFilenames.keys)
    If uniqueFilenames.Count > 0 Then
        Dim debugKey As Variant
        Dim debugCount As Long: debugCount = 0
        Debug.Print "First 5 filenames:"
        For Each debugKey In uniqueFilenames.keys
            Debug.Print "  [" & debugCount & "] " & CStr(debugKey) & " (Type: " & TypeName(debugKey) & ")"
            debugCount = debugCount + 1
            If debugCount >= 5 Then Exit For
        Next debugKey
    End If
    If uniqueDepts.Count > 0 Then
        Dim debugDept As Variant
        debugCount = 0
        Debug.Print "First 5 departments:"
        For Each debugDept In uniqueDepts.keys
            Debug.Print "  [" & debugCount & "] " & CStr(debugDept) & " (Type: " & TypeName(debugDept) & ")"
            debugCount = debugCount + 1
            If debugCount >= 5 Then Exit For
        Next debugDept
    End If
    Debug.Print "=== END DICTIONARY DEBUG ==="

    ' Validate that we have data to process
    If uniqueFilenames.Count = 0 Then
        MsgBox "No valid filenames found in Setup sheet Column Q. Please check your data.", vbExclamation
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If
    If uniqueDepts.Count = 0 Then
        MsgBox "No valid departments found in Setup sheet Column D. Please check your data.", vbExclamation
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Build master dictionary with composite keys (row|dept)
    Set masterDict = CreateObject("Scripting.Dictionary")
    Dim validationError As String
    If Not BuildMasterDictionary(masterDict, uniqueDepts, validationError) Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "P&L export stopped before clearing any values." & vbCrLf & vbCrLf & validationError, _
               vbCritical, "Invalid P&L row mapping"
        Exit Sub
    End If

    globalMasterSheet.Activate
    Debug.Print "Pre-clearing master sheet for clean export..."
    Call ClearPNL

    ' Process each unique filename
    Dim filenameKey As Variant
    For Each filenameKey In uniqueFilenames.keys
        Debug.Print "=== Processing Filename Key ==="
        Debug.Print "Raw filenameKey: " & CStr(filenameKey)
        Debug.Print "filenameKey Type: " & TypeName(filenameKey)
        Debug.Print "filenameKey IsArray: " & IsArray(filenameKey)
        CleanFilename = CleanFilenames(CStr(filenameKey))
        Debug.Print "Cleaned filename variable: '" & CleanFilename & "'"
        Debug.Print "Cleaned filename length: " & Len(CleanFilename)

        ' Skip if cleaned filename is still invalid
        If CleanFilename = "" Or Len(Trim(CleanFilename)) = 0 Then
            Debug.Print "ERROR: Skipping invalid filename: " & CStr(filenameKey)
            GoTo NextFilename
        End If

        ' Create new workbook
        Set newWorkbook = Workbooks.Add

        ' Remove default sheets except one (we'll use the copy method)
        Dim wsCount As Long
        wsCount = newWorkbook.Worksheets.Count

        ' Disable alerts for sheet deletion
        Application.DisplayAlerts = False
        Do While wsCount > 1
            newWorkbook.Worksheets(1).Delete
            wsCount = wsCount - 1
        Loop
        Application.DisplayAlerts = True
        Debug.Print "Processing filename: " & CleanFilename
        Debug.Print "Creating workbook for: " & CleanFilename

        ' Get departments that belong to this filename from Setup sheet
        Dim filenameDepts As Object
        Set filenameDepts = CreateObject("Scripting.Dictionary")

        ' Scan Setup sheet again to find departments for this specific filename
        For i = 9 To setupLastRow
            Dim setupFilename As String, setupDept As String
            setupFilename = Trim(CStr(globalsetupSheet.Cells(i, "Q").value))
            setupDept = Trim(CStr(globalsetupSheet.Cells(i, "D").value))

            ' If this row matches our current filename, add the department
            If setupFilename = CStr(filenameKey) And setupDept <> "" And setupDept <> "0" Then
                If Not filenameDepts.Exists(setupDept) Then
                    filenameDepts.Add setupDept, Nothing
                    Debug.Print "  Found department for " & CleanFilename & ": " & setupDept
                End If
            End If
        Next i
        Debug.Print "Found " & filenameDepts.Count & " departments for filename: " & CleanFilename

        ' Skip if no departments found for this filename
        If filenameDepts.Count = 0 Then
            Debug.Print "WARNING: No departments found for filename: " & CleanFilename & ", skipping workbook creation"
            GoTo NextFilename
        End If

        ' Create sheets for each department that belongs to this filename
        Dim deptKey As Variant
        Dim deptCount As Long: deptCount = 0
        Dim firstSheet As Worksheet: Set firstSheet = Nothing ' Reset for each workbook
        For Each deptKey In filenameDepts.keys
            dept = CStr(deptKey)
            deptCount = deptCount + 1
            Debug.Print "  Creating sheet " & deptCount & " of " & filenameDepts.Count & ": " & dept

            ' FAST: Use optimized copy method instead of manual copying
            Set newSheet = CopySheetStructureFast(globalMasterSheet, newWorkbook, dept)

            ' Track the first sheet created for later selection
            If firstSheet Is Nothing Then
                Set firstSheet = newSheet
                Debug.Print "  Set firstSheet to: " & firstSheet.Name & " (Workbook: " & firstSheet.Parent.Name & ")"
            End If

            ' Update P&L data for this department
            Call UpdatePnLForExportedSheets(newSheet, dept, masterDict)
        Next deptKey

        ' Remove the original default sheet if it still exists
        If newWorkbook.Worksheets.Count > filenameDepts.Count Then
            Application.DisplayAlerts = False
            On Error Resume Next
            newWorkbook.Worksheets("Sheet1").Delete
            On Error GoTo 0
            Application.DisplayAlerts = True
        End If

        ' OPTIMIZATION: Bulk operations after all sheets are created
        Debug.Print "Performing bulk optimizations for workbook: " & CleanFilename
        Call BulkBreakLinksAllSheets(newWorkbook)
        Call BulkRemoveControlsAllSheets(newWorkbook)
        Debug.Print "Completed creating " & filenameDepts.Count & " sheets for workbook: " & CleanFilename

        ' Save workbook
        Dim fullPath As String
        fullPath = savePath & "\" & CleanFilename & ".xlsm"
        Debug.Print "Attempting to save workbook to: " & fullPath

        ' Select the first sheet created before saving (ensure it's in the correct workbook context)
        If Not firstSheet Is Nothing Then
            Debug.Print "About to select firstSheet: " & firstSheet.Name & " (Workbook: " & firstSheet.Parent.Name & ")"
            Debug.Print "Current newWorkbook: " & newWorkbook.Name
            newWorkbook.Activate
            firstSheet.Select
            Debug.Print "Selected first sheet: " & firstSheet.Name & " in workbook: " & CleanFilename
        Else
            Debug.Print "WARNING: firstSheet is Nothing for workbook: " & CleanFilename
        End If
        On Error Resume Next
        newWorkbook.SaveAs fullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        If Err.Number <> 0 Then
            Debug.Print "ERROR saving file: " & fullPath & " - " & Err.Description
            Err.Clear
        Else
            Debug.Print "SUCCESS: Saved file: " & fullPath
        End If
        On Error GoTo 0
        newWorkbook.Close SaveChanges:=False
NextFilename:
    Next filenameKey
    Debug.Print "=== EXPORT SUMMARY ==="
    Debug.Print "Total unique filenames processed: " & uniqueFilenames.Count
    Debug.Print "Save path: " & savePath

    ' Cleanup
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Export completed! Check debug output for details. Files should be in: " & savePath, vbInformation

End Sub

Private Function CleanFilenames(filename As String) As String

    ' Remove or replace characters that are invalid in filenames
    Dim cleanName As String
    Dim invalidChars As String
    Dim i As Long
    Debug.Print "CleanFilenames input: '" & filename & "' (Length: " & Len(filename) & ")"

    ' Handle empty or invalid input
    If Trim(filename) = "" Or filename = "0" Then
        CleanFilenames = "DefaultName_" & Format(Now, "yyyymmdd_hhnnss")
        Debug.Print "CleanFilenames output (default): '" & CleanFilenames & "'"
        Exit Function
    End If
    invalidChars = "\/:*?""<>|"
    cleanName = Trim(filename)
    For i = 1 To Len(invalidChars)
        cleanName = Replace(cleanName, Mid(invalidChars, i, 1), "_")
    Next i

    ' Replace multiple consecutive underscores with single underscore
    Do While InStr(cleanName, "__") > 0
        cleanName = Replace(cleanName, "__", "_")
    Loop

    ' Remove leading/trailing underscores and limit length
    cleanName = Trim(cleanName)
    If Left(cleanName, 1) = "_" Then cleanName = Mid(cleanName, 2)
    If Right(cleanName, 1) = "_" Then cleanName = Left(cleanName, Len(cleanName) - 1)
    If Len(cleanName) > 200 Then
        cleanName = Left(cleanName, 200)
    End If

    ' Ensure we have a valid filename
    If cleanName = "" Then
        cleanName = "DefaultName_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    Debug.Print "CleanFilenames output: '" & cleanName & "'"
    CleanFilenames = cleanName
End Function
Private Function CleanSheetName(sheetName As String, targetWorkbook As Workbook) As String

    ' Clean and ensure unique sheet name within the workbook
    Dim cleanName As String
    Dim invalidChars As String
    Dim i As Long
    Dim counter As Long
    Dim finalName As String
    Debug.Print "CleanSheetName input: '" & sheetName & "'"

    ' Handle empty input
    If Trim(sheetName) = "" Then
        cleanName = "Sheet"
    Else

        ' Remove invalid characters for sheet names: \ / ? * [ ]
        invalidChars = "\/?*[]"
        cleanName = Trim(sheetName)
        For i = 1 To Len(invalidChars)
            cleanName = Replace(cleanName, Mid(invalidChars, i, 1), "_")
        Next i

        ' Replace multiple consecutive underscores with single underscore
        Do While InStr(cleanName, "__") > 0
            cleanName = Replace(cleanName, "__", "_")
        Loop

        ' Remove leading/trailing underscores
        If Left(cleanName, 1) = "_" Then cleanName = Mid(cleanName, 2)
        If Right(cleanName, 1) = "_" Then cleanName = Left(cleanName, Len(cleanName) - 1)

        ' Excel sheet name limit is 31 characters
        If Len(cleanName) > 31 Then
            cleanName = Left(cleanName, 31)
        End If

        ' Ensure we have something
        If cleanName = "" Then cleanName = "Sheet"
    End If

    ' Handle duplicates by adding numbers
    finalName = cleanName
    counter = 1

    ' Check if name already exists in workbook
    Dim ws As Worksheet
    Dim nameExists As Boolean
    Do
        nameExists = False
        For Each ws In targetWorkbook.Worksheets
            If ws.Name = finalName Then
                nameExists = True
                Exit For
            End If
        Next ws
        If nameExists Then
            counter = counter + 1

            ' Ensure we don't exceed 31 characters with the counter
            Dim maxBaseLength As Long
            maxBaseLength = 31 - Len(CStr(counter)) - 1 ' -1 for underscore
            If maxBaseLength < 1 Then maxBaseLength = 1
            If Len(cleanName) > maxBaseLength Then
                finalName = Left(cleanName, maxBaseLength) & "_" & counter
            Else
                finalName = cleanName & "_" & counter
            End If
        End If
    Loop While nameExists
    Debug.Print "CleanSheetName output: '" & finalName & "'"
    CleanSheetName = finalName
End Function
Private Function CopySheetStructureFast(sourceSheet As Worksheet, targetWorkbook As Workbook, newSheetName As String) As Worksheet

    ' FAST: Use Excel's native Move/Copy instead of manual copying
    Debug.Print "  Fast copying sheet structure for: " & newSheetName

    ' Disable alerts to handle prompts automatically
    Application.DisplayAlerts = False

    ' Use Excel's built-in copy (much faster than manual copy)
    sourceSheet.Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)

    ' Get reference to the newly copied sheet and rename it
    Dim newSheet As Worksheet
    Set newSheet = targetWorkbook.Sheets(targetWorkbook.Sheets.Count)

    ' Clean and ensure unique sheet name
    Dim cleanName As String
    cleanName = CleanSheetName(newSheetName, targetWorkbook)

    ' Apply the cleaned name
    On Error Resume Next
    newSheet.Name = cleanName
    If Err.Number <> 0 Then
        Debug.Print "  ERROR renaming sheet to: " & cleanName & " - " & Err.Description
        Err.Clear
    Else
        Debug.Print "  Successfully renamed sheet to: " & cleanName
    End If
    On Error GoTo 0

    ' Add department name to cell I1
    newSheet.Cells(1, "I").value = newSheetName ' Use original department name, not cleaned version
    Debug.Print "  Added department '" & newSheetName & "' to cell I1"

    ' Add additional values from globalSetupSheet to specific cells

    ' Find the department row in globalSetupSheet (data starts from row 9)
    Dim setupRow As Long: setupRow = 0
    Dim i As Long
    For i = 9 To 1000 ' Search in reasonable range starting from row 9
        If globalsetupSheet.Cells(i, "D").value = newSheetName Then
            setupRow = i
            Exit For
        End If
    Next i
    If setupRow > 0 Then

        ' Add values from globalSetupSheet to specific cells
        newSheet.Cells(1, "G").value = globalsetupSheet.Cells(setupRow, "E").value ' G1 = Column E
        newSheet.Cells(2, "G").value = globalsetupSheet.Cells(setupRow, "F").value ' G2 = Column F
        newSheet.Cells(3, "G").value = globalsetupSheet.Cells(setupRow, "G").value ' G3 = Column G
        newSheet.Cells(4, "G").value = globalsetupSheet.Cells(setupRow, "H").value ' G4 = Column H
        Debug.Print "  Added setup sheet values from row " & setupRow & " to G1-G4"
        Debug.Print "    G1 (E" & setupRow & "): " & globalsetupSheet.Cells(setupRow, "E").value
        Debug.Print "    G2 (F" & setupRow & "): " & globalsetupSheet.Cells(setupRow, "F").value
        Debug.Print "    G3 (G" & setupRow & "): " & globalsetupSheet.Cells(setupRow, "G").value
        Debug.Print "    G4 (H" & setupRow & "): " & globalsetupSheet.Cells(setupRow, "H").value
    Else
        Debug.Print "  WARNING: Could not find department '" & newSheetName & "' in globalSetupSheet"
    End If

    ' Collapse all rows and columns, then select G8
    newSheet.Activate

    ' Collapse all outlined rows and columns
    On Error Resume Next
    newSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    On Error GoTo 0

    ' Select range G8
    newSheet.Range("G8").Select
    Debug.Print "  Collapsed outlines and selected G8"

    ' Re-enable alerts
    Application.DisplayAlerts = True
    Debug.Print "  Successfully copied and configured sheet: " & newSheet.Name
    Set CopySheetStructureFast = newSheet
End Function
Private Sub BulkBreakLinksAllSheets(targetWorkbook As Workbook)

    ' FAST: Break all external links in all sheets at once
    Debug.Print "Breaking external links for all sheets in workbook..."
    On Error Resume Next

    ' Method 1: Use Excel's built-in break links if any exist
    Dim links As Variant
    links = targetWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        Dim i As Long
        For i = 1 To UBound(links)
            targetWorkbook.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        Next i
        Debug.Print "Broke " & UBound(links) & " external links"
    End If

    ' Method 2: Convert remaining external formulas to values (faster bulk approach)
    Dim ws As Worksheet
    For Each ws In targetWorkbook.Worksheets
        Dim usedRange As Range
        Set usedRange = ws.usedRange
        If Not usedRange Is Nothing Then

            ' Find and replace external references with values
            usedRange.Replace What:="=[*", Replacement:="#REF!", LookAt:=xlPart
            usedRange.Replace What:="='*", Replacement:="#REF!", LookAt:=xlPart
        End If
    Next ws
    On Error GoTo 0
    Debug.Print "Completed bulk link breaking for all sheets"

End Sub

Private Sub BulkRemoveControlsAllSheets(targetWorkbook As Workbook)

    ' FAST: Remove all form controls and buttons from all sheets at once
    Debug.Print "Removing form controls from all sheets in workbook..."
    Dim ws As Worksheet
    Dim shp As Shape
    Dim totalRemoved As Long: totalRemoved = 0

    ' Disable alerts for shape deletion
    Application.DisplayAlerts = False
    For Each ws In targetWorkbook.Worksheets
        Dim shapesToDelete As Collection
        Set shapesToDelete = New Collection

        ' Collect ALL shapes to delete (including AutoShapes, TextBoxes, etc.)
        For Each shp In ws.Shapes

            ' Delete ALL shape types to ensure clean sheets
            shapesToDelete.Add shp
        Next shp

        ' Delete collected shapes
        Dim shapeItem As Variant
        For Each shapeItem In shapesToDelete
            On Error Resume Next
            shapeItem.Delete
            If Err.Number = 0 Then totalRemoved = totalRemoved + 1
            Err.Clear
            On Error GoTo 0
        Next shapeItem
    Next ws

    ' Re-enable alerts
    Application.DisplayAlerts = True
    Debug.Print "Removed " & totalRemoved & " controls/shapes from all sheets"

End Sub

Private Function BuildMasterDictionary(masterDict As Object, uniqueDepts As Object, _
                                       ByRef validationError As String) As Boolean

    ' Build unified dictionary with composite keys (row|dept) for all departments
    Dim budgetLastRow As Long, actualLastRow As Long
    Dim i As Long, k As Long
    Dim arrBudget As Variant, arrActual As Variant
    Dim AEcol As Long
    Dim dept As Variant
    Dim targetRow As Long
    Dim compositeKey As String
    BuildMasterDictionary = False
    budgetLastRow = GetLastDataRow(globalBudgetFYSheet, "F", 12)
    actualLastRow = GetLastDataRow(globalActualFYSheet, "F", 12)
    globalBudgetFYSheet.Calculate
    globalActualFYSheet.Calculate
    AEcol = 31 ' AE column (target row mapping)
    Debug.Print "Building master dictionary with composite keys..."

    '--- Read all data into arrays ---
    arrBudget = globalBudgetFYSheet.Range("A12:AJ" & budgetLastRow).value
    arrActual = globalActualFYSheet.Range("A12:AJ" & actualLastRow).value

    ' Build department column mapping from Setup sheet
    Dim deptColMapping As Object
    Set deptColMapping = CreateObject("Scripting.Dictionary")
    Dim setupLastRow As Long
    setupLastRow = globalsetupSheet.Cells(globalsetupSheet.Rows.Count, "Q").End(xlUp).row
    Debug.Print "Building department column mapping from Setup sheet..."
    For i = 9 To setupLastRow
        Dim setupDept As String, setupCol As String
        setupDept = Trim(CStr(globalsetupSheet.Cells(i, "D").value)) ' Column D = Department
        setupCol = Trim(CStr(globalsetupSheet.Cells(i, "R").value))  ' Column R = Column reference
        If setupDept <> "" And setupCol <> "" And setupCol <> "0" Then

            ' Convert column letter to number (e.g., "AF" -> 32, "Y" -> 25)
            Dim deptColNum As Long
            deptColNum = Range(setupCol & "1").Column
            deptColMapping(setupDept) = deptColNum
            Debug.Print "  Department: " & setupDept & " -> Column: " & setupCol & " (" & deptColNum & ")"
        End If
    Next i
    Debug.Print "Department column mapping completed with " & deptColMapping.Count & " entries"

    ' Process budget data for each department
    For Each dept In uniqueDepts.keys
        Debug.Print "Processing budget data for department: " & dept

        ' Get the correct column number for this department
        Dim deptCol As Long: deptCol = 25 ' Default fallback
        If deptColMapping.Exists(dept) Then
            deptCol = deptColMapping(dept)
            Debug.Print "  Using column " & deptCol & " for department: " & dept
        Else
            Debug.Print "  WARNING: No column mapping found for department " & dept & ", using default column 25"
        End If

        If Not ValidatePnLTargetRows(arrBudget, globalBudgetFYSheet.Name, deptCol, CStr(dept), _
                                     globalMasterSheet, GetLastUsedRow(globalMasterSheet, 8), validationError) Then Exit Function
        If Not ValidatePnLTargetRows(arrActual, globalActualFYSheet.Name, deptCol, CStr(dept), _
                                     globalMasterSheet, GetLastUsedRow(globalMasterSheet, 8), validationError) Then Exit Function

        For i = 1 To UBound(arrBudget, 1)
            If arrBudget(i, deptCol) = dept And arrBudget(i, AEcol) <> 0 Then
                targetRow = arrBudget(i, AEcol)
                compositeKey = targetRow & "|" & dept & "|GtoR"
                If masterDict.Exists(compositeKey) Then

                    ' Accumulate values (preserve ledger account from first entry)
                    Dim existingRow As Variant: existingRow = masterDict(compositeKey)
                    For k = 10 To 21
                        existingRow(k - 8) = existingRow(k - 8) + arrBudget(i, k)
                    Next k
                    ' Keep existing ledger account (element 14) from first entry
                    masterDict(compositeKey) = existingRow
                Else

                    ' Create new entry with ledger account as 14th element
                    Dim arrNewRow() As Variant
                    ReDim arrNewRow(1 To 14)
                    arrNewRow(1) = targetRow
                    For k = 10 To 21
                        arrNewRow(k - 8) = arrBudget(i, k)
                    Next k
                    arrNewRow(14) = CStr(arrBudget(i, 6)) ' Store ledger account from column F
                    masterDict(compositeKey) = arrNewRow
                End If
            End If
        Next i

        ' Process actual data for each department
        For i = 1 To UBound(arrActual, 1)
            If arrActual(i, deptCol) = dept And arrActual(i, AEcol) <> 0 Then
                targetRow = arrActual(i, AEcol)
                compositeKey = targetRow & "|" & dept & "|WtoAH"
                If masterDict.Exists(compositeKey) Then

                    ' Accumulate values (preserve ledger account from first entry)
                    Dim existingRowActual As Variant: existingRowActual = masterDict(compositeKey)
                    For k = 10 To 21
                        existingRowActual(k - 8) = existingRowActual(k - 8) + arrActual(i, k)
                    Next k
                    ' Keep existing ledger account (element 14) from first entry
                    masterDict(compositeKey) = existingRowActual
                Else

                    ' Create new entry with ledger account as 14th element
                    Dim arrNewRowActual() As Variant
                    ReDim arrNewRowActual(1 To 14)
                    arrNewRowActual(1) = targetRow
                    For k = 10 To 21
                        arrNewRowActual(k - 8) = arrActual(i, k)
                    Next k
                    arrNewRowActual(14) = CStr(arrActual(i, 6)) ' Store ledger account from column F
                    masterDict(compositeKey) = arrNewRowActual
                End If
            End If
        Next i
    Next dept
    Debug.Print "Master dictionary built with " & masterDict.Count & " entries"
    BuildMasterDictionary = True

End Function

Private Sub UpdatePnLForExportedSheets(targetSheet As Worksheet, department As String, masterDict As Object)

    ' Update P&L data for exported sheets using pre-built master dictionary
    Debug.Print "Updating P&L for department: " & department

    Dim targetLastRow As Long
    targetLastRow = GetLastUsedRow(targetSheet, 8)

    ' Clear existing data through the actual copied-template boundary.
    targetSheet.Range(targetSheet.Cells(8, "G"), targetSheet.Cells(targetLastRow, "R")).ClearContents
    targetSheet.Range(targetSheet.Cells(8, "W"), targetSheet.Cells(targetLastRow, "AH")).ClearContents

    ' Build separate dictionaries for this department
    Dim dictGtoR As Object, dictWtoAH As Object
    Set dictGtoR = CreateObject("Scripting.Dictionary")
    Set dictWtoAH = CreateObject("Scripting.Dictionary")

    ' Extract data for this department from master dictionary
    Dim key As Variant
    For Each key In masterDict.keys
        Dim keyParts As Variant
        keyParts = Split(CStr(key), "|")
        If UBound(keyParts) >= 2 Then
            Dim rowNum As Long, dept As String, colType As String
            rowNum = CLng(keyParts(0))
            dept = keyParts(1)
            colType = keyParts(2)
            If dept = department Then
                Dim masterRowData As Variant: masterRowData = masterDict(key)
                
                ' Get ledger account from 14th element for sign inversion logic
                Dim ledgerAccount As String: ledgerAccount = CStr(masterRowData(14))
                Dim signMultiplier As Double: signMultiplier = 1
                
                ' Determine sign multiplier based on ledger account
                If ledgerAccount <> "" And Left(ledgerAccount, 1) <> "9" Then
                    signMultiplier = -1
                    Debug.Print "Export " & colType & " Row " & rowNum & ": Ledger '" & ledgerAccount & "' doesn't start with 9, inverting sign"
                End If
                
                ' Create corrected array without the 14th element (back to 13 elements)
                Dim correctedRowData() As Variant
                ReDim correctedRowData(1 To 13)
                correctedRowData(1) = masterRowData(1) ' Keep target row number
                
                ' Apply sign correction to elements 2-13
                For k = 2 To 13
                    If IsNumeric(masterRowData(k)) Then
                        correctedRowData(k) = masterRowData(k) * signMultiplier
                    Else
                        correctedRowData(k) = masterRowData(k) ' Keep non-numeric values as-is
                    End If
                Next k
                
                ' Store corrected data in appropriate dictionary
                If colType = "GtoR" Then
                    dictGtoR(rowNum) = correctedRowData
                ElseIf colType = "WtoAH" Then
                    dictWtoAH(rowNum) = correctedRowData
                End If
            End If
        End If
    Next key

    ' Debug: Show first 5 entries from extracted dictionaries
    Debug.Print "=== EXTRACTED DICTIONARY DEBUG for " & department & " ==="
    Debug.Print "dictGtoR entries: " & dictGtoR.Count & ", dictWtoAH entries: " & dictWtoAH.Count
    
    ' Debug GtoR dictionary (first 5 entries)
    If dictGtoR.Count > 0 Then
        Debug.Print "=== First 5 GtoR Dictionary Entries ==="
        Dim gtoRKeys As Variant: gtoRKeys = dictGtoR.keys
        Dim sortedGtoRKeys() As Long: ReDim sortedGtoRKeys(0 To UBound(gtoRKeys))
        For i = 0 To UBound(gtoRKeys)
            sortedGtoRKeys(i) = CLng(gtoRKeys(i))
        Next i
        ' Simple bubble sort
        For i = 0 To UBound(sortedGtoRKeys) - 1
            For k = i + 1 To UBound(sortedGtoRKeys)
                If sortedGtoRKeys(i) > sortedGtoRKeys(k) Then
                    Dim temp As Long: temp = sortedGtoRKeys(i)
                    sortedGtoRKeys(i) = sortedGtoRKeys(k)
                    sortedGtoRKeys(k) = temp
                End If
            Next k
        Next i
        ' Print first 5
        For i = 0 To UBound(sortedGtoRKeys)
            If i >= 5 Then Exit For
            Dim gtoRData As Variant: gtoRData = dictGtoR(sortedGtoRKeys(i))
            Debug.Print "  Row " & sortedGtoRKeys(i) & " -> Array Size: " & UBound(gtoRData) & " | Data: " & _
                "[1]=" & gtoRData(1) & " [2]=" & gtoRData(2) & " [3]=" & gtoRData(3) & " [4]=" & gtoRData(4) & " [5]=" & gtoRData(5) & " [6]=" & gtoRData(6) & " (Sign corrected)"
        Next i
    End If
    
    ' Debug WtoAH dictionary (first 5 entries)
    If dictWtoAH.Count > 0 Then
        Debug.Print "=== First 5 WtoAH Dictionary Entries ==="
        Dim wtoAHKeys As Variant: wtoAHKeys = dictWtoAH.keys
        Dim sortedWtoAHKeys() As Long: ReDim sortedWtoAHKeys(0 To UBound(wtoAHKeys))
        For i = 0 To UBound(wtoAHKeys)
            sortedWtoAHKeys(i) = CLng(wtoAHKeys(i))
        Next i
        ' Simple bubble sort
        For i = 0 To UBound(sortedWtoAHKeys) - 1
            For k = i + 1 To UBound(sortedWtoAHKeys)
                If sortedWtoAHKeys(i) > sortedWtoAHKeys(k) Then
                    Dim temp2 As Long: temp2 = sortedWtoAHKeys(i)
                    sortedWtoAHKeys(i) = sortedWtoAHKeys(k)
                    sortedWtoAHKeys(k) = temp2
                End If
            Next k
        Next i
        ' Print first 5
        For i = 0 To UBound(sortedWtoAHKeys)
            If i >= 5 Then Exit For
            Dim wtoAHData As Variant: wtoAHData = dictWtoAH(sortedWtoAHKeys(i))
            Debug.Print "  Row " & sortedWtoAHKeys(i) & " -> Array Size: " & UBound(wtoAHData) & " | Data: " & _
                "[1]=" & wtoAHData(1) & " [2]=" & wtoAHData(2) & " [3]=" & wtoAHData(3) & " [4]=" & wtoAHData(4) & " [5]=" & wtoAHData(5) & " [6]=" & wtoAHData(6) & " (Sign corrected)"
        Next i
    End If
    Debug.Print "=== END EXTRACTED DICTIONARY DEBUG ==="

    ' Add formula rows (rows with AO = "Do not clear")
    Dim i2 As Long, k2 As Long
    For i2 = 8 To GetLastUsedRow(globalMasterSheet, 8)
        If globalMasterSheet.Cells(i2, 41).value = "Do not clear" Then

            ' Add budget formulas to G-R
            Dim arrFormulaBudgetGtoR() As Variant
            ReDim arrFormulaBudgetGtoR(1 To 13)
            arrFormulaBudgetGtoR(1) = i2
            For k2 = 1 To 12
                Dim formulaText As String
                formulaText = globalMasterSheet.Cells(i2, 6 + k2).Formula
                arrFormulaBudgetGtoR(k2 + 1) = IIf(formulaText <> "", formulaText, "")
            Next k2
            dictGtoR(i2) = arrFormulaBudgetGtoR

            ' Add actual formulas to W-AH
            Dim arrFormulaActualWtoAH() As Variant
            ReDim arrFormulaActualWtoAH(1 To 13)
            arrFormulaActualWtoAH(1) = i2
            For k2 = 1 To 12
                formulaText = globalMasterSheet.Cells(i2, 22 + k2).Formula
                arrFormulaActualWtoAH(k2 + 1) = IIf(formulaText <> "", formulaText, "")
            Next k2
            dictWtoAH(i2) = arrFormulaActualWtoAH
        End If
    Next i2

    ' Process Special Rows (VIP Win, Mass Win, Total Slot Win, etc.) using dynamic configuration
    Debug.Print "Processing special rows for department: " & department

    ' Read budget and actual data arrays for special row processing
    Dim budgetLastRow As Long: budgetLastRow = GetLastDataRow(globalBudgetFYSheet, "F", 12)
    Dim actualLastRow As Long: actualLastRow = GetLastDataRow(globalActualFYSheet, "F", 12)
    Dim arrBudget As Variant, arrActual As Variant
    arrBudget = globalBudgetFYSheet.Range("A12:AJ" & budgetLastRow).value
    arrActual = globalActualFYSheet.Range("A12:AJ" & actualLastRow).value

    ' Determine the correct department column for this department
    Dim deptCol As Long: deptCol = 25 ' Default fallback

    ' Build department column mapping from Setup sheet to find correct column
    Dim setupLastRow As Long
    setupLastRow = globalsetupSheet.Cells(globalsetupSheet.Rows.Count, "Q").End(xlUp).row
    Dim setupRowIndex As Long
    For setupRowIndex = 9 To setupLastRow
        Dim setupDept As String, setupCol As String
        setupDept = Trim(CStr(globalsetupSheet.Cells(setupRowIndex, "D").value)) ' Column D = Department
        setupCol = Trim(CStr(globalsetupSheet.Cells(setupRowIndex, "R").value))  ' Column R = Column reference
        If setupDept = department And setupCol <> "" And setupCol <> "0" Then

            ' Convert column letter to number (e.g., "AF" -> 32, "Y" -> 25)
            deptCol = Range(setupCol & "1").Column
            Debug.Print "  Using column " & deptCol & " (" & setupCol & ") for department: " & department
            Exit For
        End If
    Next setupRowIndex
    If deptCol = 25 Then
        Debug.Print "  WARNING: No column mapping found for department " & department & ", using default column 25"
    End If

    ' Process special rows and add results to dictionaries
    Call ProcessSpecialWinRows(department, deptCol, dictGtoR, dictWtoAH, arrBudget, arrActual)

    ' Bulk write to target sheet
    Call BulkWriteToSheet(targetSheet, dictGtoR, "G", "R")
    Call BulkWriteToSheet(targetSheet, dictWtoAH, "W", "AH")
    Call BulkWriteToSheet(targetSheet, dictGtoR, "AV", "BG")
    Debug.Print "P&L update completed for " & department & ". G-R entries: " & dictGtoR.Count & ", W-AH entries: " & dictWtoAH.Count

End Sub


Private Sub BulkWriteToSheet(targetSheet As Worksheet, dataDict As Object, startCol As String, endCol As String)

    ' Generic bulk write routine for any column range
    If dataDict.Count = 0 Then Exit Sub
    Dim keys As Variant, sortedKeys() As Long
    Dim i As Long, j As Long, temp As Long
    keys = dataDict.keys
    ReDim sortedKeys(0 To UBound(keys))

    ' Copy and sort keys
    For i = 0 To UBound(keys)
        sortedKeys(i) = CLng(keys(i))
    Next i

    ' Simple bubble sort
    For i = 0 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Find range
    Dim minRow As Long, maxRow As Long
    minRow = sortedKeys(0)
    maxRow = sortedKeys(UBound(sortedKeys))

    ' Calculate column numbers
    Dim startColNum As Long, endColNum As Long
    startColNum = Range(startCol & "1").Column
    endColNum = Range(endCol & "1").Column

    ' Create bulk array
    Dim rangeRows As Long, rangeCols As Long
    rangeRows = maxRow - minRow + 1
    rangeCols = endColNum - startColNum + 1
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To rangeRows, 1 To rangeCols)

    ' Read existing values
    Dim existingRange As Variant
    existingRange = targetSheet.Range(targetSheet.Cells(minRow, startColNum), targetSheet.Cells(maxRow, endColNum)).value

    ' Copy existing values
    For i = 1 To rangeRows
        For j = 1 To rangeCols
            If IsArray(existingRange) Then
                bulkArray(i, j) = existingRange(i, j)
            Else
                bulkArray(i, j) = existingRange
            End If
        Next j
    Next i

    ' Apply dictionary data
    Dim key As Variant, rowData As Variant
    Dim rowNum As Long, arrayRow As Long
    For Each key In dataDict.keys
        rowData = dataDict(key)
        rowNum = rowData(1)
        arrayRow = rowNum - minRow + 1
        For j = 1 To rangeCols
            Dim cellValue As String
            cellValue = CStr(rowData(j + 1))
            If IsNumeric(cellValue) And cellValue <> "" Then
                bulkArray(arrayRow, j) = CDbl(cellValue)
            ElseIf cellValue <> "" Then
                bulkArray(arrayRow, j) = cellValue
            End If
        Next j
    Next key

    ' Write to sheet
    targetSheet.Range(targetSheet.Cells(minRow, startColNum), targetSheet.Cells(maxRow, endColNum)).value = bulkArray

End Sub

Sub updatePnLbyDept()

    '-----------------------------

    ' updatePnLbyDept: Optimized P&L update by department using single dictionary and bulk writes

    ' Flow: 1) Read data arrays 2) Build unified dictionary by departments 3) Bulk write to master sheet

    '-----------------------------
    Dim dateFilter As String
    Dim lastRow As Long
    Dim i As Long, k As Long
    Dim startTime As Double, elapsedTime As Double
    Dim arrBudget As Variant
    Dim AEcol As Long, dateCol As Long
    Dim budgetLastRow As Long
    Dim validationError As String
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    ' Initialize
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    startTime = Timer
    Call declareGlobal

    lastRow = GetLastUsedRow(globalMasterSheetDept, 8)
    budgetLastRow = GetLastDataRow(globalBudgetFYSheet, "F", 12)
    globalBudgetFYSheet.Calculate
    arrBudget = globalBudgetFYSheet.Range("A12:AJ" & budgetLastRow).value

    If Not ValidatePnLTargetRows(arrBudget, globalBudgetFYSheet.Name, 0, "", _
                                 globalMasterSheetDept, lastRow, validationError) Then GoTo ValidationFailed

    '--- PRE-STEP: Read column formulas BEFORE clearing the sheet ---
    Debug.Print "=== PRE-READING Column Formulas BEFORE Clear ==="
    Dim lastDeptCol As Long
    lastDeptCol = globalMasterSheetDept.Cells(7, globalMasterSheetDept.Columns.Count).End(xlToLeft).Column
    If lastDeptCol < 7 Then lastDeptCol = 7 ' Ensure at least column G

    ' Find columns that have "Formula" in row 6
    Dim formulaColumns As Object
    Set formulaColumns = CreateObject("Scripting.Dictionary")
    For k = 7 To lastDeptCol
        Dim formulaIndicator As String
        formulaIndicator = Trim(CStr(globalMasterSheetDept.Cells(6, k).value))
        If formulaIndicator = "Formula" Then
            formulaColumns.Add k, "Formula_Col_" & k
            Debug.Print "Found formula column: " & k
        End If
    Next k

    ' Pre-read all existing formulas from formula columns before clearing
    Dim dictExistingFormulas As Object
    Set dictExistingFormulas = CreateObject("Scripting.Dictionary")
    Dim preReadColKey As Variant
    For Each preReadColKey In formulaColumns.keys
        Dim preReadColNum As Long: preReadColNum = CLng(preReadColKey)
        Dim preReadRowIndex As Long
        For preReadRowIndex = 8 To lastRow
            Dim existingFormula As String
            existingFormula = globalMasterSheetDept.Cells(preReadRowIndex, preReadColNum).Formula
            If existingFormula <> "" And Left(existingFormula, 1) = "=" Then
                Dim formulaKey As String: formulaKey = preReadColNum & "|" & preReadRowIndex
                dictExistingFormulas.Add formulaKey, existingFormula
                If preReadRowIndex >= 16 And preReadRowIndex <= 21 Then
                    Debug.Print "Pre-read formula at row " & preReadRowIndex & ", col " & preReadColNum & ": " & existingFormula
                End If
            End If
        Next preReadRowIndex
    Next preReadColKey
    Debug.Print "Pre-read " & dictExistingFormulas.Count & " existing formulas before clear"
    Call ClearDeptPNL

    ' Enhanced debug for I2 cell value - check both potential sheets
    Debug.Print "=== Enhanced Date Filter Debug ==="
    Debug.Print "globalMasterSheet Name: " & globalMasterSheet.Name
    Debug.Print "globalMasterSheetDept Name: " & globalMasterSheetDept.Name
    Debug.Print "globalMasterSheet I2 Address: " & globalMasterSheet.Range("I2").Address
    Debug.Print "globalMasterSheet I2 Formula: " & globalMasterSheet.Range("I2").Formula
    Debug.Print "globalMasterSheet I2 Value: '" & globalMasterSheet.Range("I2").value & "'"
    Debug.Print "globalMasterSheet I2 Text: '" & globalMasterSheet.Range("I2").Text & "'"
    Debug.Print "globalMasterSheetDept I2 Value: '" & globalMasterSheetDept.Range("I2").value & "'"
    Debug.Print "globalMasterSheetDept I2 Text: '" & globalMasterSheetDept.Range("I2").Text & "'"

    ' Try getting dateFilter from the correct sheet (globalMasterSheetDept since this function works with it)
    dateFilter = globalMasterSheetDept.Range("I2").value

    ' If still empty, try the original globalMasterSheet
    If IsEmpty(dateFilter) Or IsNull(dateFilter) Or Trim(CStr(dateFilter)) = "" Then
        Debug.Print "dateFilter from globalMasterSheetDept is empty, trying globalMasterSheet..."
        dateFilter = globalMasterSheet.Range("I2").value
    End If

    ' Debug: Show raw dateFilter value and add validation
    Debug.Print "=== Date Filter Debug ==="
    Debug.Print "Raw dateFilter from I2: '" & dateFilter & "'"
    Debug.Print "dateFilter Length: " & Len(dateFilter)
    Debug.Print "dateFilter Type: " & TypeName(dateFilter)

    ' Clean and validate dateFilter
    If IsEmpty(dateFilter) Or IsNull(dateFilter) Then
        Debug.Print "WARNING: dateFilter is Empty or Null, defaulting to 'Jan'"
        dateFilter = "Jan"
    Else
        dateFilter = Trim(CStr(dateFilter))
        If dateFilter = "" Then
            Debug.Print "WARNING: dateFilter is blank after trimming, defaulting to 'Jan'"
            dateFilter = "Jan"
        End If
    End If
    Debug.Print "Final cleaned dateFilter: '" & dateFilter & "'"

    '--- POST-STEP: Read column formulas AFTER clearing the sheet (since ClearDeptPNL preserves formulas) ---
    Debug.Print "=== POST-READING Column Formulas AFTER Clear ==="

    ' Re-read all existing formulas from formula columns after clearing (since ClearDeptPNL doesn't clear formulas)
    Dim postClearColKey As Variant
    For Each postClearColKey In formulaColumns.keys
        Dim postClearColNum As Long: postClearColNum = CLng(postClearColKey)
        Dim postClearRowIndex As Long
        For postClearRowIndex = 8 To lastRow
            Dim postClearFormula As String
            postClearFormula = globalMasterSheetDept.Cells(postClearRowIndex, postClearColNum).Formula
            If postClearFormula <> "" And Left(postClearFormula, 1) = "=" Then
                Dim postClearKey As String: postClearKey = postClearColNum & "|" & postClearRowIndex
                If Not dictExistingFormulas.Exists(postClearKey) Then
                    dictExistingFormulas.Add postClearKey, postClearFormula
                    If postClearRowIndex >= 16 And postClearRowIndex <= 21 Then
                        Debug.Print "Post-clear formula at row " & postClearRowIndex & ", col " & postClearColNum & ": " & postClearFormula
                    End If
                End If
            End If
        Next postClearRowIndex
    Next postClearColKey
    Debug.Print "Total formulas preserved after clear: " & dictExistingFormulas.Count
    AEcol = 31 ' AE column (target row mapping)

    ' Map dateFilter to column number
    Select Case dateFilter
        Case "Jan": dateCol = 10
        Case "Feb": dateCol = 11
        Case "Mar": dateCol = 12
        Case "Apr": dateCol = 13
        Case "May": dateCol = 14
        Case "Jun": dateCol = 15
        Case "Jul": dateCol = 16
        Case "Aug": dateCol = 17
        Case "Sep": dateCol = 18
        Case "Oct": dateCol = 19
        Case "Nov": dateCol = 20
        Case "Dec": dateCol = 21
        Case "Full Year": dateCol = 26
        Case Else:
            Debug.Print "WARNING: Unrecognized dateFilter '" & dateFilter & "', defaulting to Jan (Column 10)"
            dateCol = 10 ' Default to Jan
    End Select
    Debug.Print "=== Final Date Filter Mapping ==="
    Debug.Print "dateFilter: '" & dateFilter & "' -> Column: " & dateCol
    Debug.Print "=== P&L by Dept Update Started for Date Filter: " & dateFilter & " (Column " & dateCol & ") ==="

    '--- STEP 1: Budget data was read and validated before clearing the target sheet ---

    '--- STEP 2: Get department columns from globalMasterSheetDept row 7 ---

    ' lastDeptCol already declared above, reuse it
    lastDeptCol = globalMasterSheetDept.Cells(7, globalMasterSheetDept.Columns.Count).End(xlToLeft).Column
    If lastDeptCol < 7 Then lastDeptCol = 7 ' Ensure at least column G
    Dim deptColumns As Object
    Set deptColumns = CreateObject("Scripting.Dictionary")

    ' Build department column mapping (including columns with "Formula" in row 6)
    For k = 7 To lastDeptCol ' G7 onwards
        Dim deptName As String
        Dim formulaIndicator1 As String
        deptName = Trim(CStr(globalMasterSheetDept.Cells(7, k).value))
        formulaIndicator1 = Trim(CStr(globalMasterSheetDept.Cells(6, k).value))

        ' Include columns with department names OR "Formula" indicator
        If (deptName <> "" And deptName <> "0") Or formulaIndicator1 = "Formula" Then
            If deptName = "" Then deptName = "Formula_Col_" & k ' Give formula columns a name
            deptColumns.Add k, deptName ' Column number -> Department name
            Debug.Print "Found department/formula: " & deptName & " in column " & k
        End If
    Next k

    '--- STEP 3: Build dictionary for departments ---
    Dim dictPnlRowsbyDept As Object
    Set dictPnlRowsbyDept = CreateObject("Scripting.Dictionary")
    Debug.Print "Building dictPnlRowsbyDept (Budget data by department)..."

    ' 3a) Budget data from globalBudgetFYSheet -> department columns
    Dim budgetCount As Long: budgetCount = 0
    Dim rowIndex As Long
    For rowIndex = 1 To UBound(arrBudget, 1)
        If arrBudget(rowIndex, AEcol) <> 0 Then
            Dim targetRow As Long: targetRow = arrBudget(rowIndex, AEcol)
            Dim budgetAmount As Double: budgetAmount = arrBudget(rowIndex, dateCol)

            ' Apply sign inversion for non-revenue accounts (ledger account in column F)
            Dim ledgerAccount As String: ledgerAccount = CStr(arrBudget(rowIndex, 6)) ' Column F
            If Left(ledgerAccount, 1) <> "9" Then
                budgetAmount = budgetAmount * -1
                Debug.Print "Sign inverted for row " & rowIndex & ", ledger: " & ledgerAccount & ", new amount: " & budgetAmount
            End If

            ' Check if this target row already exists in dictionary
            If dictPnlRowsbyDept.Exists(targetRow) Then

                ' Get existing row data
                Dim existingRowData As Variant: existingRowData = dictPnlRowsbyDept(targetRow)

                ' Find which department this row belongs to and add amount
                Dim currentDeptCol As Variant
                For Each currentDeptCol In deptColumns.keys
                    Dim currentDept As String: currentDept = Trim(CStr(arrBudget(rowIndex, 25))) ' Y column = department
                    If currentDept = deptColumns(currentDeptCol) Then

                        ' Calculate correct array index: currentDeptCol is actual column number, need to map to array index
                        Dim deptArrayIndex As Long
                        deptArrayIndex = 1 ' Start from 1 after row number
                        Dim tempDeptCol As Variant
                        For Each tempDeptCol In deptColumns.keys
                            deptArrayIndex = deptArrayIndex + 1
                            If CLng(tempDeptCol) = CLng(currentDeptCol) Then Exit For
                        Next tempDeptCol
                        existingRowData(deptArrayIndex) = existingRowData(deptArrayIndex) + budgetAmount
                        Exit For
                    End If
                Next currentDeptCol
                dictPnlRowsbyDept(targetRow) = existingRowData
            Else

                ' Create new entry
                Dim arrNewRowbyDept() As Variant
                ReDim arrNewRowbyDept(1 To deptColumns.Count + 1) ' Row + department columns
                arrNewRowbyDept(1) = targetRow ' Target row number

                ' Initialize all department columns to 0
                For k = 2 To deptColumns.Count + 1
                    arrNewRowbyDept(k) = 0
                Next k

                ' Add amount to correct department column
                For Each currentDeptCol In deptColumns.keys
                    currentDept = Trim(CStr(arrBudget(rowIndex, 25))) ' Y column = department
                    If currentDept = deptColumns(currentDeptCol) Then

                        ' Calculate correct array index
                        deptArrayIndex = 1 ' Start from 1 after row number
                        For Each tempDeptCol In deptColumns.keys
                            deptArrayIndex = deptArrayIndex + 1
                            If CLng(tempDeptCol) = CLng(currentDeptCol) Then Exit For
                        Next tempDeptCol
                        arrNewRowbyDept(deptArrayIndex) = budgetAmount
                        Exit For
                    End If
                Next currentDeptCol
                dictPnlRowsbyDept(targetRow) = arrNewRowbyDept
            End If
            budgetCount = budgetCount + 1
        End If
    Next rowIndex
    Debug.Print "Budget rows added to department dictionary: " & budgetCount

    ' 3a.1) Process Special Win Rows (VIP Win, Mass Win, Total Slot Win)
    Call ProcessSpecialWinRowsByDept(deptColumns, dictPnlRowsbyDept, arrBudget, dateCol)

    ' 3b) Formulas from globalMasterSheetDept (rows with column D = "Do not clear")
    Dim formulaCount As Long: formulaCount = 0
    Debug.Print "Scanning formula rows from row 8 to " & lastRow & " for 'Do not clear' in column D..."
    Dim formulaRowIndex As Long
    For formulaRowIndex = 8 To lastRow
        If globalMasterSheetDept.Cells(formulaRowIndex, 4).value = "Do not clear" Then
            Dim arrFormulabyDept() As Variant
            ReDim arrFormulabyDept(1 To deptColumns.Count + 1) ' Row + department columns
            arrFormulabyDept(1) = formulaRowIndex ' Target row number

            ' Initialize array index counter
            Dim arrayIndex As Long: arrayIndex = 1
            Dim formulaDeptCol As Variant
            For Each formulaDeptCol In deptColumns.keys
                arrayIndex = arrayIndex + 1
                Dim formulaText As String
                formulaText = globalMasterSheetDept.Cells(formulaRowIndex, formulaDeptCol).Formula

                ' Write formula directly, no prefix needed
                If formulaText <> "" Then
                    arrFormulabyDept(arrayIndex) = formulaText
                Else
                    arrFormulabyDept(arrayIndex) = ""
                End If
            Next formulaDeptCol
            dictPnlRowsbyDept(formulaRowIndex) = arrFormulabyDept
            formulaCount = formulaCount + 1
        End If
    Next formulaRowIndex
    Debug.Print "Formula rows added to department dictionary: " & formulaCount

    '--- STEP 3c: Build separate dictionary for column formulas ---
    Dim dictColumnFormulas As Object
    Set dictColumnFormulas = CreateObject("Scripting.Dictionary")
    Debug.Print "=== Building dictColumnFormulas using pre-read formulas ==="

    ' Use the pre-read existing formulas first
    Dim existingFormulaCount As Long: existingFormulaCount = 0
    Dim existingFormulaKey As Variant
    For Each existingFormulaKey In dictExistingFormulas.keys
        dictColumnFormulas.Add existingFormulaKey, dictExistingFormulas(existingFormulaKey)
        existingFormulaCount = existingFormulaCount + 1

        ' Debug output for rows 16-21
        Dim debugKeyParts() As String: debugKeyParts = Split(CStr(existingFormulaKey), "|")
        Dim debugRowNum As Long: debugRowNum = CLng(debugKeyParts(1))
        If debugRowNum >= 16 And debugRowNum <= 21 Then
            Debug.Print "Restored existing formula at row " & debugRowNum & ": " & dictExistingFormulas(existingFormulaKey)
        End If
    Next existingFormulaKey
    Debug.Print "Restored " & existingFormulaCount & " existing formulas"

    ' For each formula column, generate NEW formulas for rows that have data but no existing formula
    Dim newFormulaCount As Long: newFormulaCount = 0
    Dim formulaColKey1 As Variant
    For Each formulaColKey1 In formulaColumns.keys
        Dim colNum As Long: colNum = CLng(formulaColKey1)

        ' Find the range of data columns (excluding formula columns)
        Dim startColNum As Long, endColNum As Long
        startColNum = 999: endColNum = 0
        Dim dataColKey As Variant
        For Each dataColKey In deptColumns.keys
            If Not formulaColumns.Exists(dataColKey) Then
                If CLng(dataColKey) < startColNum Then startColNum = CLng(dataColKey)
                If CLng(dataColKey) > endColNum Then endColNum = CLng(dataColKey)
            End If
        Next dataColKey

        ' Skip if no valid data columns found
        If startColNum = 999 Or endColNum = 0 Or startColNum > endColNum Then
            Debug.Print "No valid data columns found for formula column " & colNum & ", skipping"
            GoTo NextFormulaColumn
        End If

        ' Convert column numbers to letters
        Dim startCol As String, endCol As String
        startCol = Split(Cells(1, startColNum).Address, "$")(1)
        endCol = Split(Cells(1, endColNum).Address, "$")(1)

        ' For each row that exists in dictPnlRowsbyDept, create formula if not already exists
        Dim dictRowKey As Variant
        For Each dictRowKey In dictPnlRowsbyDept.keys
            Dim formulaRowNum As Long: formulaRowNum = CLng(dictRowKey)
            Dim colFormulaKey As String: colFormulaKey = colNum & "|" & formulaRowNum

            ' Only add if not already exists (from pre-read)
            If Not dictColumnFormulas.Exists(colFormulaKey) Then
                Dim formulaForCell As String
                formulaForCell = "=SUM(" & startCol & formulaRowNum & ":" & endCol & formulaRowNum & ")"
                dictColumnFormulas.Add colFormulaKey, formulaForCell
                newFormulaCount = newFormulaCount + 1
            End If
        Next dictRowKey

        ' Also add formulas for blank rows between data (like 10, 11) within first 50 rows
        Dim blankRowIndex As Long
        For blankRowIndex = 8 To 50
            If Not dictPnlRowsbyDept.Exists(blankRowIndex) Then
                colFormulaKey = colNum & "|" & blankRowIndex

                ' Only add if not already exists (from pre-read)
                If Not dictColumnFormulas.Exists(colFormulaKey) Then

                    ' Check if there are data rows within 2 rows before and after
                    Dim hasRowBefore As Boolean, hasRowAfter As Boolean
                    hasRowBefore = dictPnlRowsbyDept.Exists(blankRowIndex - 1) Or dictPnlRowsbyDept.Exists(blankRowIndex - 2)
                    hasRowAfter = dictPnlRowsbyDept.Exists(blankRowIndex + 1) Or dictPnlRowsbyDept.Exists(blankRowIndex + 2)
                    If hasRowBefore And hasRowAfter Then
                        formulaForCell = "=SUM(" & startCol & blankRowIndex & ":" & endCol & blankRowIndex & ")"
                        dictColumnFormulas.Add colFormulaKey, formulaForCell
                        newFormulaCount = newFormulaCount + 1
                    End If
                End If
            End If
        Next blankRowIndex
NextFormulaColumn:
    Next formulaColKey1
    Debug.Print "Added " & newFormulaCount & " new column formulas"
    Debug.Print "Total column formulas in dictColumnFormulas: " & dictColumnFormulas.Count

    ' Debug: Show formulas for rows 16-21
    Debug.Print "=== dictColumnFormulas Debug for rows 16-21 ==="
    Dim debugFormulaKey As Variant
    For Each debugFormulaKey In dictColumnFormulas.keys
        debugKeyParts = Split(CStr(debugFormulaKey), "|")
        debugRowNum = CLng(debugKeyParts(1))
        If debugRowNum >= 16 And debugRowNum <= 21 Then
            Debug.Print "dictColumnFormulas[" & debugFormulaKey & "] = " & dictColumnFormulas(debugFormulaKey)
        End If
    Next debugFormulaKey

    '--- STEP 4: Debug output for the first 15 entries of dictionary (sorted by row) ---
    Debug.Print "=== Department Dictionary Data (first 15, sorted by row) ==="

    ' Sort keys
    Dim keysDept As Variant, sortedKeysDept() As Long
    Dim countDept As Long: countDept = 0
    keysDept = dictPnlRowsbyDept.keys
    ReDim sortedKeysDept(0 To UBound(keysDept))

    ' Copy and sort
    Dim copyIndex As Long
    For copyIndex = 0 To UBound(keysDept)
        sortedKeysDept(copyIndex) = CLng(keysDept(copyIndex))
    Next copyIndex

    ' Simple bubble sort
    Dim sortOuterIndex As Long, sortInnerIndex As Long
    For sortOuterIndex = 0 To UBound(sortedKeysDept) - 1
        For sortInnerIndex = sortOuterIndex + 1 To UBound(sortedKeysDept)
            If sortedKeysDept(sortOuterIndex) > sortedKeysDept(sortInnerIndex) Then
                Dim sortTemp As Long
                sortTemp = sortedKeysDept(sortOuterIndex)
                sortedKeysDept(sortOuterIndex) = sortedKeysDept(sortInnerIndex)
                sortedKeysDept(sortInnerIndex) = sortTemp
            End If
        Next sortInnerIndex
    Next sortOuterIndex

    ' Print first 15 sorted entries
    Dim debugIndex As Long
    For debugIndex = 0 To UBound(sortedKeysDept)
        If countDept >= 15 Then Exit For
        Dim rowDataDept As Variant: rowDataDept = dictPnlRowsbyDept(sortedKeysDept(debugIndex))
        Dim debugStrDept As String: debugStrDept = "Row " & sortedKeysDept(debugIndex) & " -> Departments: "
        For k = 2 To UBound(rowDataDept) ' Skip row number, show department data
            debugStrDept = debugStrDept & "[" & rowDataDept(k) & "] "
        Next k
        Debug.Print debugStrDept
        countDept = countDept + 1
    Next debugIndex

    '--- STEP 5: Bulk write dictionary to master sheet ---
    Call BulkWriteGtoRbyDept(dictPnlRowsbyDept, deptColumns, dictColumnFormulas)

    ' Finalize
    globalMasterSheetDept.Calculate
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
     Application.DisplayStatusBar = True
    elapsedTime = Timer - startTime
    Debug.Print "=== P&L by Dept Update Completed ==="
    Debug.Print "Total department dictionary entries: " & dictPnlRowsbyDept.Count
    Debug.Print "Date filter: " & dateFilter & " (Column " & dateCol & ")"
    Exit Sub

ValidationFailed:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    MsgBox "P&L by Dept refresh stopped before clearing any values." & vbCrLf & vbCrLf & validationError, _
           vbCritical, "Invalid P&L row mapping"

End Sub

Private Sub BulkWriteGtoRbyDept(dictPnlRowsbyDept As Object, deptColumns As Object, dictColumnFormulas As Object)

    '--- Write department data (budget values + formulas) in bulk ---
    Debug.Print "=== Writing Department Dictionary (Budget + Formulas) ==="
    Debug.Print "Department dictionary entries: " & dictPnlRowsbyDept.Count
    If dictPnlRowsbyDept.Count = 0 Then Exit Sub

    ' Get sorted keys (row numbers)
    Dim dictKeys As Variant, sortedKeys() As Long
    Dim sortIndexI As Long, sortIndexJ As Long, sortTemp As Long
    dictKeys = dictPnlRowsbyDept.keys
    ReDim sortedKeys(0 To UBound(dictKeys))

    ' Copy keys to sortedKeys array
    For sortIndexI = 0 To UBound(dictKeys)
        sortedKeys(sortIndexI) = CLng(dictKeys(sortIndexI))
    Next sortIndexI

    ' Simple bubble sort by row number
    For sortIndexI = 0 To UBound(sortedKeys) - 1
        For sortIndexJ = sortIndexI + 1 To UBound(sortedKeys)
            If sortedKeys(sortIndexI) > sortedKeys(sortIndexJ) Then
                sortTemp = sortedKeys(sortIndexI)
                sortedKeys(sortIndexI) = sortedKeys(sortIndexJ)
                sortedKeys(sortIndexJ) = sortTemp
            End If
        Next sortIndexJ
    Next sortIndexI

    ' Find min and max rows for bulk range
    Dim minRow As Long, maxRow As Long
    minRow = sortedKeys(0)
    maxRow = sortedKeys(UBound(sortedKeys))

    ' Find min and max columns for bulk range
    Dim minCol As Long, maxCol As Long
    Dim deptColKey As Variant
    minCol = 999: maxCol = 0
    For Each deptColKey In deptColumns.keys
        If CLng(deptColKey) < minCol Then minCol = CLng(deptColKey)
        If CLng(deptColKey) > maxCol Then maxCol = CLng(deptColKey)
    Next deptColKey

    ' Create bulk arrays for the entire range
    Dim rangeRows As Long, rangeCols As Long
    rangeRows = maxRow - minRow + 1
    rangeCols = maxCol - minCol + 1
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To rangeRows, 1 To rangeCols)

    ' Read existing values first
    Dim existingRange As Variant
    existingRange = globalMasterSheetDept.Range(globalMasterSheetDept.Cells(minRow, minCol), globalMasterSheetDept.Cells(maxRow, maxCol)).value

    ' Copy existing values to our array
    Dim copyRowIndex As Long, copyColIndex As Long
    For copyRowIndex = 1 To rangeRows
        For copyColIndex = 1 To rangeCols
            If IsArray(existingRange) Then
                bulkArray(copyRowIndex, copyColIndex) = existingRange(copyRowIndex, copyColIndex)
            Else
                bulkArray(copyRowIndex, copyColIndex) = existingRange ' Single cell case
            End If
        Next copyColIndex
    Next copyRowIndex

    ' Apply our dictionary data to the array
    Dim dictRowKey As Variant, rowData As Variant
    Dim rowNum As Long, arrayRow As Long, arrayCol As Long
    Dim cellValue As String
    For Each dictRowKey In dictPnlRowsbyDept.keys
        rowData = dictPnlRowsbyDept(dictRowKey)
        rowNum = rowData(1)
        arrayRow = rowNum - minRow + 1

        ' Apply each department column value
        Dim dataArrayIndex As Long: dataArrayIndex = 1 ' Reset array index counter
        For Each deptColKey In deptColumns.keys
            dataArrayIndex = dataArrayIndex + 1
            arrayCol = CLng(deptColKey) - minCol + 1
            cellValue = CStr(rowData(dataArrayIndex))

            ' Apply values directly - formulas start with =, numbers are numeric
            If IsNumeric(cellValue) And cellValue <> "" Then
                bulkArray(arrayRow, arrayCol) = CDbl(cellValue)
            ElseIf cellValue <> "" Then
                bulkArray(arrayRow, arrayCol) = cellValue ' This includes formulas starting with =
            End If
        Next deptColKey
    Next dictRowKey

    ' Write entire array to sheet in one operation
    globalMasterSheetDept.Range(globalMasterSheetDept.Cells(minRow, minCol), globalMasterSheetDept.Cells(maxRow, maxCol)).value = bulkArray
    Debug.Print "Department bulk write completed for range " & minRow & ":" & maxRow & " (columns " & minCol & " to " & maxCol & ") (" & dictPnlRowsbyDept.Count & " entries)"

    '--- Apply column formulas in BULK by column ---
    Debug.Print "=== Applying Column Formulas in BULK ==="
    Debug.Print "Column formula entries: " & dictColumnFormulas.Count
    If dictColumnFormulas.Count > 0 Then

        ' Group formulas by column number
        Dim formulasByColumn As Object
        Set formulasByColumn = CreateObject("Scripting.Dictionary")
        Dim formulaKeyVar As Variant, formulaValue As String
        Dim colNum As Long, rowNum1 As Long
        Dim keyParts() As String

        ' First pass: Group formulas by column
        For Each formulaKeyVar In dictColumnFormulas.keys
            formulaValue = dictColumnFormulas(formulaKeyVar)
            keyParts = Split(CStr(formulaKeyVar), "|")
            colNum = CLng(keyParts(0))
            rowNum1 = CLng(keyParts(1))

            ' Create column entry if it doesn't exist
            If Not formulasByColumn.Exists(colNum) Then
                Set formulasByColumn(colNum) = CreateObject("Scripting.Dictionary")
            End If

            ' Add row and formula to this column
            formulasByColumn(colNum).Add rowNum1, formulaValue
        Next formulaKeyVar

        ' Second pass: Apply formulas in bulk by column
        Dim columnKey As Variant, columnFormulas As Object
        Dim totalApplied As Long: totalApplied = 0
        For Each columnKey In formulasByColumn.keys
            Set columnFormulas = formulasByColumn(columnKey)
            colNum = CLng(columnKey)

            ' Get sorted row numbers for this column
            Dim rowKeys As Variant, sortedRows() As Long
            Dim sortIdx As Long, sortIdx2 As Long, sortTemp1 As Long
            rowKeys = columnFormulas.keys
            ReDim sortedRows(0 To UBound(rowKeys))

            ' Copy and sort row numbers
            For sortIdx = 0 To UBound(rowKeys)
                sortedRows(sortIdx) = CLng(rowKeys(sortIdx))
            Next sortIdx

            ' Simple bubble sort
            For sortIdx = 0 To UBound(sortedRows) - 1
                For sortIdx2 = sortIdx + 1 To UBound(sortedRows)
                    If sortedRows(sortIdx) > sortedRows(sortIdx2) Then
                        sortTemp1 = sortedRows(sortIdx)
                        sortedRows(sortIdx) = sortedRows(sortIdx2)
                        sortedRows(sortIdx2) = sortTemp1
                    End If
                Next sortIdx2
            Next sortIdx

            ' Apply formulas for this column using optimized bulk ranges
            On Error Resume Next
            Dim currentRowIdx As Long, rangeStart As Long, rangeEnd As Long
            Dim lastFormulaPattern As String, currentFormulaPattern As String
            Dim rangeCount As Long: rangeCount = 0
            For currentRowIdx = 0 To UBound(sortedRows)
                rowNum1 = sortedRows(currentRowIdx)
                formulaValue = columnFormulas(rowNum1)

                ' Extract formula pattern (remove row numbers to detect identical patterns)
                currentFormulaPattern = formulaValue

                ' Replace row numbers with placeholder to detect pattern
                Dim tempPattern As String: tempPattern = formulaValue
                Dim patternRowNum As String: patternRowNum = CStr(rowNum1)

                ' Simple pattern detection - if it contains the row number, it's likely a pattern
                If InStr(tempPattern, patternRowNum) > 0 Then
                    currentFormulaPattern = "ROW_PATTERN"
                Else
                    currentFormulaPattern = formulaValue
                End If

                ' Check if we can extend current range or start new one
                If currentRowIdx = 0 Then

                    ' Start first range
                    rangeStart = rowNum1
                    rangeEnd = rowNum1
                    lastFormulaPattern = currentFormulaPattern
                ElseIf currentFormulaPattern = lastFormulaPattern And rowNum1 = rangeEnd + 1 Then

                    ' Extend current range
                    rangeEnd = rowNum1
                Else

                    ' Apply previous range and start new one
                    Call ApplyFormulaRange(colNum, rangeStart, rangeEnd, columnFormulas, totalApplied, rangeCount)
                    rangeStart = rowNum1
                    rangeEnd = rowNum1
                    lastFormulaPattern = currentFormulaPattern
                End If

                ' Handle last range
                If currentRowIdx = UBound(sortedRows) Then
                    Call ApplyFormulaRange(colNum, rangeStart, rangeEnd, columnFormulas, totalApplied, rangeCount)
                End If
            Next currentRowIdx
            On Error GoTo 0
            Debug.Print "Applied " & columnFormulas.Count & " formulas to column " & colNum & " in " & rangeCount & " range(s)"
        Next columnKey
        Debug.Print "BULK APPLIED " & totalApplied & " column formulas across " & formulasByColumn.Count & " columns"
    End If

End Sub

Private Sub ApplyFormulaRange(colNum As Long, rangeStart As Long, rangeEnd As Long, columnFormulas As Object, ByRef totalApplied As Long, ByRef rangeCount As Long)

    ' Apply formulas to a contiguous range for maximum efficiency
    On Error Resume Next

    ' Validate range
    If rangeStart > rangeEnd Or rangeStart <= 0 Or rangeEnd <= 0 Or colNum <= 0 Then Exit Sub
    If rangeStart > 1048576 Or rangeEnd > 1048576 Or colNum > 16384 Then Exit Sub
    rangeCount = rangeCount + 1
    If rangeStart = rangeEnd Then

        ' Single cell
        Dim singleFormula As String: singleFormula = columnFormulas(rangeStart)
        globalMasterSheetDept.Cells(rangeStart, colNum).Formula = singleFormula
        If Err.Number = 0 Then totalApplied = totalApplied + 1
    Else

        ' Multiple cells - apply one by one but in sequence for better performance
        Dim currentRow As Long
        For currentRow = rangeStart To rangeEnd
            If columnFormulas.Exists(currentRow) Then
                globalMasterSheetDept.Cells(currentRow, colNum).Formula = columnFormulas(currentRow)
                If Err.Number = 0 Then totalApplied = totalApplied + 1
                Err.Clear
            End If
        Next currentRow
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

End Sub
Private Function FindRowByText(searchText As String) As Long

    ' Find row number in globalMasterSheet Column F that exactly matches searchText
    Dim i As Long
    Debug.Print "Searching for exact match of '" & searchText & "' in globalMasterSheet Column F..."
    For i = 8 To GetLastUsedRow(globalMasterSheet, 8)
        If StrComp(Trim(CStr(globalMasterSheet.Cells(i, "F").value)), Trim(searchText), vbTextCompare) = 0 Then
            FindRowByText = i
            Debug.Print "Found exact match of '" & searchText & "' at row " & i
            Exit Function
        End If
    Next i
    FindRowByText = 0 ' Not found
    Debug.Print "WARNING: Exact match for '" & searchText & "' not found in Column F"
End Function

Private Function ReadSpecialRowsConfig() As Object

    ' Read configuration from "Special Rows" sheet and return as dictionary
    Dim configDict As Object
    Set configDict = CreateObject("Scripting.Dictionary")
    
    ' Find the Special Rows sheet
    Dim specialRowsSheet As Worksheet
    Dim ws As Worksheet
    Set specialRowsSheet = Nothing
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Special Rows" Then
            Set specialRowsSheet = ws
            Exit For
        End If
    Next ws
    
    If specialRowsSheet Is Nothing Then
        Set ReadSpecialRowsConfig = configDict
        Exit Function
    End If
    
    ' Read header row (row 6) to get column mappings (OPTIMIZED - removed debug prints)
    Dim headerRow As Long: headerRow = 6
    Dim lastCol As Long: lastCol = specialRowsSheet.Cells(headerRow, specialRowsSheet.Columns.Count).End(xlToLeft).Column
    
    ' Map column headers to column numbers - handle multiple "Include if" columns
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim colIndex As Long
    Dim lastFieldName As String: lastFieldName = ""
    
    For colIndex = 1 To lastCol
        Dim headerText As String: headerText = Trim(CStr(specialRowsSheet.Cells(headerRow, colIndex).value))
        If headerText <> "" Then
            Dim mapKey As String
            
            ' Handle "Include if" columns by creating unique keys based on the previous field
            If LCase(headerText) = "include if" Then
                If lastFieldName <> "" Then
                    mapKey = LCase(lastFieldName) & "_include_if"
                Else
                    GoTo NextColumn
                End If
            Else
                mapKey = LCase(headerText)
                lastFieldName = headerText ' Remember this field name for the next "Include if"
            End If
            
            ' Add to dictionary with unique key
            If Not colMap.Exists(mapKey) Then
                colMap.Add mapKey, colIndex
            End If
        End If
NextColumn:
    Next colIndex
    
    ' Read data rows (starting from row 7)
    Dim dataRow As Long
    For dataRow = 7 To 1000 ' Search reasonable range
        Dim textValue As String: textValue = Trim(CStr(specialRowsSheet.Cells(dataRow, 1).value)) ' First column should be Text
        
        If textValue = "" Then Exit For ' Stop at first empty row
        
        ' Create configuration object for this special row
        Dim configObj As Object
        Set configObj = CreateObject("Scripting.Dictionary")
        
        ' Add text name
        configObj.Add "Text", textValue
        
        ' Read all criteria from this row
        Dim criteriaCount As Long: criteriaCount = 0
        
        ' Company
        If colMap.Exists("company") And colMap.Exists("company_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "company", "company_include_if", "Company", configObj, criteriaCount)
        End If
        
        ' Cost Center
        If colMap.Exists("cost center") And colMap.Exists("cost center_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "cost center", "cost center_include_if", "CostCenter", configObj, criteriaCount)
        End If
        
        ' Customer Segment
        If colMap.Exists("customer segment") And colMap.Exists("customer segment_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "customer segment", "customer segment_include_if", "CustomerSegment", configObj, criteriaCount)
        End If
        
        ' Venue
        If colMap.Exists("venue") And colMap.Exists("venue_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "venue", "venue_include_if", "Venue", configObj, criteriaCount)
        End If
        
        ' Ledger Account
        If colMap.Exists("ledger account") And colMap.Exists("ledger account_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "ledger account", "ledger account_include_if", "LedgerAccount", configObj, criteriaCount)
        End If
        
        ' Revenue Category
        If colMap.Exists("revenue category") And colMap.Exists("revenue category_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "revenue category", "revenue category_include_if", "RevenueCategory", configObj, criteriaCount)
        End If
        
        ' Spend Category as Worktag
        If colMap.Exists("spend category as worktag") And colMap.Exists("spend category as worktag_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "spend category as worktag", "spend category as worktag_include_if", "SpendCategory", configObj, criteriaCount)
        End If
        
        ' US GAAP 20F
        If colMap.Exists("us gaap 20f") And colMap.Exists("us gaap 20f_include_if") Then
            Call AddCriteriaIfExists(specialRowsSheet, dataRow, colMap, "us gaap 20f", "us gaap 20f_include_if", "USGAAP20F", configObj, criteriaCount)
        End If
        
        ' Invert setting
        If colMap.Exists("invert") Then
            Dim invertCol As Long: invertCol = colMap("invert")
            Dim invertValue As String: invertValue = Trim(CStr(specialRowsSheet.Cells(dataRow, invertCol).value))
            If LCase(invertValue) = "yes" Then
                configObj.Add "Invert", True
            Else
                configObj.Add "Invert", False
            End If
        Else
            ' Default to False if Invert column not found
            configObj.Add "Invert", False
        End If
        
        ' Add to main dictionary with unique key handling for duplicates
        Dim uniqueKey As String: uniqueKey = textValue
        Dim duplicateCount As Long: duplicateCount = 1
        
        ' Check if key already exists, if so create unique key
        While configDict.Exists(uniqueKey)
            duplicateCount = duplicateCount + 1
            uniqueKey = textValue & "_" & duplicateCount
        Wend
        
        configDict.Add uniqueKey, configObj
        
    Next dataRow
    
    Set ReadSpecialRowsConfig = configDict

End Function

Private Sub AddCriteriaIfExists(specialRowsSheet As Worksheet, dataRow As Long, colMap As Object, _
                                headerKey As String, includeIfKey As String, configKey As String, configObj As Object, ByRef criteriaCount As Long)
    
    ' Helper function to add criteria to config object if data exists
    If colMap.Exists(headerKey) And colMap.Exists(includeIfKey) Then
        Dim valueCol As Long: valueCol = colMap(headerKey)
        Dim criteriaValue As String: criteriaValue = Trim(CStr(specialRowsSheet.Cells(dataRow, valueCol).value))
        
        If criteriaValue <> "" Then
            ' Get the corresponding "Include if" column using the unique key
            Dim includeIfCol As Long: includeIfCol = colMap(includeIfKey)
            Dim includeIfValue As String: includeIfValue = Trim(CStr(specialRowsSheet.Cells(dataRow, includeIfCol).value))
            
            If includeIfValue <> "" Then
                ' Add both value and operator to config
                configObj.Add configKey & "_Value", criteriaValue
                configObj.Add configKey & "_Operator", includeIfValue
                criteriaCount = criteriaCount + 1
            End If
        End If
    End If
    
End Sub

Private Function EvaluateCriteria(actualValue As String, criteriaValue As String, operator As String) As Boolean

    ' Evaluate if actualValue meets criteria based on operator (OPTIMIZED - removed debug prints)
    Dim actualValueTrimmed As String: actualValueTrimmed = Trim(actualValue)
    Dim criteriaValueTrimmed As String: criteriaValueTrimmed = Trim(criteriaValue)
    
    ' Handle multiple values separated by semicolon
    If InStr(criteriaValueTrimmed, ";") > 0 Then
        Dim criteriaValues() As String: criteriaValues = Split(criteriaValueTrimmed, ";")
        Dim i As Long
        For i = 0 To UBound(criteriaValues)
            If EvaluateSingleCriteria(actualValueTrimmed, Trim(criteriaValues(i)), operator) Then
                EvaluateCriteria = True
                Exit Function
            End If
        Next i
        EvaluateCriteria = False
    Else
        EvaluateCriteria = EvaluateSingleCriteria(actualValueTrimmed, criteriaValueTrimmed, operator)
    End If

End Function

Private Function EvaluateSingleCriteria(actualValue As String, criteriaValue As String, operator As String) As Boolean

    ' Evaluate single criteria comparison
    Select Case LCase(Trim(operator))
        Case "equals to"
            EvaluateSingleCriteria = (LCase(actualValue) = LCase(criteriaValue))
        Case "does not equals to"
            EvaluateSingleCriteria = (LCase(actualValue) <> LCase(criteriaValue))
        Case "contains"
            EvaluateSingleCriteria = (InStr(1, LCase(actualValue), LCase(criteriaValue), vbTextCompare) > 0)
        Case "starts with"
            EvaluateSingleCriteria = (Left(LCase(actualValue), Len(criteriaValue)) = LCase(criteriaValue))
        Case "ends with"
            EvaluateSingleCriteria = (Right(LCase(actualValue), Len(criteriaValue)) = LCase(criteriaValue))
        Case Else
            Debug.Print "WARNING: Unknown operator '" & operator & "', defaulting to False"
            EvaluateSingleCriteria = False
    End Select

End Function
Private Sub ProcessSpecialWinRows(department As String, deptCol As Long, dictGtoR As Object, dictWtoAH As Object, arrBudget As Variant, arrActual As Variant)

    ' Process special rows dynamically based on Special Rows sheet configuration (OPTIMIZED - removed debug, handle duplicates)
    
    ' Read configuration from Special Rows sheet
    Dim specialRowsConfig As Object
    Set specialRowsConfig = ReadSpecialRowsConfig()
    
    If specialRowsConfig.Count = 0 Then Exit Sub
    
    ' Group configurations by row text to handle multiple configs for same row (like Total Slot Win)
    Dim rowGroups As Object
    Set rowGroups = CreateObject("Scripting.Dictionary")
    
    Dim configKey As Variant
    For Each configKey In specialRowsConfig.keys
        Dim configKeyStr As String: configKeyStr = CStr(configKey)
        
        ' Extract base row text (remove _2, _3, etc. suffixes for duplicates)
        Dim rowText As String
        If InStr(configKeyStr, "_") > 0 And IsNumeric(Right(configKeyStr, 1)) Then
            ' This looks like a duplicate (ends with _number)
            Dim lastUnderscorePos As Long: lastUnderscorePos = InStrRev(configKeyStr, "_")
            Dim suffix As String: suffix = Mid(configKeyStr, lastUnderscorePos + 1)
            
            ' Check if suffix is numeric (confirming it's our duplicate naming)
            If IsNumeric(suffix) Then
                rowText = Left(configKeyStr, lastUnderscorePos - 1)
            Else
                rowText = configKeyStr ' Not our duplicate format, use as-is
            End If
        Else
            rowText = configKeyStr ' No underscore or non-numeric suffix, use as-is
        End If
        
        If Not rowGroups.Exists(rowText) Then
            Set rowGroups(rowText) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Add this config to the group (using the original unique key)
        rowGroups(rowText).Add configKeyStr, specialRowsConfig(configKey)
    Next configKey
    
    ' Process each row group (may have multiple configurations per row)
    Dim rowGroupKey As Variant
    For Each rowGroupKey In rowGroups.keys
        Dim rowText1 As String: rowText1 = CStr(rowGroupKey)
        Dim targetRow As Long: targetRow = FindRowByText(rowText1)
        
        If targetRow > 0 Then
            Call ProcessDynamicWinRowMultiple(rowText1, targetRow, department, deptCol, dictGtoR, dictWtoAH, arrBudget, arrActual, rowGroups(rowGroupKey))
        End If
    Next rowGroupKey

End Sub

Private Sub ProcessDynamicWinRow(winType As String, targetRow As Long, department As String, deptCol As Long, dictGtoR As Object, dictWtoAH As Object, arrBudget As Variant, arrActual As Variant, rowConfig As Object)

    ' Process a specific win row using dynamic criteria from configuration (OPTIMIZED - removed debug prints)

    ' Initialize arrays for this win row
    Dim arrWinRowGtoR() As Variant, arrWinRowWtoAH() As Variant
    ReDim arrWinRowGtoR(1 To 13): ReDim arrWinRowWtoAH(1 To 13)
    arrWinRowGtoR(1) = targetRow: arrWinRowWtoAH(1) = targetRow

    ' Initialize all amounts to 0
    Dim k As Long
    For k = 2 To 13
        arrWinRowGtoR(k) = 0: arrWinRowWtoAH(k) = 0
    Next k

    ' Process Budget data (G-R columns)
    Dim budgetCount As Long: budgetCount = 0
    Dim i As Long
    Dim debugRowCount As Long: debugRowCount = 0
    
    For i = 1 To UBound(arrBudget, 1)
        ' Add detailed debugging for first few rows to understand data structure
        If debugRowCount < 5 Then
            Debug.Print "=== SAMPLE BUDGET ROW " & i & " DATA STRUCTURE ==="
            Debug.Print "  Col 1(A): '" & CStr(arrBudget(i, 1)) & "'"
            Debug.Print "  Col 2(B): '" & CStr(arrBudget(i, 2)) & "'"
            Debug.Print "  Col 3(C): '" & CStr(arrBudget(i, 3)) & "'"
            Debug.Print "  Col 4(D): '" & CStr(arrBudget(i, 4)) & "'"
            Debug.Print "  Col 5(E): '" & CStr(arrBudget(i, 5)) & "'"
            Debug.Print "  Col 6(F): '" & CStr(arrBudget(i, 6)) & "'"
            Debug.Print "  Col 7(G): '" & CStr(arrBudget(i, 7)) & "'"
            Debug.Print "  Col 8(H): '" & CStr(arrBudget(i, 8)) & "'"
            Debug.Print "  Col 9(I): '" & CStr(arrBudget(i, 9)) & "'"
            Debug.Print "  Col " & deptCol & "(Dept): '" & CStr(arrBudget(i, deptCol)) & "'"
            debugRowCount = debugRowCount + 1
        End If
        
        If EvaluateRowCriteria(arrBudget, i, department, deptCol, rowConfig) Then
            ' If criteria match, accumulate amounts
            For k = 10 To 21 ' Columns J to U (months)
                arrWinRowGtoR(k - 8) = arrWinRowGtoR(k - 8) + arrBudget(i, k)
            Next k
            budgetCount = budgetCount + 1
        End If
    Next i

    ' Process Actual data (W-AH columns) with same logic
    Dim actualCount As Long: actualCount = 0
    For i = 1 To UBound(arrActual, 1)
        If EvaluateRowCriteria(arrActual, i, department, deptCol, rowConfig) Then
            ' If criteria match, accumulate amounts
            For k = 10 To 21 ' Columns J to U (months)
                arrWinRowWtoAH(k - 8) = arrWinRowWtoAH(k - 8) + arrActual(i, k)
            Next k
            actualCount = actualCount + 1
        End If
    Next i

    ' Conditionally invert the values based on configuration
    Dim shouldInvert As Boolean: shouldInvert = False
    If rowConfig.Exists("Invert") Then
        shouldInvert = rowConfig("Invert")
    End If
    
    If shouldInvert Then
        For k = 2 To 13
            arrWinRowGtoR(k) = arrWinRowGtoR(k) * -1
            arrWinRowWtoAH(k) = arrWinRowWtoAH(k) * -1
        Next k
    End If
    
    dictGtoR(targetRow) = arrWinRowGtoR
    dictWtoAH(targetRow) = arrWinRowWtoAH

End Sub

Private Sub ProcessDynamicWinRowMultiple(winType As String, targetRow As Long, department As String, deptCol As Long, dictGtoR As Object, dictWtoAH As Object, arrBudget As Variant, arrActual As Variant, configGroup As Object)

    ' Process a specific win row with multiple configurations (OPTIMIZED - handles multiple Total Slot Win etc.)
    
    ' Initialize arrays for this win row
    Dim arrWinRowGtoR() As Variant, arrWinRowWtoAH() As Variant
    ReDim arrWinRowGtoR(1 To 13): ReDim arrWinRowWtoAH(1 To 13)
    arrWinRowGtoR(1) = targetRow: arrWinRowWtoAH(1) = targetRow

    ' Initialize all amounts to 0
    Dim k As Long
    For k = 2 To 13
        arrWinRowGtoR(k) = 0: arrWinRowWtoAH(k) = 0
    Next k

    ' Process each configuration in the group
    Dim configIndex As Variant
    For Each configIndex In configGroup.keys
        Dim rowConfig As Object: Set rowConfig = configGroup(configIndex)
        
        ' Process Budget data for this configuration
        Dim i As Long
        For i = 1 To UBound(arrBudget, 1)
            If EvaluateRowCriteria(arrBudget, i, department, deptCol, rowConfig) Then
                ' If criteria match, accumulate amounts
                For k = 10 To 21 ' Columns J to U (months)
                    arrWinRowGtoR(k - 8) = arrWinRowGtoR(k - 8) + arrBudget(i, k)
                Next k
            End If
        Next i

        ' Process Actual data for this configuration
        For i = 1 To UBound(arrActual, 1)
            If EvaluateRowCriteria(arrActual, i, department, deptCol, rowConfig) Then
                ' If criteria match, accumulate amounts
                For k = 10 To 21 ' Columns J to U (months)
                    arrWinRowWtoAH(k - 8) = arrWinRowWtoAH(k - 8) + arrActual(i, k)
                Next k
            End If
        Next i
        
        ' Apply invert setting for this configuration
        Dim shouldInvert As Boolean: shouldInvert = False
        If rowConfig.Exists("Invert") Then
            shouldInvert = rowConfig("Invert")
        End If
        
        If shouldInvert Then
            For k = 2 To 13
                arrWinRowGtoR(k) = arrWinRowGtoR(k) * -1
                arrWinRowWtoAH(k) = arrWinRowWtoAH(k) * -1
            Next k
        End If
    Next configIndex
    
    dictGtoR(targetRow) = arrWinRowGtoR
    dictWtoAH(targetRow) = arrWinRowWtoAH

End Sub

Private Function EvaluateRowCriteria(dataArray As Variant, rowIndex As Long, department As String, deptCol As Long, rowConfig As Object) As Boolean

    ' Evaluate if a data row matches all the criteria in rowConfig (OPTIMIZED - removed debug prints)
    
    ' Check department first (this is always required)
    If Trim(CStr(dataArray(rowIndex, deptCol))) <> department Then
        EvaluateRowCriteria = False
        Exit Function
    End If
    
    ' Pre-extract all data values to avoid repeated array access
    Dim actualCompany As String, actualCostCenter As String, actualCustomerSegment As String
    Dim actualVenue As String, actualLedger As String, actualRevenue As String
    Dim actualSpend As String, actualUSGAAP As String
    
    actualCompany = Trim(CStr(dataArray(rowIndex, 2)))           ' Column B
    actualCostCenter = Trim(CStr(dataArray(rowIndex, 3)))        ' Column C
    actualCustomerSegment = Trim(CStr(dataArray(rowIndex, 4)))   ' Column D
    actualVenue = Trim(CStr(dataArray(rowIndex, 5)))            ' Column E
    actualLedger = Trim(CStr(dataArray(rowIndex, 6)))           ' Column F
    actualRevenue = Trim(CStr(dataArray(rowIndex, 7)))          ' Column G
    actualSpend = Trim(CStr(dataArray(rowIndex, 8)))            ' Column H
    actualUSGAAP = Trim(CStr(dataArray(rowIndex, 9)))           ' Column I
    
    ' Check each configured criteria (optimized with pre-extracted values and early exit)
    If rowConfig.Exists("Company_Value") And rowConfig.Exists("Company_Operator") Then
        If Not EvaluateCriteria(actualCompany, rowConfig("Company_Value"), rowConfig("Company_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("CostCenter_Value") And rowConfig.Exists("CostCenter_Operator") Then
        If Not EvaluateCriteria(actualCostCenter, rowConfig("CostCenter_Value"), rowConfig("CostCenter_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("CustomerSegment_Value") And rowConfig.Exists("CustomerSegment_Operator") Then
        If Not EvaluateCriteria(actualCustomerSegment, rowConfig("CustomerSegment_Value"), rowConfig("CustomerSegment_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("Venue_Value") And rowConfig.Exists("Venue_Operator") Then
        If Not EvaluateCriteria(actualVenue, rowConfig("Venue_Value"), rowConfig("Venue_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("LedgerAccount_Value") And rowConfig.Exists("LedgerAccount_Operator") Then
        If Not EvaluateCriteria(actualLedger, rowConfig("LedgerAccount_Value"), rowConfig("LedgerAccount_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("RevenueCategory_Value") And rowConfig.Exists("RevenueCategory_Operator") Then
        If Not EvaluateCriteria(actualRevenue, rowConfig("RevenueCategory_Value"), rowConfig("RevenueCategory_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("SpendCategory_Value") And rowConfig.Exists("SpendCategory_Operator") Then
        If Not EvaluateCriteria(actualSpend, rowConfig("SpendCategory_Value"), rowConfig("SpendCategory_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("USGAAP20F_Value") And rowConfig.Exists("USGAAP20F_Operator") Then
        If Not EvaluateCriteria(actualUSGAAP, rowConfig("USGAAP20F_Value"), rowConfig("USGAAP20F_Operator")) Then
            EvaluateRowCriteria = False: Exit Function
        End If
    End If
    
    EvaluateRowCriteria = True

End Function

Private Sub ProcessSpecialWinRowsByDept(deptColumns As Object, dictPnlRowsbyDept As Object, arrBudget As Variant, dateCol As Long)

    ' Process special rows dynamically for department-based function (OPTIMIZED - removed debug, handle duplicates)
    
    ' Read configuration from Special Rows sheet
    Dim specialRowsConfig As Object
    Set specialRowsConfig = ReadSpecialRowsConfig()
    
    If specialRowsConfig.Count = 0 Then Exit Sub
    
    ' Group configurations by row text to handle multiple configs for same row (like Total Slot Win)
    Dim rowGroups As Object
    Set rowGroups = CreateObject("Scripting.Dictionary")
    
    Dim configKey As Variant
    For Each configKey In specialRowsConfig.keys
        Dim configKeyStr As String: configKeyStr = CStr(configKey)
        
        ' Extract base row text (remove _2, _3, etc. suffixes for duplicates)
        Dim rowText As String
        If InStr(configKeyStr, "_") > 0 And IsNumeric(Right(configKeyStr, 1)) Then
            ' This looks like a duplicate (ends with _number)
            Dim lastUnderscorePos As Long: lastUnderscorePos = InStrRev(configKeyStr, "_")
            Dim suffix As String: suffix = Mid(configKeyStr, lastUnderscorePos + 1)
            
            ' Check if suffix is numeric (confirming it's our duplicate naming)
            If IsNumeric(suffix) Then
                rowText = Left(configKeyStr, lastUnderscorePos - 1)
            Else
                rowText = configKeyStr ' Not our duplicate format, use as-is
            End If
        Else
            rowText = configKeyStr ' No underscore or non-numeric suffix, use as-is
        End If
        
        If Not rowGroups.Exists(rowText) Then
            Set rowGroups(rowText) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Add this config to the group (using the original unique key)
        rowGroups(rowText).Add configKeyStr, specialRowsConfig(configKey)
    Next configKey
    
    ' Process each row group (may have multiple configurations per row)
    Dim rowGroupKey As Variant
    For Each rowGroupKey In rowGroups.keys
        Dim rowText1 As String: rowText1 = CStr(rowGroupKey)
        Dim targetRow As Long: targetRow = FindRowByText(rowText1)
        
        If targetRow > 0 Then
            Call ProcessDynamicWinRowByDeptMultiple(rowText1, targetRow, deptColumns, dictPnlRowsbyDept, arrBudget, dateCol, rowGroups(rowGroupKey))
        End If
    Next rowGroupKey

End Sub

Private Sub ProcessDynamicWinRowByDept(winType As String, targetRow As Long, deptColumns As Object, dictPnlRowsbyDept As Object, arrBudget As Variant, dateCol As Long, rowConfig As Object)

    ' Process a specific win row for department-based function using dynamic criteria (OPTIMIZED - removed debug prints)

    ' Initialize array for this win row across all departments
    Dim arrWinRowbyDept() As Variant
    ReDim arrWinRowbyDept(1 To deptColumns.Count + 1) ' Row + department columns
    arrWinRowbyDept(1) = targetRow ' Target row number

    ' Initialize all department columns to 0
    Dim k As Long
    For k = 2 To deptColumns.Count + 1
        arrWinRowbyDept(k) = 0
    Next k

    ' Process Budget data for all departments
    Dim totalCount As Long: totalCount = 0
    Dim i As Long
    For i = 1 To UBound(arrBudget, 1)
        If EvaluateRowCriteriaForDept(arrBudget, i, rowConfig) Then
            ' If criteria match, find which department this belongs to and add amount
            Dim currentDeptCol As Variant
            For Each currentDeptCol In deptColumns.keys
                Dim currentDept As String: currentDept = Trim(CStr(arrBudget(i, 25))) ' Y column = department
                If currentDept = deptColumns(currentDeptCol) Then

                    ' Calculate correct array index
                    Dim deptArrayIndex As Long: deptArrayIndex = 1
                    Dim tempDeptCol As Variant
                    For Each tempDeptCol In deptColumns.keys
                        deptArrayIndex = deptArrayIndex + 1
                        If CLng(tempDeptCol) = CLng(currentDeptCol) Then Exit For
                    Next tempDeptCol

                    ' Add the amount for the selected date column
                    arrWinRowbyDept(deptArrayIndex) = arrWinRowbyDept(deptArrayIndex) + arrBudget(i, dateCol)
                    totalCount = totalCount + 1
                    Exit For
                End If
            Next currentDeptCol
        End If
    Next i

    ' Conditionally invert the values based on configuration
    Dim shouldInvert As Boolean: shouldInvert = False
    If rowConfig.Exists("Invert") Then
        shouldInvert = rowConfig("Invert")
    End If
    
    If shouldInvert Then
        For k = 2 To deptColumns.Count + 1
            arrWinRowbyDept(k) = arrWinRowbyDept(k) * -1
        Next k
    End If
    
    dictPnlRowsbyDept(targetRow) = arrWinRowbyDept

End Sub

Private Sub ProcessDynamicWinRowByDeptMultiple(winType As String, targetRow As Long, deptColumns As Object, dictPnlRowsbyDept As Object, arrBudget As Variant, dateCol As Long, configGroup As Object)

    ' Process a specific win row for department-based function with multiple configurations (OPTIMIZED)
    
    ' Initialize array for this win row across all departments
    Dim arrWinRowbyDept() As Variant
    ReDim arrWinRowbyDept(1 To deptColumns.Count + 1) ' Row + department columns
    arrWinRowbyDept(1) = targetRow ' Target row number

    ' Initialize all department columns to 0
    Dim k As Long
    For k = 2 To deptColumns.Count + 1
        arrWinRowbyDept(k) = 0
    Next k

    ' Process each configuration in the group
    Dim configIndex As Variant
    For Each configIndex In configGroup.keys
        Dim rowConfig As Object: Set rowConfig = configGroup(configIndex)
        
        ' Process Budget data for all departments for this configuration
        Dim i As Long
        For i = 1 To UBound(arrBudget, 1)
            If EvaluateRowCriteriaForDept(arrBudget, i, rowConfig) Then
                ' If criteria match, find which department this belongs to and add amount
                Dim currentDeptCol As Variant
                For Each currentDeptCol In deptColumns.keys
                    Dim currentDept As String: currentDept = Trim(CStr(arrBudget(i, 25))) ' Y column = department
                    If currentDept = deptColumns(currentDeptCol) Then

                        ' Calculate correct array index
                        Dim deptArrayIndex As Long: deptArrayIndex = 1
                        Dim tempDeptCol As Variant
                        For Each tempDeptCol In deptColumns.keys
                            deptArrayIndex = deptArrayIndex + 1
                            If CLng(tempDeptCol) = CLng(currentDeptCol) Then Exit For
                        Next tempDeptCol

                        ' Add the amount for the selected date column
                        arrWinRowbyDept(deptArrayIndex) = arrWinRowbyDept(deptArrayIndex) + arrBudget(i, dateCol)
                        Exit For
                    End If
                Next currentDeptCol
            End If
        Next i
        
        ' Apply invert setting for this configuration
        Dim shouldInvert As Boolean: shouldInvert = False
        If rowConfig.Exists("Invert") Then
            shouldInvert = rowConfig("Invert")
        End If
        
        If shouldInvert Then
            For k = 2 To deptColumns.Count + 1
                arrWinRowbyDept(k) = arrWinRowbyDept(k) * -1
            Next k
        End If
    Next configIndex
    
    dictPnlRowsbyDept(targetRow) = arrWinRowbyDept

End Sub

Private Function EvaluateRowCriteriaForDept(dataArray As Variant, rowIndex As Long, rowConfig As Object) As Boolean

    ' Evaluate if a data row matches all the criteria in rowConfig (OPTIMIZED - no department filter, removed debug, early exit)
    
    ' Pre-extract all data values to avoid repeated array access and type conversions
    Dim actualCompany As String, actualCostCenter As String, actualCustomerSegment As String
    Dim actualVenue As String, actualLedger As String, actualRevenue As String
    Dim actualSpend As String, actualUSGAAP As String
    
    actualCompany = CStr(dataArray(rowIndex, 2))           ' Column B
    actualCostCenter = CStr(dataArray(rowIndex, 3))        ' Column C
    actualCustomerSegment = CStr(dataArray(rowIndex, 4))   ' Column D
    actualVenue = CStr(dataArray(rowIndex, 5))            ' Column E
    actualLedger = CStr(dataArray(rowIndex, 6))           ' Column F
    actualRevenue = CStr(dataArray(rowIndex, 7))          ' Column G
    actualSpend = CStr(dataArray(rowIndex, 8))            ' Column H
    actualUSGAAP = CStr(dataArray(rowIndex, 9))           ' Column I
    
    ' Check each configured criteria (optimized with pre-extracted values and early exit)
    If rowConfig.Exists("Company_Value") And rowConfig.Exists("Company_Operator") Then
        If Not EvaluateCriteria(actualCompany, rowConfig("Company_Value"), rowConfig("Company_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("CostCenter_Value") And rowConfig.Exists("CostCenter_Operator") Then
        If Not EvaluateCriteria(actualCostCenter, rowConfig("CostCenter_Value"), rowConfig("CostCenter_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("CustomerSegment_Value") And rowConfig.Exists("CustomerSegment_Operator") Then
        If Not EvaluateCriteria(actualCustomerSegment, rowConfig("CustomerSegment_Value"), rowConfig("CustomerSegment_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("Venue_Value") And rowConfig.Exists("Venue_Operator") Then
        If Not EvaluateCriteria(actualVenue, rowConfig("Venue_Value"), rowConfig("Venue_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("LedgerAccount_Value") And rowConfig.Exists("LedgerAccount_Operator") Then
        If Not EvaluateCriteria(actualLedger, rowConfig("LedgerAccount_Value"), rowConfig("LedgerAccount_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("RevenueCategory_Value") And rowConfig.Exists("RevenueCategory_Operator") Then
        If Not EvaluateCriteria(actualRevenue, rowConfig("RevenueCategory_Value"), rowConfig("RevenueCategory_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("SpendCategory_Value") And rowConfig.Exists("SpendCategory_Operator") Then
        If Not EvaluateCriteria(actualSpend, rowConfig("SpendCategory_Value"), rowConfig("SpendCategory_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    If rowConfig.Exists("USGAAP20F_Value") And rowConfig.Exists("USGAAP20F_Operator") Then
        If Not EvaluateCriteria(actualUSGAAP, rowConfig("USGAAP20F_Value"), rowConfig("USGAAP20F_Operator")) Then
            EvaluateRowCriteriaForDept = False: Exit Function
        End If
    End If
    
    EvaluateRowCriteriaForDept = True

End Function


Sub updateManual()

    ' Optimized version using dictionary-based bulk operations and proportional allocation

    ' Updated Aug 2025
    Call declareGlobal
    Dim startTime As Double
    startTime = Timer
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    globalMasterSheet.Calculate
    globalBudgetFYSheet.Calculate
    If globalMasterSheet.Range("A1").value <> "Individual P&L" Then
        MsgBox "Cannot update a summary."
        Exit Sub
    End If

    ' --- Step 1: Load existing Budget-FY data into dictionary ---
    Dim globalBudgetDict As Object
    Set globalBudgetDict = CreateObject("Scripting.Dictionary")
    Dim lastRowBudget As Long, maxRowBudget As Long
    Dim arrBudgetData As Variant
    Dim i As Long, j As Long
    lastRowBudget = globalBudgetFYSheet.Cells(globalBudgetFYSheet.Rows.Count, "F").End(xlUp).row
    maxRowBudget = lastRowBudget
    If lastRowBudget >= 12 Then

        ' Read all existing data in one bulk operation (A to X columns - includes summaryType)
        arrBudgetData = globalBudgetFYSheet.Range("A12:X" & lastRowBudget).value

        ' Build dictionary from existing data
        For i = 1 To UBound(arrBudgetData, 1)
            If arrBudgetData(i, 1) <> "" Then ' If lookup key exists
                Dim budgetRow(1 To 25) As Variant ' Row + 24 columns (A to X)
                budgetRow(1) = i + 11 ' Actual row number in sheet
                For j = 1 To 24 ' A to X (24 columns)
                    budgetRow(j + 1) = arrBudgetData(i, j)
                Next j
                globalBudgetDict(arrBudgetData(i, 1)) = budgetRow
            End If
        Next i
        Debug.Print "Loaded " & globalBudgetDict.Count & " existing budget rows from Budget-FY sheet"
    Else
        Debug.Print "No existing budget data found - starting with empty dictionary"
        maxRowBudget = 11 ' Start from row 12 when adding new data
    End If

    ' --- Step 2: Find last row in globalMasterSheet column F and collect Manual Changes ---
    Dim lastRowMaster As Long
    lastRowMaster = GetLastUsedRow(globalMasterSheet, 8)

    ' Read globalMasterSheet data in bulk
    Dim arrMasterData As Variant
    arrMasterData = globalMasterSheet.Range("A8:AM" & lastRowMaster).value

    ' Collect Manual Changes rows
    Dim manualChangesDict As Object
    Set manualChangesDict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arrMasterData, 1)
        If arrMasterData(i, 1) <> "" And arrMasterData(i, 21) = "Manual Changes" Then ' Column U = 21
            Dim manualRow(1 To 40) As Variant ' A to AM plus row number
            manualRow(1) = i + 7 ' Actual row number in sheet
            For j = 1 To 39 ' A to AM (39 columns)
                manualRow(j + 1) = arrMasterData(i, j)
            Next j
            manualChangesDict(i + 7) = manualRow ' Use row number as key
        End If
    Next i
    Debug.Print "Found " & manualChangesDict.Count & " Manual Changes rows"
    If manualChangesDict.Count = 0 Then
        MsgBox "No manual changes found to process."
        Exit Sub
    End If

    ' --- Step 3: Get filter criteria from globalMasterSheet ---
    Dim company_code As String, cost_center As String, customer_segment As String, venue As String
    company_code = Trim(CStr(globalMasterSheet.Cells(1, "G").value))
    cost_center = Trim(CStr(globalMasterSheet.Cells(2, "G").value))
    customer_segment = Trim(CStr(globalMasterSheet.Cells(3, "G").value))
    venue = Trim(CStr(globalMasterSheet.Cells(4, "G").value))
    Debug.Print "Filter criteria - Company: '" & company_code & "', Cost Center: '" & cost_center & "', Customer Segment: '" & customer_segment & "', Venue: '" & venue & "'"

    ' --- Step 3.1: Create UsGaap lookup dictionary from filtered budget data ---
    Dim usGaapDict As Object
    Set usGaapDict = CreateObject("Scripting.Dictionary")
    Dim budgetKey As Variant, budgetRowData As Variant
    For Each budgetKey In globalBudgetDict.keys
        budgetRowData = globalBudgetDict(budgetKey)

        ' Check if row matches filter criteria
        Dim isFilterMatch As Boolean
        isFilterMatch = True

        ' Compare filter criteria (empty means include all)
        If company_code <> "" And Trim(CStr(budgetRowData(3))) <> company_code Then isFilterMatch = False ' Column B
        If cost_center <> "" And Trim(CStr(budgetRowData(4))) <> cost_center Then isFilterMatch = False ' Column C
        If customer_segment <> "" And Trim(CStr(budgetRowData(5))) <> customer_segment Then isFilterMatch = False ' Column D
        If venue <> "" And Trim(CStr(budgetRowData(6))) <> venue Then isFilterMatch = False ' Column E
        If isFilterMatch Then
            Dim summaryType As String, usGaap As String
            summaryType = Trim(CStr(budgetRowData(25))) ' Column X = position 25 (summaryType)
            usGaap = Trim(CStr(budgetRowData(10))) ' Column I = position 10 (UsGaap)

            ' Only add if both summaryType and usGaap are not empty
            If summaryType <> "" And usGaap <> "" Then
                If Not usGaapDict.Exists(summaryType) Then
                    usGaapDict.Add summaryType, usGaap
                End If
            End If
        End If
    Next budgetKey
    Debug.Print "Created UsGaap lookup dictionary with " & usGaapDict.Count & " entries"

    ' --- Step 4: Process each Manual Changes row ---
    Dim manualRowKey As Variant, manualRowData As Variant
    Dim processedCount As Long
    processedCount = 0
    For Each manualRowKey In manualChangesDict.keys
        manualRowData = manualChangesDict(manualRowKey)

        ' Extract transaction category components
        Dim ledger_account As String, spend_category As String, revenue_category As String, transactionCategory As String
        ledger_account = Left(Trim(CStr(manualRowData(2))), 6) ' Column A
        spend_category = Trim(CStr(manualRowData(3))) ' Column B
        revenue_category = Trim(CStr(manualRowData(4))) ' Column C
        transactionCategory = ledger_account & spend_category & revenue_category

        ' Extract summaryType from column AM (39)
        Dim manualSummaryType As String
        manualSummaryType = Trim(CStr(manualRowData(40))) ' Column AM = position 40

        ' Extract new amounts (columns G to R)
        Dim newAmounts(1 To 12) As Variant
        For j = 1 To 12
            newAmounts(j) = manualRowData(j + 7) ' G=8, H=9, ..., R=19
        Next j

        ' Apply sign inversion for non-revenue accounts (ledger account in column A)
        If Left(ledger_account, 1) <> "9" Then
            For j = 1 To 12
                If IsNumeric(newAmounts(j)) Then
                    newAmounts(j) = CDbl(newAmounts(j)) * -1
                End If
            Next j
            Debug.Print "Sign inverted for ledger: " & ledger_account & " in manual changes"
        End If

        ' Find matching rows in globalBudgetDict
        Call UpdateBudgetRowsWithAllocation(globalBudgetDict, company_code, cost_center, customer_segment, venue, transactionCategory, newAmounts, processedCount, maxRowBudget, ledger_account, spend_category, revenue_category, usGaapDict, manualSummaryType)
    Next manualRowKey

    ' --- Step 5: Write updated data back to Budget-FY sheet ---
    If processedCount > 0 Then
        Call WriteBulkBudgetDataManual(globalBudgetDict, maxRowBudget)
        Debug.Print "Updated " & processedCount & " budget categories"
    Else
        Debug.Print "No matching budget rows found for manual changes"
    End If

    ' --- Step 6: Apply formulas in bulk ---

    ' Recalculate maxRowBudget before applying formulas
    maxRowBudget = globalBudgetFYSheet.Cells(globalBudgetFYSheet.Rows.Count, "F").End(xlUp).row
    If maxRowBudget >= 12 Then

        ' Copy formulas from V10:AJ10 to V12:AJ[maxRowBudget] using copy-paste
        globalBudgetFYSheet.Range("V10:AJ10").Copy
        globalBudgetFYSheet.Range("V12:AJ" & maxRowBudget).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If

    ' Finalize
    globalMasterSheet.Calculate
    globalBudgetFYSheet.Calculate

    ' Convert formulas to values in columns W12:Y[lastRow] and AC12:AJ[lastRow]
    With globalBudgetFYSheet
        .Range("W12:Y" & maxRowBudget).value = .Range("W12:Y" & maxRowBudget).value
        .Range("AB12:AJ" & maxRowBudget).value = .Range("AB12:AJ" & maxRowBudget).value
    End With
    Call updatePnL(25)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    MsgBox "Manual update completed!" & vbCrLf & "Processed " & processedCount & " categories in " & Format(elapsedTime, "0.00") & " seconds", vbInformation

End Sub


Private Sub UpdateBudgetRowsWithAllocation(globalBudgetDict As Object, company_code As String, cost_center As String, customer_segment As String, venue As String, transactionCategory As String, newAmounts As Variant, ByRef processedCount As Long, ByRef maxRowBudget As Long, ledger_account As String, spend_category As String, revenue_category As String, usGaapDict As Object, manualSummaryType As String)

    ' Find matching budget rows and apply proportional allocation
    Dim matchingRows As Object
    Set matchingRows = CreateObject("Scripting.Dictionary")

    ' Find all rows that match the criteria
    Dim budgetKey As Variant, budgetRowData As Variant
    For Each budgetKey In globalBudgetDict.keys
        budgetRowData = globalBudgetDict(budgetKey)

        ' Check if row matches filter criteria
        Dim isMatch As Boolean
        isMatch = True

        ' Compare filter criteria (empty means include all)
        If company_code <> "" And Trim(CStr(budgetRowData(3))) <> company_code Then isMatch = False ' Column B
        If cost_center <> "" And Trim(CStr(budgetRowData(4))) <> cost_center Then isMatch = False ' Column C
        If customer_segment <> "" And Trim(CStr(budgetRowData(5))) <> customer_segment Then isMatch = False ' Column D
        If venue <> "" And Trim(CStr(budgetRowData(6))) <> venue Then isMatch = False ' Column E

        ' Check transaction category match (exact match required)
        Dim budgetTransactionCategory As String
        budgetTransactionCategory = Left(Trim(CStr(budgetRowData(7))), 6) & Trim(CStr(budgetRowData(8))) & Trim(CStr(budgetRowData(9))) ' F + G + H
        If budgetTransactionCategory <> transactionCategory Then isMatch = False
        If isMatch Then
            matchingRows.Add budgetKey, budgetRowData
        End If
    Next budgetKey
    If matchingRows.Count = 0 Then
        Debug.Print "No matching budget rows found for transaction category: " & transactionCategory
        Debug.Print "Creating new budget row for transaction category: " & transactionCategory

        ' Create a new budget row
        Call CreateNewBudgetRow(globalBudgetDict, company_code, cost_center, customer_segment, venue, transactionCategory, newAmounts, maxRowBudget, ledger_account, spend_category, revenue_category, usGaapDict, manualSummaryType)
        processedCount = processedCount + 1
        Debug.Print "Successfully created new budget row for transaction category: " & transactionCategory
        Exit Sub
    End If
    Debug.Print "Found " & matchingRows.Count & " matching budget rows for transaction category: " & transactionCategory

    ' Calculate allocation for each month
    For monthIndex = 1 To 12
        Dim targetAmount As Double
        targetAmount = CDbl(newAmounts(monthIndex))
        If targetAmount <> 0 Then

            ' Calculate current total for this month across all matching rows
            Dim currentTotal As Double
            currentTotal = 0
            Dim matchKey As Variant
            For Each matchKey In matchingRows.keys
                Dim matchRowData As Variant
                matchRowData = matchingRows(matchKey)

                ' Columns J-U = positions 11-22 in our budgetRow array

                ' Column J is 10th column (A=1,B=2...J=10), so position 11 in array (1=rownum, 2=A, 3=B...11=J)

                ' monthIndex 1-12 should map to positions 11-22 (J-U)
                currentTotal = currentTotal + CDbl(matchRowData(monthIndex + 10)) ' monthIndex 1 -> pos 11 (J), monthIndex 12 -> pos 22 (U)
            Next matchKey

            ' Allocate proportionally if current total > 0, otherwise split equally
            For Each matchKey In matchingRows.keys
                matchRowData = matchingRows(matchKey)
                Dim currentValue As Double

                ' Columns J-U = positions 11-22 in our budgetRow array
                currentValue = CDbl(matchRowData(monthIndex + 10))
                Dim newValue As Double
                If currentTotal <> 0 Then

                    ' Proportional allocation
                    newValue = targetAmount * (currentValue / currentTotal)
                Else

                    ' Equal split if no existing amounts
                    newValue = targetAmount / matchingRows.Count
                End If

                ' Update the row in the main dictionary
                Dim updatedRowData As Variant
                updatedRowData = globalBudgetDict(matchKey)
                updatedRowData(monthIndex + 10) = newValue
                globalBudgetDict(matchKey) = updatedRowData
            Next matchKey
        End If
    Next monthIndex
    processedCount = processedCount + 1
    Debug.Print "Successfully allocated amounts for transaction category: " & transactionCategory & " across " & matchingRows.Count & " rows"

End Sub

Private Sub CreateNewBudgetRow(globalBudgetDict As Object, company_code As String, cost_center As String, customer_segment As String, venue As String, transactionCategory As String, newAmounts As Variant, ByRef maxRowBudget As Long, ledger_account As String, spend_category As String, revenue_category As String, usGaapDict As Object, manualSummaryType As String)

    ' Create a new budget row when no matching row exists
    maxRowBudget = maxRowBudget + 1
    Dim newRowNumber As Long
    newRowNumber = maxRowBudget

    ' Create lookup key for the new row (you may need to adjust this based on your lookup key format)
    Dim lookupKey As String
    lookupKey = company_code & cost_center & customer_segment & venue & ledger_account & spend_category & revenue_category

    ' Create new budget row array
    Dim newBudgetRow(1 To 25) As Variant ' Row + 24 columns (A to X)
    newBudgetRow(1) = newRowNumber ' Row number

    ' Set the basic data columns
    newBudgetRow(2) = lookupKey ' Column A - Lookup key
    newBudgetRow(3) = company_code ' Column B - Company
    newBudgetRow(4) = cost_center ' Column C - Cost Center
    newBudgetRow(5) = customer_segment ' Column D - Customer Segment
    newBudgetRow(6) = venue ' Column E - Venue
    newBudgetRow(7) = ledger_account ' Column F - Ledger Account
    newBudgetRow(8) = spend_category ' Column G - Spend Category
    newBudgetRow(9) = revenue_category ' Column H - Revenue Category

    ' Set UsGaap from usGaapDict based on summaryType
    Dim usGaapValue As String
    usGaapValue = ""
    If manualSummaryType <> "" And usGaapDict.Exists(manualSummaryType) Then
        usGaapValue = usGaapDict(manualSummaryType)
        Debug.Print "Found UsGaap '" & usGaapValue & "' for summaryType '" & manualSummaryType & "'"
    Else
        Debug.Print "WARNING: No UsGaap found for summaryType '" & manualSummaryType & "'"
    End If
    newBudgetRow(10) = usGaapValue ' Column I - UsGaap

    ' Set the monthly amounts (columns J-U = positions 11-22)
    For monthIndex = 1 To 12
        newBudgetRow(monthIndex + 10) = CDbl(newAmounts(monthIndex))
    Next monthIndex

    ' Set other columns to empty/default values for positions 23-24 (columns V-W)
    newBudgetRow(23) = "" ' Column V
    newBudgetRow(24) = "" ' Column W
    
    ' Set summaryType (column X = position 25)
    newBudgetRow(25) = manualSummaryType ' Column X - summaryType

    ' Add the new row to the dictionary
    globalBudgetDict.Add lookupKey, newBudgetRow
    Debug.Print "Created new budget row at row " & newRowNumber & " with key: " & lookupKey

End Sub

Private Sub WriteBulkBudgetDataManual(globalBudgetDict As Object, maxRowBudget As Long)

    ' Write all dictionary data back to Budget-FY sheet in one bulk operation
    If globalBudgetDict.Count = 0 Then Exit Sub

    ' Create bulk array for all data
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To maxRowBudget - 11, 1 To 24) ' Rows 12 to maxRowBudget, Columns A to X

    ' Initialize array with empty values
    Dim i As Long, j As Long
    For i = 1 To maxRowBudget - 11
        For j = 1 To 24
            bulkArray(i, j) = ""
        Next j
    Next i

    ' Fill array with dictionary data
    Dim key As Variant, rowData As Variant
    For Each key In globalBudgetDict.keys
        rowData = globalBudgetDict(key)
        Dim arrayRow As Long
        arrayRow = rowData(1) - 11 ' Convert sheet row to array index
        If arrayRow >= 1 And arrayRow <= UBound(bulkArray, 1) Then
            For j = 1 To 24 ' A to X columns
                bulkArray(arrayRow, j) = rowData(j + 1) ' +1 because rowData(1) is row number
            Next j
        End If
    Next key

    ' Write entire array to sheet in one operation
    If maxRowBudget >= 12 Then
        globalBudgetFYSheet.Range("A12:X" & maxRowBudget).value = bulkArray
        Debug.Print "Bulk write completed for " & globalBudgetDict.Count & " budget rows"
    End If

End Sub













