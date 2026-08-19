Attribute VB_Name = "GlobalVariables"
'Public globalDbPath As String
Public globalsetupSheet As Worksheet
Public globalMasterSheet As Worksheet
Public globalMasterSheetUSD As Worksheet
Public globalMasterSheetDept As Worksheet
Public globalOpexSheet As Worksheet
Public globalPayrollSheet As Worksheet
Public globalBelowEbitda As Worksheet
Public globalAveragesSheet As Worksheet
Public globalopexArray() As String
Public globalOpexRowCount As Long
Public globalpayrollArray() As String
Public globalPayrollRowCount As Long
Public globalMassSheet As Worksheet
Public globalvipSheet As Worksheet
Public globalSlotsSheet As Worksheet
Public globalFnBSheet As Worksheet
Public tableNameCollection As New Collection
Public collectionItem As Variant
Public globalHotelsSheet As Worksheet
Public globalMiscRevSheet As Worksheet
Public globalcompSheet As Worksheet
Public globalSummarySheet As Worksheet
Public globalActualFYSheet As Worksheet
Public globalBudgetFYSheet As Worksheet
Public globalForecastSheet As Worksheet

'Public globalOpexSheet As Worksheet
Public arr() As Variant

'Public globalSheet As Worksheet
Public DeptRow As Long
Public colStart As Long
Public colEnd As Long
Public opexRowStart As Long
Public opexRowEnd As Long
Public payrollRowStart As Long
Public payrollRowEnd As Long
Public massRange_ColEnd As Long
Public lookupValues As Collection
Public matchIndex As Variant
Public appendRow As Long
Public cellNewRow As Range
Public new_RowRange As Range
Public allRanges As New Collection
Public currentCell As Range
Public currentRange As Range
Public totalItems As Long
Public typeTrans As String
Public Categories As New Collection
Public CategoriesItems As Variant
Public CategoriesStr As String
Public Sub declareGlobal()

globalDbPath = ThisWorkbook.Worksheets("Setup").Range("B309")
Set globalsetupSheet = ThisWorkbook.Worksheets("Setup")

'Set globalopexSheet = ThisWorkbook.Worksheets("(A) Opex")
Set globalPayrollSheet = ThisWorkbook.Worksheets("(A) Payroll")
Set globalBelowEbitda = ThisWorkbook.Worksheets("(A) Below EBITDA")

'Set globalAveragesSheet = ThisWorkbook.Worksheets("(A) Averages")
Set globalMasterSheet = ThisWorkbook.Worksheets("P&L")
Set globalMasterSheetUSD = ThisWorkbook.Worksheets("P&L (USD)")
'Set globalMasterSheetDept = ThisWorkbook.Worksheets("P&L by Dept")
Set globalMassSheet = ThisWorkbook.Worksheets("(A) Mass")
Set globalvipSheet = ThisWorkbook.Worksheets("(A) VIP")
Set globalSlotsSheet = ThisWorkbook.Worksheets("(A) Slots")
Set globalFnBSheet = ThisWorkbook.Worksheets("(A) F&B")
Set globalHotelsSheet = ThisWorkbook.Worksheets("(A) Hotels")
Set globalMiscRevSheet = ThisWorkbook.Worksheets("(A) Others")

'Set globalSummarySheet = ThisWorkbook.Worksheets("Parent P&Ls")
Set globalActualFYSheet = ThisWorkbook.Worksheets("Actual-FY")
Set globalBudgetFYSheet = ThisWorkbook.Worksheets("Budget-FY")
Set globalcompSheet = ThisWorkbook.Worksheets("(A) Comps")
Set globalOpexSheet = ThisWorkbook.Worksheets("(A) Opex")
'Set globalForecastSheet = ThisWorkbook.Worksheets("Forecast")
Set allRanges = Nothing

End Sub

Public Sub manualCategoriesItems()

    Application.Calculation = xlManual
    Dim payrollRowRange As Range
    Dim payrollDepRange As Range
    Dim m As Long
    Dim depCell As Range
    Dim colIndex As Long
    Dim payrollValue As String
    Dim targetItem As String
    Dim itemFound As Boolean

    ' Assign the ranges
    Set payrollRowRange = Sheets("(A) Payroll").Range("A12:L300")
    Set payrollDepRange = Sheets("(A) Payroll").Range("O12:O300")

    ' Determine the row count based on non-blank cells in opexDepRange
    globalPayrollRowCount = Application.WorksheetFunction.CountA(payrollDepRange)

    ' Resize the array to match the number of non-blank rows in opexDepRange
    ReDim globalpayrollArray(1 To globalPayrollRowCount)

    ' Loop through each row in opexRowRange and concatenate with opexDepRange (A to L with O)
    For m = 1 To globalPayrollRowCount
        payrollValue = ""
        For colIndex = 1 To 12 ' Columns A to L
            If Not IsEmpty(payrollRowRange.Cells(m, colIndex).value) Then
                payrollValue = payrollValue & payrollRowRange.Cells(m, colIndex).value & " " & payrollDepRange.Cells(m, 1).value
            End If
        Next colIndex

        ' Trim the trailing space, if any
        globalpayrollArray(m) = Trim(payrollValue)
    Next m

    ' Output the array elements
    For m = 1 To globalPayrollRowCount
        If globalpayrollArray(m) <> "" Then
            Debug.Print globalpayrollArray(m)
        End If
    Next m

End Sub

Public Sub newRows(globalSheet As Worksheet)

    ' Updated Aug 2025 - Optimized version using bulk operations and in-memory processing
    Call declareGlobal
    globalSheet.Calculate
    Dim startTime As Double
    startTime = Timer
    Dim totalItems As Long, processedItems As Long
    Dim i As Long, j As Long, k As Long
    Dim lastRowBudget As Long, maxRowBudget As Long
    Dim arrBudgetData As Variant, arrGlobalData As Variant
    
    ' Unfilter Budget-FY sheet if filtered
    If globalBudgetFYSheet.AutoFilterMode Then globalBudgetFYSheet.AutoFilterMode = False

    ' Initialize progress
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Step 1: Loading existing budget data..."
    Progress_Bar.percentDone 0

    ' --- Step 1: Load existing Budget-FY data into dictionary ---
    Dim globalBudgetDict As Object
    Set globalBudgetDict = CreateObject("Scripting.Dictionary")
    lastRowBudget = globalBudgetFYSheet.Cells(globalBudgetFYSheet.Rows.Count, "F").End(xlUp).row
    maxRowBudget = lastRowBudget
    If lastRowBudget >= 12 Then

        ' Read all existing data in one bulk operation (A to U columns)
        arrBudgetData = globalBudgetFYSheet.Range("A12:U" & lastRowBudget).value

        ' Build dictionary from existing data
        For i = 1 To UBound(arrBudgetData, 1)
            If arrBudgetData(i, 1) <> "" Then ' If lookup key exists
                Dim budgetRow(1 To 22) As Variant ' Row + 21 columns (A to U)
                budgetRow(1) = i + 11 ' Actual row number in sheet
                For j = 1 To 21 ' A to U (21 columns)
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

    ' Clear existing data range for fresh bulk write later
    If lastRowBudget >= 12 Then
        globalBudgetFYSheet.Range("A12:AJ" & lastRowBudget).ClearContents
    End If
    Progress_Bar.percentDone 0.2
    Progress_Bar.setStatusText "Step 2: Processing global sheet data..."

    ' --- Step 2: Process globalSheet data and build updates dictionary ---
    Dim globalSheetDict As Object
    Set globalSheetDict = CreateObject("Scripting.Dictionary")
    Dim cellNewRow As Range, cell As Range, new_RowRange As Range
    Dim colStart As Long, colEnd As Long

    ' Read globalSheet data into array for faster processing
    Dim lastRowGlobal As Long
    lastRowGlobal = globalSheet.Cells(globalSheet.Rows.Count, "A").End(xlUp).row
    If lastRowGlobal < 5000 Then lastRowGlobal = 5000
    arrGlobalData = globalSheet.Range("A10:ZZ" & lastRowGlobal).value

    ' Process each row type in memory
    For i = 1 To UBound(arrGlobalData, 1)
        Dim rowType As String
        rowType = arrGlobalData(i, 1) ' Column A
        If rowType = "KPI" Or rowType = "Amount" Or rowType = "Percentage" Or rowType = "Amount with Allocation" Or rowType = "Adjustment" Then
            Dim actualRow As Long
            actualRow = i + 9 ' Convert back to actual sheet row

            ' Debug output for row type detection
            If rowType = "Amount with Allocation" Then
                Debug.Print "Found 'Amount with Allocation' at rowIndex " & i & " (actualRow " & actualRow & ")"
            End If

            ' Handle Adjustment transactions differently
            If rowType = "Adjustment" Then
                colStart = 12 ' Column L
                colEnd = 23   ' Column W

                ' Process each cell in the range for Budget Adjustment
                For j = colStart To colEnd
                    If j <= UBound(arrGlobalData, 2) And arrGlobalData(i, j) <> "" Then
                        Call ProcessAdjustmentTransaction(arrGlobalData, globalBudgetDict, i, j)
                        totalItems = totalItems + 1
                    End If
                Next j
            Else
                ' Handle regular transaction types (KPI, Amount, Percentage, Amount with Allocation)
                
                ' Determine column range based on row type
                If rowType = "Percentage" Or rowType = "Amount with Allocation" Then

                    ' Find column markers in the array
                    colStart = 0
                    colEnd = 0
                    For j = 1 To UBound(arrGlobalData, 2)
                        If arrGlobalData(i, j) = "Column Start>>" Or arrGlobalData(i, j) = "Allocation Start>>" Then
                            colStart = j + 1
                        ElseIf arrGlobalData(i, j) = "<<Column End" Or arrGlobalData(i, j) = "<<Allocation End" Then
                            colEnd = j - 1
                            Exit For
                        End If
                    Next j
                    If colStart = 0 Then colStart = 12
                    If colEnd = 0 Then colEnd = 23
                Else
                    colStart = 12 ' Column L
                    colEnd = 23   ' Column W
                End If

                ' Process each cell in the range for Budget
                For j = colStart To colEnd
                    If j <= UBound(arrGlobalData, 2) And arrGlobalData(i, j) <> "" Then
                        Call ProcessGlobalSheetCell(globalSheetDict, arrGlobalData, globalBudgetDict, i, j, rowType, colStart, actualRow, maxRowBudget)
                        totalItems = totalItems + 1
                    End If
                Next j
            End If
        End If
    Next i
    Progress_Bar.percentDone 0.6
    Progress_Bar.setStatusText "Step 3: Writing updated data back to Budget-FY..."

    ' --- Step 3: Write all data back to Budget-FY in bulk ---
    Call WriteBulkBudgetData(globalBudgetDict, maxRowBudget)

    ' Debug output: Show first 10 dictionary entries

    'Debug.Print "=== DEBUG: First 10 entries in globalBudgetDict ==="

    'Dim debugCount As Long: debugCount = 0

    'Dim debugKey As Variant

    'For Each debugKey In globalBudgetDict.keys

        'If debugCount >= 10 Then Exit For

        'Dim debugRowData As Variant: debugRowData = globalBudgetDict(debugKey)

        'Debug.Print "Key: " & Left(debugKey, 50) & "... | Row: " & debugRowData(1) & " | Company: " & debugRowData(3) & " | J-Col: " & debugRowData(11)

        'debugCount = debugCount + 1

    'Next debugKey

    'Debug.Print "Total dictionary entries: " & globalBudgetDict.Count

    'Debug.Print "=== END DEBUG ==="
    Progress_Bar.percentDone 0.8
    Progress_Bar.setStatusText "Step 4: Applying formulas..."

    ' --- Step 4: Apply formulas in bulk ---

    ' Recalculate maxRowBudget before applying formulas
    maxRowBudget = globalBudgetFYSheet.Cells(globalBudgetFYSheet.Rows.Count, "F").End(xlUp).row
    If maxRowBudget >= 12 Then

        ' Copy formulas from V10:AJ10 to V12:AJ[maxRowBudget] using copy-paste
        globalBudgetFYSheet.Range("V10:AJ10").Copy
        globalBudgetFYSheet.Range("V12:AJ" & maxRowBudget).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If
    Progress_Bar.percentDone 0.9
    Progress_Bar.setStatusText "Step 5: Finalizing..."

    ' Calculate and finalize
    globalBudgetFYSheet.Calculate
    globalSheet.Calculate

    'Set as values the formulas
    Call SetGlobalFYSheetValues
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    Unload Progress_Bar
    MsgBox "Update finished!" & vbCrLf & "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf & "Total items processed: " & totalItems, vbInformation

End Sub

Private Sub ProcessGlobalSheetCell(globalSheetDict As Object, arrGlobalData As Variant, globalBudgetDict As Object, rowIndex As Long, colIndex As Long, mode As String, colStart As Long, actualRow As Long, ByRef maxRowBudget As Long)

    ' Process individual cell from globalSheet and update dictionaries
    Dim company As String, CostCenter As String, customerSegment As String
    Dim venue As String, ledgerAccount As String, revCat As String
    Dim spendCat As String, usGaap As String, lookupValue As String, updateValue As Variant
    Dim targetCol As Long

    ' Extract metadata from the row
    company = arrGlobalData(rowIndex, 2)     ' Column B
    CostCenter = arrGlobalData(rowIndex, 3)  ' Column C
    customerSegment = arrGlobalData(rowIndex, 4) ' Column D
    venue = arrGlobalData(rowIndex, 5)       ' Column E
    usGaap = arrGlobalData(rowIndex, 9)      ' Column I

    ' Process based on mode
    If mode = "Amount" Or mode = "KPI" Then
        ledgerAccount = Left(arrGlobalData(rowIndex, 6), 6) ' Column F
        revCat = arrGlobalData(rowIndex, 7)  ' Column G
        spendCat = arrGlobalData(rowIndex, 8) ' Column H
        lookupValue = company & CostCenter & customerSegment & venue & ledgerAccount & revCat & spendCat & usGaap
        If mode = "Amount" Then
            updateValue = -arrGlobalData(rowIndex, colIndex)
        ElseIf mode = "KPI" Then
            updateValue = arrGlobalData(rowIndex, colIndex)
        End If
        targetCol = colIndex - 2 ' Convert array index to target column (J=10, K=11, etc.)
        Call UpdateDictionaryValue(globalBudgetDict, lookupValue, updateValue, targetCol, company, CostCenter, customerSegment, venue, ledgerAccount, revCat, spendCat, usGaap, maxRowBudget)
    ElseIf mode = "Percentage" Then
        ledgerAccount = Left(arrGlobalData(rowIndex, 6), 6)
        revCat = arrGlobalData(rowIndex, 7)
        spendCat = arrGlobalData(rowIndex, 8)
        lookupValue = company & CostCenter & customerSegment & venue & ledgerAccount & revCat & spendCat & usGaap
        updateValue = -arrGlobalData(rowIndex, colIndex)
        Dim percentageCol As Long
        percentageCol = CInt(colIndex) - CInt(colStart) + 10
        Call UpdateDictionaryValue(globalBudgetDict, lookupValue, updateValue, percentageCol, company, CostCenter, customerSegment, venue, ledgerAccount, revCat, spendCat, usGaap, maxRowBudget)
    ElseIf mode = "Amount with Allocation" Then

        ' Find allocation markers and related rows using corrected logic
        Dim allocationStartCol As Long
        allocationStartCol = 0

        ' Find "Allocation Start>>" in the current row
        For j = 1 To UBound(arrGlobalData, 2)
            If arrGlobalData(rowIndex, j) = "Allocation Start>>" Then
                allocationStartCol = j
                Exit For
            End If
        Next j

        ' Early exit if no allocation marker found
        If allocationStartCol = 0 Then
            Debug.Print "Warning: 'Amount with Allocation' row at rowIndex " & rowIndex & " has no 'Allocation Start>>' marker - skipping allocation processing"
            Exit Sub
        End If

        ' Find ledger, spend, and revenue category rows with separate searches
        Dim ledgerRow As Long, spendRow As Long, revRow As Long
        ledgerRow = 0: spendRow = 0: revRow = 0

        ' Search for LedgerAcct>> marker
        For searchRow = rowIndex - 1 To 1 Step -1
            If searchRow >= 1 And searchRow <= UBound(arrGlobalData, 1) And allocationStartCol >= 1 And allocationStartCol <= UBound(arrGlobalData, 2) Then
                If arrGlobalData(searchRow, allocationStartCol) = "LedgerAcct>>" Then
                    ledgerRow = searchRow
                    Exit For
                End If
            End If
        Next searchRow

        ' Search for SpendCat>> marker
        For searchRow = rowIndex - 1 To 1 Step -1
            If searchRow >= 1 And searchRow <= UBound(arrGlobalData, 1) And allocationStartCol >= 1 And allocationStartCol <= UBound(arrGlobalData, 2) Then
                If arrGlobalData(searchRow, allocationStartCol) = "SpendCat>>" Then
                    spendRow = searchRow
                    Exit For
                End If
            End If
        Next searchRow

        ' Search for RevCat>> marker
        For searchRow = rowIndex - 1 To 1 Step -1
            If searchRow >= 1 And searchRow <= UBound(arrGlobalData, 1) And allocationStartCol >= 1 And allocationStartCol <= UBound(arrGlobalData, 2) Then
                If arrGlobalData(searchRow, allocationStartCol) = "RevCat>>" Then
                    revRow = searchRow
                    Exit For
                End If
            End If
        Next searchRow

        ' Debug output for allocation search
        Debug.Print "Allocation search for rowIndex " & rowIndex & ": allocationStartCol=" & allocationStartCol & ", ledgerRow=" & ledgerRow & ", spendRow=" & spendRow & ", revRow=" & revRow

        ' Only process if all required rows were found
        If ledgerRow > 0 And spendRow > 0 And revRow > 0 And allocationStartCol > 0 Then

            ' Process allocation for 12 months
            For i = 0 To 11
                If ledgerRow <= UBound(arrGlobalData, 1) And spendRow <= UBound(arrGlobalData, 1) And revRow <= UBound(arrGlobalData, 1) And colIndex <= UBound(arrGlobalData, 2) Then
                    ledgerAccount = Left(arrGlobalData(ledgerRow, colIndex), 6)
                    spendCat = arrGlobalData(spendRow, colIndex)
                    revCat = arrGlobalData(revRow, colIndex)
                    lookupValue = company & CostCenter & customerSegment & venue & ledgerAccount & revCat & spendCat & usGaap

                    ' Check bounds for allocation data
                    If (12 + i) <= UBound(arrGlobalData, 2) Then
                        updateValue = (arrGlobalData(rowIndex, 12 + i) * -arrGlobalData(rowIndex, colIndex))
                        Dim allocationCol As Long
                        allocationCol = 10 + i
                        Call UpdateDictionaryValue(globalBudgetDict, lookupValue, updateValue, allocationCol, company, CostCenter, customerSegment, venue, ledgerAccount, revCat, spendCat, usGaap, maxRowBudget)
                    End If
                End If
            Next i
        Else
            Debug.Print "Warning: Could not find all required allocation rows for rowIndex " & rowIndex & " (allocationStartCol=" & allocationStartCol & ", ledgerRow=" & ledgerRow & ", spendRow=" & spendRow & ", revRow=" & revRow & ")"
        End If
    End If

End Sub

Private Sub ProcessAdjustmentTransaction(arrGlobalData As Variant, globalBudgetDict As Object, rowIndex As Long, colIndex As Long)

    ' Process Adjustment transaction type - updates existing rows only, never creates new rows
    Dim company As String, CostCenter As String, customerSegment As String
    Dim venue As String, ledgerAccount As String, revCat As String
    Dim spendCat As String, usGaap As String, updateValue As Variant
    Dim targetCol As Long

    ' Extract metadata from the row - some may be empty for partial matching
    company = arrGlobalData(rowIndex, 2)     ' Column B
    CostCenter = arrGlobalData(rowIndex, 3)  ' Column C
    customerSegment = arrGlobalData(rowIndex, 4) ' Column D
    venue = arrGlobalData(rowIndex, 5)       ' Column E
    ledgerAccount = Left(arrGlobalData(rowIndex, 6), 6) ' Column F
    revCat = arrGlobalData(rowIndex, 7)      ' Column G
    spendCat = arrGlobalData(rowIndex, 8)    ' Column H
    usGaap = arrGlobalData(rowIndex, 9)      ' Column I
    updateValue = arrGlobalData(rowIndex, colIndex)
    targetCol = colIndex - 2 ' Convert array index to target column (J=10, K=11, etc.)

    ' Find all matching existing entries in globalBudgetDict and update them
    Dim key As Variant, rowData As Variant
    For Each key In globalBudgetDict.keys
        rowData = globalBudgetDict(key)

        ' Check if this row matches the adjustment criteria
        Dim isMatch As Boolean
        isMatch = True

        ' Only check fields that have values in the adjustment row
        If company <> "" And rowData(3) <> company Then isMatch = False
        If CostCenter <> "" And rowData(4) <> CostCenter Then isMatch = False
        If customerSegment <> "" And rowData(5) <> customerSegment Then isMatch = False
        If venue <> "" And rowData(6) <> venue Then isMatch = False
        If ledgerAccount <> "" And rowData(7) <> ledgerAccount Then isMatch = False
        If revCat <> "" And rowData(8) <> revCat Then isMatch = False
        If spendCat <> "" And rowData(9) <> spendCat Then isMatch = False
        If usGaap <> "" And rowData(10) <> usGaap Then isMatch = False

        ' If match found, update the target column
        If isMatch Then
            rowData(targetCol + 1) = updateValue ' +1 because row(1) is row number, data starts at row(2)
            globalBudgetDict(key) = rowData
            Debug.Print "Adjustment applied to key: " & Left(key, 30) & "... | TargetCol=" & targetCol & " | Value=" & updateValue
        End If
    Next key

End Sub

Private Sub UpdateDictionaryValue(globalBudgetDict As Object, lookupValue As String, updateValue As Variant, targetCol As Long, company As String, CostCenter As String, customerSegment As String, venue As String, ledgerAccount As String, revCat As String, spendCat As String, usGaap As String, ByRef maxRowBudget As Long)

    ' Update or create dictionary entry
    If globalBudgetDict.Exists(lookupValue) Then

        ' Update existing entry
        Dim existingRow As Variant
        existingRow = globalBudgetDict(lookupValue)
        existingRow(targetCol + 1) = updateValue ' +1 because row(1) is row number, data starts at row(2)
        globalBudgetDict(lookupValue) = existingRow

        ' Debug output for first 10 matches

        'Static matchCount As Long

        'If matchCount < 10 Then

            'matchCount = matchCount + 1

            'Debug.Print "MATCH " & matchCount & ": Key=" & Left(lookupValue, 30) & "... | Row=" & existingRow(1) & " | TargetCol=" & targetCol & " | Value=" & updateValue

        'End If
    Else

        ' Create new entry
        If updateValue <> 0 Then
            maxRowBudget = maxRowBudget + 1
            Dim newRow(1 To 22) As Variant ' Row + 21 columns (A to U)
            newRow(1) = maxRowBudget ' Row number
            newRow(2) = lookupValue  ' Column A
            newRow(3) = company      ' Column B
            newRow(4) = CostCenter   ' Column C
            newRow(5) = customerSegment ' Column D
            newRow(6) = venue        ' Column E
            newRow(7) = ledgerAccount ' Column F
            newRow(8) = revCat       ' Column G
            newRow(9) = spendCat     ' Column H
            newRow(10) = usGaap      ' Column I

            ' Initialize J to U (columns 10-21) with zeros
            For i = 11 To 22
                newRow(i) = 0
            Next i
            newRow(targetCol + 1) = updateValue
            globalBudgetDict(lookupValue) = newRow

            ' Debug output for first 10 new rows

            'Static newRowCount As Long

            'If newRowCount < 10 Then

                'newRowCount = newRowCount + 1

                'Debug.Print "NEW " & newRowCount & ": Key=" & Left(lookupValue, 30) & "... | Row=" & maxRowBudget & " | TargetCol=" & targetCol & " | Value=" & updateValue & " | Company=" & company

            'End If
        End If
    End If

End Sub

Private Sub WriteBulkBudgetData(globalBudgetDict As Object, maxRowBudget As Long)

    ' Write all dictionary data back to Budget-FY sheet in one bulk operation
    If globalBudgetDict.Count = 0 Then Exit Sub

    ' Create bulk array for all data
    Dim bulkArray() As Variant
    ReDim bulkArray(1 To maxRowBudget - 11, 1 To 21) ' Rows 12 to maxRowBudget, Columns A to U

    ' Initialize array with empty values
    For i = 1 To maxRowBudget - 11
        For j = 1 To 21
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
            For j = 2 To 22 ' Skip row number, get columns A to U
                bulkArray(arrayRow, j - 1) = rowData(j)
            Next j
        End If
    Next key

    ' Write entire array to sheet in one operation
    If maxRowBudget >= 12 Then
        globalBudgetFYSheet.Range("A12:U" & maxRowBudget).value = bulkArray
    End If

End Sub



Sub clearBudget()

'
    Sheets("Budget-FY").Select
    Sheets("Budget-FY").Range("A12:AT10000").ClearContents

End Sub

Public Sub UpdateBudgetFYWithAllocations(filter5YP As String, sourceSheet As Worksheet)

    ' Optimized version using dictionary lookup and reduced worksheet operations

    ' Updated Jul 2025
    Call declareGlobal
    Dim wsActual As Worksheet
    Dim wsBudget As Worksheet
    Dim wsExclusion As Worksheet
    Dim lastRowActual As Long, lastRowBudget As Long, lastRowExclusion As Long
    Dim dictBudgetKeys As Object, dictDept As Object
    Dim i As Long, j As Long, k As Long
    Dim ytdAverage As Variant, department As String
    Dim cell As Range
    Dim allocArr As Variant
    Dim processedItems As Long, totalItems As Long
    Dim startTime As Double, elapsedTime As Double
    Dim arrBudgetKeys As Variant
    startTime = Timer
    Set wsActual = ThisWorkbook.Sheets("Actual-FY")
    Set wsBudget = ThisWorkbook.Sheets("Budget-FY")
    Set wsExclusion = ThisWorkbook.Sheets("Exclusions")
    sourceSheet.Calculate
    wsActual.Calculate
    wsBudget.Calculate
    wsExclusion.Calculate

    ' --- Load exclusionArr from UpdateExclusion sheet ---
    lastRowExclusion = wsExclusion.Cells(wsExclusion.Rows.Count, "A").End(xlUp).row
    Dim tempArrExcl() As Variant
    Dim exclusionArr() As Variant
    Dim exclusionCount As Long
    exclusionCount = 0
    For i = 2 To lastRowExclusion
        If wsExclusion.Cells(i, 1).value = filter5YP Then
            exclusionCount = exclusionCount + 1
            ReDim Preserve tempArrExcl(1 To exclusionCount)
            Dim rowArr(1 To 9) As Variant
            For j = 1 To 9
                If IsEmpty(wsExclusion.Cells(i, j).value) Or wsExclusion.Cells(i, j).value = "" Then
                    rowArr(j) = ""
                Else
                    rowArr(j) = wsExclusion.Cells(i, j).value
                End If
            Next j
            tempArrExcl(exclusionCount) = rowArr
        End If
    Next i
    If exclusionCount > 0 Then
        ReDim exclusionArr(1 To exclusionCount, 1 To 9)
        For i = 1 To exclusionCount
            For j = 1 To 9
                exclusionArr(i, j) = tempArrExcl(i)(j)
            Next j
        Next i
    Else
        ReDim exclusionArr(1 To 1, 1 To 9)
        For j = 1 To 9
            exclusionArr(1, j) = ""
        Next j
    End If

    ' --- Step 1: Build dictionary of existing keys in Budget-FY for fast lookup ---
    lastRowBudget = wsBudget.Cells(wsBudget.Rows.Count, "A").End(xlUp).row
    Set dictBudgetKeys = CreateObject("Scripting.Dictionary")
    If lastRowBudget >= 2 Then
        arrBudgetKeys = wsBudget.Range("A2:A" & lastRowBudget).value
        For i = 1 To UBound(arrBudgetKeys, 1)
            If Not IsEmpty(arrBudgetKeys(i, 1)) Then
                dictBudgetKeys(arrBudgetKeys(i, 1)) = True
            End If
        Next i
    End If

    ' --- Step 2: Copy new rows from Actual-FY to Budget-FY if not already present ---
    lastRowActual = wsActual.Cells(wsActual.Rows.Count, "A").End(xlUp).row
    wsActual.Range("A1:Y" & lastRowActual).AutoFilter Field:=23, Criteria1:=filter5YP
    On Error Resume Next
    totalItems = wsActual.Range("A12:A" & lastRowActual).SpecialCells(xlCellTypeVisible).Count
    On Error GoTo 0
    If totalItems = 0 Then totalItems = 1
    processedItems = 0
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Copying Actual-FY..."
    Progress_Bar.percentDone 0
    On Error Resume Next
    For Each cell In wsActual.Range("A12:A" & lastRowActual).SpecialCells(xlCellTypeVisible)
        If Not dictBudgetKeys.Exists(cell.value) Then
            lastRowBudget = wsBudget.Cells(wsBudget.Rows.Count, "B").End(xlUp).row + 1
            wsBudget.Range("B" & lastRowBudget & ":I" & lastRowBudget).value = wsActual.Range("B" & cell.row & ":I" & cell.row).value
            wsBudget.Range("A" & lastRowBudget).value = _
                wsBudget.Range("B" & lastRowBudget).value & _
                wsBudget.Range("C" & lastRowBudget).value & _
                wsBudget.Range("D" & lastRowBudget).value & _
                wsBudget.Range("E" & lastRowBudget).value & _
                wsBudget.Range("F" & lastRowBudget).value & _
                wsBudget.Range("G" & lastRowBudget).value & _
                wsBudget.Range("H" & lastRowBudget).value & _
                wsBudget.Range("I" & lastRowBudget).value
            wsBudget.Range("V" & lastRowBudget).Formula = "=XLOOKUP(A" & lastRowBudget & ",'Actual-FY'!$A:$A,'Actual-FY'!$V:$V,"""")"
            wsBudget.Range("W" & lastRowBudget & ":AJ" & lastRowBudget).FormulaR1C1 = wsBudget.Range("W10:AJ10").FormulaR1C1
            dictBudgetKeys(cell.value) = True
        End If
        processedItems = processedItems + 1
        If processedItems Mod 50 = 0 Or processedItems = totalItems Then
            Progress_Bar.percentDone processedItems / totalItems
            Progress_Bar.setStatusText "Copying Actual-FY: " & processedItems & " of " & totalItems
            DoEvents
        End If
    Next cell
    On Error GoTo 0
    wsActual.AutoFilterMode = False

    ' --- Step 3: Build allocation dictionary from sourceSheet ---
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).row
    Set dictDept = CreateObject("Scripting.Dictionary")
    totalItems = lastRowSource - 9
    processedItems = 0
    Progress_Bar.setStatusText "Building Allocation..."
    Progress_Bar.percentDone 0
    For i = 10 To lastRowSource
        If sourceSheet.Cells(i, "A").value = "Allocated" Then
            department = sourceSheet.Cells(i, 11).value
            For j = 1 To sourceSheet.Columns.Count
                If sourceSheet.Cells(i, j).value = "Column Start>>" Then
                    Exit For
                End If
            Next j
            If j <= sourceSheet.Columns.Count Then
                Dim tempArr(1 To 12) As Double
                For k = 0 To 11
                    tempArr(k + 1) = sourceSheet.Cells(i, j + k + 1).value
                Next k
                dictDept(department) = tempArr
            End If
        End If
        processedItems = processedItems + 1
        If processedItems Mod 100 = 0 Or processedItems = totalItems Then
            Progress_Bar.percentDone processedItems / totalItems
            Progress_Bar.setStatusText "Building Allocation: " & processedItems & " of " & totalItems
            DoEvents
        End If
    Next i

    ' --- Step 4: Update Budget-FY using allocation dictionary, skip excluded rows ---
    Dim lastRowBudgetFY As Long
    lastRowBudgetFY = wsBudget.Cells(wsBudget.Rows.Count, "W").End(xlUp).row
    totalItems = lastRowBudgetFY - 1
    processedItems = 0
    Progress_Bar.setStatusText "Populating Budget-FY..."
    Progress_Bar.percentDone 0
    For i = 2 To lastRowBudgetFY

        ' Gather candidate values from Budget-FY row
        Dim candidate(1 To 9) As Variant
        For k = 1 To 9
            candidate(k) = wsBudget.Cells(i, k).value ' A=1, ..., I=9
        Next k

        ' Check exclusionArr for multi-field AND match
        Dim isExcluded As Boolean
        isExcluded = False
        For x = 1 To exclusionCount
            Dim allMatch As Boolean
            allMatch = True
            For y = 2 To 9 ' Skip column 1 (Type)
                If exclusionArr(x, y) <> "" Then
                    If candidate(y) <> exclusionArr(x, y) Then
                        allMatch = False
                        Exit For
                    End If
                End If
            Next y
            If allMatch Then
                isExcluded = True
                Exit For
            End If
        Next x
        If Not isExcluded Then
            If wsBudget.Cells(i, "W").value = filter5YP Then
                ytdAverage = wsBudget.Cells(i, "V").value
                If IsEmpty(ytdAverage) Or Not IsNumeric(ytdAverage) Then ytdAverage = 0
                department = wsBudget.Cells(i, "Y").value
                If dictDept.Exists(department) Then
                    allocArr = dictDept(department)
                    For j = 1 To 12
                        wsBudget.Cells(i, 9 + j).value = Round((1 + allocArr(j)) * ytdAverage, 6) ' Columns J to U
                    Next j
                End If
            End If
        End If
        processedItems = processedItems + 1
        If processedItems Mod 100 = 0 Or processedItems = totalItems Then
            Progress_Bar.percentDone processedItems / totalItems
            Progress_Bar.setStatusText "Populating Budget-FY: " & processedItems & " of " & totalItems
            DoEvents
        End If
    Next i
    sourceSheet.Calculate
    wsActual.Calculate
    wsBudget.Calculate
    wsExclusion.Calculate

    'Set as values the formulas
    Call SetGlobalFYSheetValues
    elapsedTime = Timer - startTime
    Unload Progress_Bar
    MsgBox "Budget-FY sheet has been populated successfully!" & vbCrLf & _
           "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds", vbInformation

End Sub

Public Sub UpdateBudgetFYWithAllocationsNoDept(filterArr As Variant, sourceSheet As Worksheet, typeTrans As String)

    ' Optimized version using dictionary lookup and reduced worksheet operations for multiple filter categories

    ' For OPEX, PAYROLL and Below Ebitda

    ' Updated Jul 2025
    Call declareGlobal
    Dim wsActual As Worksheet
    Dim wsBudget As Worksheet
    Dim wsExclusion As Worksheet
    Dim lastRowActual As Long, lastRowBudget As Long, lastRowSource As Long, lastRowExclusion As Long
    Dim dictBudgetKeys As Object, dictAlloc As Object
    Dim i As Long, j As Long, k As Long
    Dim ytdAverage As Variant, department As String
    Dim cell As Range
    Dim processedItems As Long, totalItems As Long
    Dim startTime As Double, elapsedTime As Double
    Dim arrBudgetKeys As Variant
    Dim rowIndex As Long
    Dim colStart As Long
    startTime = Timer

    ' Set worksheet references for performance
    Set wsActual = ThisWorkbook.Sheets("Actual-FY")
    Set wsBudget = ThisWorkbook.Sheets("Budget-FY")
    Set wsExclusion = ThisWorkbook.Sheets("Exclusions")

    ' Calculate all sheets to ensure formulas are up to date
    sourceSheet.Calculate
    wsActual.Calculate
    wsBudget.Calculate
    wsExclusion.Calculate

    ' --- Load exclusionArr from UpdateExclusion sheet for matching type ---
    lastRowExclusion = wsExclusion.Cells(wsExclusion.Rows.Count, "A").End(xlUp).row
    Dim tempArr() As Variant
    Dim exclusionArr() As Variant
    Dim exclusionCount As Long
    exclusionCount = 0

    ' Collect matching rows into tempArr as array of arrays
    For i = 2 To lastRowExclusion
        If wsExclusion.Cells(i, 1).value = typeTrans Then
            exclusionCount = exclusionCount + 1
            ReDim Preserve tempArr(1 To exclusionCount)
            Dim rowArr(1 To 9) As Variant
            For j = 1 To 9
                If IsEmpty(wsExclusion.Cells(i, j).value) Or wsExclusion.Cells(i, j).value = "" Then
                    rowArr(j) = ""
                Else
                    rowArr(j) = wsExclusion.Cells(i, j).value
                End If
            Next j
            tempArr(exclusionCount) = rowArr
        End If
    Next i

    ' Copy tempArr to a 2D array exclusionArr
    If exclusionCount > 0 Then
        ReDim exclusionArr(1 To exclusionCount, 1 To 9)
        For i = 1 To exclusionCount
            For j = 1 To 9
                exclusionArr(i, j) = tempArr(i)(j)
            Next j
        Next i
    Else
        ReDim exclusionArr(1 To 1, 1 To 9)
        For j = 1 To 9
            exclusionArr(1, j) = ""
        Next j
    End If

    ' --- Build exclusionLookupArr for filterArr removal ---
    Dim exclusionLookupArr() As String
    If exclusionCount > 0 Then
        ReDim exclusionLookupArr(1 To exclusionCount)
        For i = 1 To exclusionCount
            Dim lookupStr As String
            lookupStr = ""
            For j = 2 To 9 ' Skip Type
                If exclusionArr(i, j) <> "" Then
                    lookupStr = lookupStr & "|" & exclusionArr(i, j)
                End If
            Next j
            exclusionLookupArr(i) = lookupStr
        Next i
    Else
        ReDim exclusionLookupArr(1 To 1)
        exclusionLookupArr(1) = ""
    End If

    ' --- Remove matching values from filterArr ---
    Dim newFilterArr() As Variant
    Dim newCount As Long
    newCount = 0
    For i = LBound(filterArr) To UBound(filterArr)
        Dim isExcluded As Boolean
        isExcluded = False
        For j = 1 To exclusionCount
            If filterArr(i) = exclusionLookupArr(j) Then
                isExcluded = True
                Exit For
            End If
        Next j
        If Not isExcluded Then
            newCount = newCount + 1
            ReDim Preserve newFilterArr(1 To newCount)
            newFilterArr(newCount) = filterArr(i)
        End If
    Next i
    filterArr = newFilterArr

    ' Debug.Print first 5 rows of exclusionArr
    For i = 1 To Application.Min(5, exclusionCount)
        Dim debugStr As String
        debugStr = ""
        For j = 1 To 9
            debugStr = debugStr & exclusionArr(i, j) & ";"
        Next j
        Debug.Print "ExclusionArr Row " & i & ": " & debugStr
    Next i

    ' Initialize progress tracking for 3 main steps
    totalItems = 3
    processedItems = 0
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Step 1: Copying Actual-FY..."
    Progress_Bar.percentDone processedItems / totalItems

    ' --- Step 1: Build dictionary of existing keys in Budget-FY for fast lookup ---
    lastRowBudget = wsBudget.Cells(wsBudget.Rows.Count, "A").End(xlUp).row
    Set dictBudgetKeys = CreateObject("Scripting.Dictionary")
    If lastRowBudget >= 2 Then
        arrBudgetKeys = wsBudget.Range("A2:A" & lastRowBudget).value
        For i = 1 To UBound(arrBudgetKeys, 1)
            If Not IsEmpty(arrBudgetKeys(i, 1)) Then
                dictBudgetKeys(arrBudgetKeys(i, 1)) = True
            End If
        Next i
    End If

    ' --- Step 1: Copy new rows from Actual-FY to Budget-FY if not already present ---
    lastRowActual = wsActual.Cells(wsActual.Rows.Count, "A").End(xlUp).row
    For Each filterCategory In filterArr
        wsActual.Range("A1:Y" & lastRowActual).AutoFilter Field:=23, Criteria1:=filterCategory
        On Error Resume Next
        For Each cell In wsActual.Range("A12:A" & lastRowActual).SpecialCells(xlCellTypeVisible)

            ' Gather candidate values from Actual-FY row
            Dim candidate(1 To 9) As Variant
            For k = 1 To 9
                candidate(k) = wsActual.Cells(cell.row, k).value ' A=1, ..., I=9
            Next k

            ' Check exclusionArr for multi-field AND match
            isExcluded = False
            For i = 1 To exclusionCount
                Dim allMatch As Boolean
                allMatch = True
                For j = 2 To 9 ' Skip column 1 (Type)
                    If exclusionArr(i, j) <> "" Then
                        If candidate(j) <> exclusionArr(i, j) Then
                            allMatch = False
                            Exit For
                        End If
                    End If
                Next j
                If allMatch Then
                    isExcluded = True
                    Exit For
                End If
            Next i
            If Not isExcluded Then
                If Not dictBudgetKeys.Exists(cell.value) Then
                    lastRowBudget = wsBudget.Cells(wsBudget.Rows.Count, "B").End(xlUp).row + 1
                    wsBudget.Range("B" & lastRowBudget & ":I" & lastRowBudget).value = wsActual.Range("B" & cell.row & ":I" & cell.row).value
                    wsBudget.Range("A" & lastRowBudget).value = _
                        wsBudget.Range("B" & lastRowBudget).value & _
                        wsBudget.Range("C" & lastRowBudget).value & _
                        wsBudget.Range("D" & lastRowBudget).value & _
                        wsBudget.Range("E" & lastRowBudget).value & _
                        wsBudget.Range("F" & lastRowBudget).value & _
                        wsBudget.Range("G" & lastRowBudget).value & _
                        wsBudget.Range("H" & lastRowBudget).value & _
                        wsBudget.Range("I" & lastRowBudget).value
                    wsBudget.Range("V" & lastRowBudget).Formula = "=XLOOKUP(A" & lastRowBudget & ",'Actual-FY'!$A:$A,'Actual-FY'!$V:$V,"""")"
                    wsBudget.Range("W" & lastRowBudget & ":AJ" & lastRowBudget).FormulaR1C1 = wsBudget.Range("W10:AJ10").FormulaR1C1
                    dictBudgetKeys(cell.value) = True
                End If
            End If
        Next cell
        On Error GoTo 0
    Next filterCategory
    wsActual.AutoFilterMode = False
    processedItems = processedItems + 1
    Progress_Bar.percentDone processedItems / totalItems
    Progress_Bar.setStatusText "Step 2: Building Allocation..."

    ' --- Step 2: Build allocation dictionary from sourceSheet ---
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).row
    Set dictAlloc = CreateObject("Scripting.Dictionary")
    For i = 10 To lastRowSource
        If sourceSheet.Cells(i, "A").value = "Allocated" Then
            department = sourceSheet.Cells(i, 11).value
            For j = 1 To sourceSheet.Columns.Count
                If sourceSheet.Cells(i, j).value = "Column Start>>" Then
                    colStart = j + 1
                    Exit For
                End If
            Next j
            If colStart > 0 Then
                Dim tempArrAlloc(1 To 12) As Double
                For k = 0 To 11
                    tempArrAlloc(k + 1) = sourceSheet.Cells(i, colStart + k).value
                Next k
                dictAlloc(department) = tempArrAlloc
            End If
        End If
        If i Mod 100 = 0 Then DoEvents
    Next i
    processedItems = processedItems + 1
    Progress_Bar.percentDone processedItems / totalItems
    Progress_Bar.setStatusText "Step 3: Populating Budget-FY..."

    ' --- Step 3: Update Budget-FY using allocation dictionary, skip excluded rows ---
    Dim lastRowBudgetFY As Long
    Dim allocArr As Variant
    lastRowBudgetFY = wsBudget.Cells(wsBudget.Rows.Count, "W").End(xlUp).row
    For i = 2 To lastRowBudgetFY

        ' Gather candidate values from Budget-FY row
        For k = 1 To 9
            candidate(k) = wsBudget.Cells(i, k).value ' A=1, ..., I=9
        Next k

        ' Check exclusionArr for multi-field AND match
        isExcluded = False
        For x = 1 To exclusionCount
            allMatch = True
            For y = 2 To 9 ' Skip column 1 (Type)
                If exclusionArr(x, y) <> "" Then
                    If candidate(y) <> exclusionArr(x, y) Then
                        allMatch = False
                        Exit For
                    End If
                End If
            Next y
            If allMatch Then
                isExcluded = True
                Exit For
            End If
        Next x
        If Not isExcluded Then
            For j = LBound(filterArr) To UBound(filterArr)
                If wsBudget.Cells(i, "W").value = filterArr(j) Then
                    ytdAverage = wsBudget.Cells(i, "V").value
                    If IsEmpty(ytdAverage) Or Not IsNumeric(ytdAverage) Then ytdAverage = 0
                    department = filterArr(j)
                    If dictAlloc.Exists(department) Then
                        allocArr = dictAlloc(department)
                        For colStart = 10 To 21
                            wsBudget.Cells(i, colStart).value = Round((1 + allocArr(colStart - 9)) * ytdAverage, 6)
                        Next colStart
                    End If
                    Exit For
                End If
            Next j
        End If
        If i Mod 100 = 0 Then DoEvents
    Next i
    processedItems = processedItems + 1
    Progress_Bar.percentDone processedItems / totalItems
    Progress_Bar.setStatusText "Done!"
    sourceSheet.Calculate
    wsActual.Calculate
    wsBudget.Calculate
    wsExclusion.Calculate

    'Set as values the formulas
    Call SetGlobalFYSheetValues
    elapsedTime = Timer - startTime
    Unload Progress_Bar
    MsgBox "Budget-FY sheet has been populated successfully!" & vbCrLf & _
           "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds", vbInformation

End Sub



Public Sub newRowsActual(globalSheet As Worksheet)

    ' Updated Aug 2025 - Optimized version using bulk operations and in-memory processing for Actual data
    Call declareGlobal
    globalSheet.Calculate
    Dim startTime As Double
    startTime = Timer
    Dim totalItems As Long, processedItems As Long
    Dim i As Long, j As Long, k As Long
    Dim lastRowActual As Long, maxRowActual As Long
    Dim arrActualData As Variant, arrGlobalData As Variant

    ' Unfilter Actual-FY sheet if filtered
    If globalActualFYSheet.AutoFilterMode Then globalActualFYSheet.AutoFilterMode = False

    ' Initialize progress
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Step 1: Loading existing actual data..."
    Progress_Bar.percentDone 0

    ' --- Step 1: Load existing Actual-FY data into two separate dictionaries ---

    ' Detect actual vs forecast columns from globalSheet row 7
    Dim actualForecastData As Variant
    actualForecastData = globalSheet.Range("A7:ZZ7").value

    ' Determine actual and forecast column ranges dynamically
    Dim actualColStart As Long, actualColEnd As Long
    Dim forecastColStart As Long, forecastColEnd As Long

    ' Initialize to detect ranges
    forecastColStart = 0  ' Will find minimum
    forecastColEnd = 0    ' Will find maximum

    ' Debug: Show what's in row 7 of Global sheet
    Debug.Print "=== GLOBAL SHEET ROW 7 ANALYSIS ==="
    For i = 1 To UBound(actualForecastData, 2)
        If actualForecastData(1, i) <> "" Then
            Debug.Print "Global Row7[" & i & "] = '" & actualForecastData(1, i) & "'"
        End If
    Next i

    ' Detect forecast columns in Global sheet
    For i = 1 To UBound(actualForecastData, 2)
        If actualForecastData(1, i) = "Forecast" Then
            If forecastColStart = 0 Then forecastColStart = i
            forecastColEnd = i
            Debug.Print "Found Forecast at Global column " & i
        End If
    Next i

    ' Error if no forecast columns detected
    If forecastColStart = 0 Then
        Err.Raise vbObjectError + 1001, "newRowsActual", "No 'Forecast' columns found in Global sheet row 7. Cannot proceed with actual updates."
    End If

    ' Map Global sheet forecast columns to Actual-FY sheet

    ' Global columns AJ-AU (36-47) map to Actual-FY columns J-U (10-21)

    ' Subtract 26 to map Global column 36->10 (J), 44->18 (R), etc.
    Dim mappingOffset As Long
    mappingOffset = 26  ' Global col 36 (AJ) -> Actual-FY col 10 (J)
    forecastColStart = forecastColStart - mappingOffset
    forecastColEnd = forecastColEnd - mappingOffset

    ' Validate mapping results
    If forecastColStart < 10 Or forecastColEnd > 21 Then
        Err.Raise vbObjectError + 1002, "newRowsActual", "Forecast column mapping out of valid range. Global forecast columns " & (forecastColStart + mappingOffset) & "-" & (forecastColEnd + mappingOffset) & " map to Actual-FY columns " & forecastColStart & "-" & forecastColEnd & " which is outside valid range J-U (10-21)."
    End If

    ' Actual columns end where forecast begins
    actualColStart = 10      ' J (January) - always start with January
    actualColEnd = forecastColStart - 1  ' End the month before forecast starts
    Debug.Print "FINAL DETECTED RANGES:"
    Debug.Print "  Global forecast columns: " & (forecastColStart + mappingOffset) & "-" & (forecastColEnd + mappingOffset)
    Debug.Print "  Actual-FY actual columns: " & Chr(64 + actualColStart) & " to " & Chr(64 + actualColEnd) & " (" & actualColStart & "-" & actualColEnd & ")"
    Debug.Print "  Actual-FY forecast columns: " & Chr(64 + forecastColStart) & " to " & Chr(64 + forecastColEnd) & " (" & forecastColStart & "-" & forecastColEnd & ")"

    ' Create separate dictionaries for actual and forecast data
    Dim globalActualDictActualCols As Object
    Dim globalActualDictForecastCols As Object
    Dim forecastUpdatedCols As Object  ' Track which columns were actually updated
    Set globalActualDictActualCols = CreateObject("Scripting.Dictionary")
    Set globalActualDictForecastCols = CreateObject("Scripting.Dictionary")
    Set forecastUpdatedCols = CreateObject("Scripting.Dictionary")
    lastRowActual = globalActualFYSheet.Cells(globalActualFYSheet.Rows.Count, "F").End(xlUp).row
    maxRowActual = lastRowActual
    If lastRowActual >= 12 Then

        ' Read all existing data in one bulk operation (A to U columns)
        arrActualData = globalActualFYSheet.Range("A12:U" & lastRowActual).value

        ' Build separate dictionaries from existing data
        For i = 1 To UBound(arrActualData, 1)
            If arrActualData(i, 1) <> "" Then ' If lookup key exists

                ' Create actual columns dictionary entry (A-I + actual month columns)
                Dim actualRowActuals(1 To 23) As Variant ' Row + 22 columns max (A-V)
                actualRowActuals(1) = i + 11 ' Actual row number in sheet

                ' Copy A to I (metadata)
                For j = 1 To 9
                    actualRowActuals(j + 1) = arrActualData(i, j)
                Next j

                ' Copy actual month columns (J to Q typically) - fix array indexing
                For j = actualColStart To actualColEnd
                    actualRowActuals(j + 1) = arrActualData(i, j)
                Next j

                ' Initialize forecast columns as empty in actual dictionary
                For j = forecastColStart To forecastColEnd
                    actualRowActuals(j + 1) = ""
                Next j
                globalActualDictActualCols(arrActualData(i, 1)) = actualRowActuals

                ' Create forecast columns dictionary entry (A-I + forecast month columns only)
                Dim actualRowForecasts(1 To 23) As Variant ' Row + 22 columns max (A-V)
                actualRowForecasts(1) = i + 11 ' Actual row number in sheet

                ' Copy A to I (metadata)
                For j = 1 To 9
                    actualRowForecasts(j + 1) = arrActualData(i, j)
                Next j

                ' Initialize actual columns as empty in forecast dictionary
                For j = actualColStart To actualColEnd
                    actualRowForecasts(j + 1) = ""
                Next j

                ' Copy forecast month columns (P to U typically) - fix array indexing
                For j = forecastColStart To forecastColEnd
                    actualRowForecasts(j + 1) = arrActualData(i, j)
                Next j
                globalActualDictForecastCols(arrActualData(i, 1)) = actualRowForecasts
            End If
        Next i
        maxRowActual = 11 ' Start from row 12 when adding new data
    End If

    ' Debug print first 10 entries from globalActualDictActualCols
    Debug.Print "=== First 10 entries in globalActualDictActualCols (DETAILED) ==="
    Dim entryCount As Long, key As Variant
    entryCount = 0
    For Each key In globalActualDictActualCols.keys
        entryCount = entryCount + 1
        If entryCount <= 10 Then
            Dim rowData As Variant
            rowData = globalActualDictActualCols(key)
            Debug.Print "Entry " & entryCount & ": Key=" & Left(key, 40) & "..."
            Debug.Print "  Row=" & rowData(1) & " | A=" & Left(rowData(2), 20) & "... | B=" & rowData(3) & " | C=" & rowData(4) & " | D=" & Left(rowData(5), 15) & "..."
            Debug.Print "  J=" & rowData(11) & " | K=" & rowData(12) & " | L=" & rowData(13) & " | M=" & rowData(14) & " | N=" & rowData(15) & " | O=" & rowData(16) & " | P=" & rowData(17) & " | Q=" & rowData(18)
            Debug.Print "  R=" & rowData(19) & " | S=" & rowData(20) & " | T=" & rowData(21) & " | U=" & rowData(22)
        Else
            Exit For
        End If
    Next key

    ' Debug print first 10 entries from globalActualDictForecastCols
    Debug.Print "=== First 10 entries in globalActualDictForecastCols (DETAILED) ==="
    entryCount = 0
    For Each key In globalActualDictForecastCols.keys
        entryCount = entryCount + 1
        If entryCount <= 10 Then
            rowData = globalActualDictForecastCols(key)
            Debug.Print "Entry " & entryCount & ": Key=" & Left(key, 40) & "..."
            Debug.Print "  Row=" & rowData(1) & " | A=" & Left(rowData(2), 20) & "... | B=" & rowData(3) & " | C=" & rowData(4) & " | D=" & Left(rowData(5), 15) & "..."
            Debug.Print "  J=" & rowData(11) & " | K=" & rowData(12) & " | L=" & rowData(13) & " | M=" & rowData(14) & " | N=" & rowData(15) & " | O=" & rowData(16) & " | P=" & rowData(17) & " | Q=" & rowData(18)
            Debug.Print "  R=" & rowData(19) & " | S=" & rowData(20) & " | T=" & rowData(21) & " | U=" & rowData(22)
        Else
            Exit For
        End If
    Next key
    Progress_Bar.percentDone 0.2
    Progress_Bar.setStatusText "Step 2: Processing global sheet data for actuals..."

    ' --- Step 2: Process globalSheet data and build updates for forecast dictionary only ---
    Dim globalSheetDict As Object
    Set globalSheetDict = CreateObject("Scripting.Dictionary")
    Dim cellNewRow As Range, cell As Range, new_RowRange As Range
    Dim colStart As Long, colEnd As Long

    ' Read globalSheet data into array for faster processing
    Dim lastRowGlobal As Long
    lastRowGlobal = globalSheet.Cells(globalSheet.Rows.Count, "A").End(xlUp).row
    If lastRowGlobal < 5000 Then lastRowGlobal = 5000
    arrGlobalData = globalSheet.Range("A10:ZZ" & lastRowGlobal).value

    ' Detect forecast columns from globalSheet row 7 and create dynamic mapping
    Dim forecastCols As Object
    Set forecastCols = CreateObject("Scripting.Dictionary")
    Dim forecastColMapping As Object
    Set forecastColMapping = CreateObject("Scripting.Dictionary")
    Dim forecastRowData As Variant
    forecastRowData = globalSheet.Range("A7:ZZ7").value

    ' Map forecast columns from globalSheet to target forecast columns in Actual-FY sheet
    Dim forecastIndex As Long
    forecastIndex = 0
    For i = 1 To UBound(forecastRowData, 2)
        If forecastRowData(1, i) = "Forecast" Then

            ' Check if we would exceed U column (21) OR the detected forecast range BEFORE creating the mapping
            If (forecastColStart + forecastIndex) > 21 Or (forecastColStart + forecastIndex) > forecastColEnd Then
                Exit For
            End If
            forecastCols(i) = True

            ' Map to forecast columns in Actual-FY sheet dynamically based on detected range
            forecastColMapping(i) = forecastColStart + forecastIndex
            forecastIndex = forecastIndex + 1
        End If
    Next i

    ' Process each row type in memory
    For i = 1 To UBound(arrGlobalData, 1)
        Dim rowType As String
        rowType = arrGlobalData(i, 1) ' Column A

        ' Check if row should be processed (only REFORECAST in column AI)
        Dim exclusionFlag As String
        If UBound(arrGlobalData, 2) >= 35 Then ' Ensure column AI (35) exists in array
            exclusionFlag = Trim(CStr(arrGlobalData(i, 35))) ' Column AI = 35
            If UCase(exclusionFlag) <> "REFORECAST" Then
                Debug.Print "Skipping row " & (i + 9) & " - only processing rows with REFORECAST flag"
                GoTo NextRow ' Skip this row entirely
            End If
        End If
        If rowType = "KPI" Or rowType = "Amount" Or rowType = "Percentage" Or rowType = "Amount with Allocation" Or rowType = "Adjustment" Then
            Dim actualRow1 As Long
            actualRow1 = i + 9 ' Convert back to actual sheet row

            ' Handle Adjustment transactions differently
            If rowType = "Adjustment" Then

                ' Process each forecast column for Actual Adjustment
                For Each forecastCol In forecastCols.keys
                    If forecastCol <= UBound(arrGlobalData, 2) And arrGlobalData(i, forecastCol) <> "" Then
                        Call ProcessActualAdjustmentTransaction(arrGlobalData, globalActualDictForecastCols, i, CLng(forecastCol), forecastColMapping, forecastUpdatedCols)
                        totalItems = totalItems + 1
                    End If
                Next forecastCol
            Else

                ' Handle regular transaction types (KPI, Amount, Percentage, Amount with Allocation)

                ' Determine column range based on row type
                If rowType = "Percentage" Or rowType = "Amount with Allocation" Then

                    ' Find column markers in the array
                    colStart = 0
                    colEnd = 0
                    For j = 1 To UBound(arrGlobalData, 2)
                        If arrGlobalData(i, j) = "Column Start>>" Or arrGlobalData(i, j) = "Allocation Start>>" Then
                            colStart = j + 1
                        ElseIf arrGlobalData(i, j) = "<<Column End" Or arrGlobalData(i, j) = "<<Allocation End" Then
                            colEnd = j - 1
                            Exit For
                        End If
                    Next j
                    If colStart = 0 Then colStart = 12
                    If colEnd = 0 Then colEnd = 23
                Else
                    colStart = 12 ' Column L
                    colEnd = 23   ' Column W
                End If

                ' Process each forecast column for Actual (only forecast columns)
                For Each forecastCol In forecastCols.keys
                    If forecastCol <= UBound(arrGlobalData, 2) And arrGlobalData(i, forecastCol) <> "" Then
                        Call ProcessActualSheetCell(globalSheetDict, arrGlobalData, globalActualDictForecastCols, i, CLng(forecastCol), rowType, colStart, actualRow1, maxRowActual, forecastColMapping, forecastUpdatedCols, forecastCols)
                        totalItems = totalItems + 1
                    End If
                Next forecastCol
            End If
        End If
NextRow:
    Next i
    Progress_Bar.percentDone 0.6
    Progress_Bar.setStatusText "Step 3: Combining dictionaries..."

    ' --- Step 3: Combine actual and forecast dictionaries into globalActualDict ---
    Dim globalActualDict As Object
    Set globalActualDict = CreateObject("Scripting.Dictionary")
    For Each key In globalActualDictActualCols.keys
        Dim actualRowData As Variant, forecastRowData1 As Variant, mergedRow(1 To 23) As Variant
        actualRowData = globalActualDictActualCols(key)

        ' Start with actual data
        For j = 1 To 23
            mergedRow(j) = actualRowData(j)
        Next j

        ' Overlay forecast data if exists
        If globalActualDictForecastCols.Exists(key) Then
            forecastRowData1 = globalActualDictForecastCols(key)

            ' Copy forecast columns only
            For j = forecastColStart To forecastColEnd
                If j + 1 <= 23 Then
                    mergedRow(j + 1) = forecastRowData1(j + 1)
                End If
            Next j
        End If
        globalActualDict(key) = mergedRow
    Next key

    ' Add any new entries that exist only in forecast dictionary
    For Each key In globalActualDictForecastCols.keys
        If Not globalActualDictActualCols.Exists(key) Then
            globalActualDict(key) = globalActualDictForecastCols(key)
        End If
    Next key

    ' Debug print first 10 entries from combined globalActualDict
    Debug.Print "=== First 10 entries in combined globalActualDict (DETAILED) ==="
    entryCount = 0
    For Each key In globalActualDict.keys
        entryCount = entryCount + 1
        If entryCount <= 10 Then
            rowData = globalActualDict(key)
            Debug.Print "Entry " & entryCount & ": Key=" & Left(key, 40) & "..."
            Debug.Print "  Row=" & rowData(1) & " | A=" & Left(rowData(2), 20) & "... | B=" & rowData(3) & " | C=" & rowData(4) & " | D=" & Left(rowData(5), 15) & "..."
            Debug.Print "  J=" & rowData(11) & " | K=" & rowData(12) & " | L=" & rowData(13) & " | M=" & rowData(14) & " | N=" & rowData(15) & " | O=" & rowData(16) & " | P=" & rowData(17) & " | Q=" & rowData(18)
            Debug.Print "  R=" & rowData(19) & " | S=" & rowData(20) & " | T=" & rowData(21) & " | U=" & rowData(22)
            Debug.Print "  Forecast Range: " & Chr(64 + forecastColStart) & "-" & Chr(64 + forecastColEnd) & " (" & forecastColStart & "-" & forecastColEnd & ")"
        Else
            Exit For
        End If
    Next key
    Progress_Bar.setStatusText "Step 4: Writing updated data back to Actual-FY..."

    ' --- Step 4: Write all data back to Actual-FY in bulk ---
    Call WriteBulkActualData(globalActualDict, maxRowActual, forecastUpdatedCols, actualColStart, actualColEnd, forecastColStart, forecastColEnd)
    Progress_Bar.percentDone 0.8
    Progress_Bar.setStatusText "Step 5: Applying formulas..."

    ' --- Step 5: Apply formulas in bulk ---

    ' Recalculate maxRowActual before applying formulas
    maxRowActual = globalActualFYSheet.Cells(globalActualFYSheet.Rows.Count, "F").End(xlUp).row
    If maxRowActual >= 12 Then

        ' Copy formulas from V10:AJ10 to V12:AJ[maxRowActual] using copy-paste
        globalActualFYSheet.Range("V10:AJ10").Copy
        globalActualFYSheet.Range("V12:AJ" & maxRowActual).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If
    Progress_Bar.percentDone 0.9
    Progress_Bar.setStatusText "Step 5: Finalizing..."

    ' Calculate and finalize
    globalActualFYSheet.Calculate
    globalSheet.Calculate

    'Set as values the formulas
    Call SetGlobalFYSheetValues
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    Unload Progress_Bar
    MsgBox "Actual update finished!" & vbCrLf & "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf & "Total items processed: " & totalItems, vbInformation

End Sub

Private Sub ProcessActualSheetCell(globalSheetDict As Object, arrGlobalData As Variant, globalActualDictForecastCols As Object, rowIndex As Long, colIndex As Long, mode As String, colStart As Long, actualRow As Long, ByRef maxRowActual As Long, forecastColMapping As Object, forecastUpdatedCols As Object, forecastCols As Object)

    ' Process individual cell from globalSheet and update actual dictionaries
    Dim company As String, CostCenter As String, customerSegment As String
    Dim venue As String, ledgerAccount As String, revCat As String
    Dim spendCat As String, usGaap As String, lookupValue As String, updateValue As Variant
    Dim targetCol As Long

    ' Extract metadata from the row
    company = arrGlobalData(rowIndex, 2)     ' Column B
    CostCenter = arrGlobalData(rowIndex, 3)  ' Column C
    customerSegment = arrGlobalData(rowIndex, 4) ' Column D
    venue = arrGlobalData(rowIndex, 5)       ' Column E
    usGaap = arrGlobalData(rowIndex, 9)      ' Column I

    ' Process based on mode
    If mode = "Amount" Or mode = "KPI" Then
        ledgerAccount = Left(arrGlobalData(rowIndex, 6), 6) ' Column F
        revCat = arrGlobalData(rowIndex, 7)  ' Column G
        spendCat = arrGlobalData(rowIndex, 8) ' Column H
        lookupValue = company & CostCenter & customerSegment & venue & ledgerAccount & revCat & spendCat & usGaap
        If mode = "Amount" Then
            updateValue = -arrGlobalData(rowIndex, colIndex)
        ElseIf mode = "KPI" Then
            updateValue = arrGlobalData(rowIndex, colIndex)
        End If

        ' Map forecast column to target column dynamically based on forecast column mapping
        If forecastColMapping.Exists(colIndex) Then
            targetCol = forecastColMapping(colIndex)
        Else

            ' Fallback if mapping not found (shouldn't happen)
            targetCol = 10 ' Default to J
        End If
        Call UpdateActualDictionaryValue(globalActualDictForecastCols, lookupValue, updateValue, targetCol, company, CostCenter, customerSegment, venue, ledgerAccount, revCat, spendCat, usGaap, maxRowActual, forecastUpdatedCols)
    ElseIf mode = "Percentage" Then
        ledgerAccount = Left(arrGlobalData(rowIndex, 6), 6)
        revCat = arrGlobalData(rowIndex, 7)
        spendCat = arrGlobalData(rowIndex, 8)
        lookupValue = company & CostCenter & customerSegment & venue & ledgerAccount & revCat & spendCat & usGaap
        updateValue = -arrGlobalData(rowIndex, colIndex)

        ' Map forecast column to target column - same logic as Amount/KPI
        If forecastColMapping.Exists(colIndex) Then
            targetCol = forecastColMapping(colIndex)
        Else

            ' Fallback if mapping not found (shouldn't happen)
            targetCol = 10 ' Default to J
        End If
        Call UpdateActualDictionaryValue(globalActualDictForecastCols, lookupValue, updateValue, targetCol, company, CostCenter, customerSegment, venue, ledgerAccount, revCat, spendCat, usGaap, maxRowActual, forecastUpdatedCols)
    ElseIf mode = "Amount with Allocation" Then

        ' Find allocation markers and related rows using corrected logic
        Dim allocationStartCol As Long
        allocationStartCol = 0

        ' Find "Allocation Start>>" in the current row
        For j = 1 To UBound(arrGlobalData, 2)
            If arrGlobalData(rowIndex, j) = "Allocation Start>>" Then
                allocationStartCol = j
                Exit For
            End If
        Next j

        ' Early exit if no allocation marker found
        If allocationStartCol = 0 Then
            Debug.Print "Warning: 'Amount with Allocation' row at rowIndex " & rowIndex & " has no 'Allocation Start>>' marker - skipping allocation processing"
            Exit Sub
        End If

        ' Find ledger, spend, and revenue category rows with separate searches
        Dim ledgerRow As Long, spendRow As Long, revRow As Long
        ledgerRow = 0: spendRow = 0: revRow = 0

        ' Search for LedgerAcct>> marker
        For searchRow = rowIndex - 1 To 1 Step -1
            If searchRow >= 1 And searchRow <= UBound(arrGlobalData, 1) And allocationStartCol >= 1 And allocationStartCol <= UBound(arrGlobalData, 2) Then
                If arrGlobalData(searchRow, allocationStartCol) = "LedgerAcct>>" Then
                    ledgerRow = searchRow
                    Exit For
                End If
            End If
        Next searchRow

        ' Search for SpendCat>> marker
        For searchRow = rowIndex - 1 To 1 Step -1
            If searchRow >= 1 And searchRow <= UBound(arrGlobalData, 1) And allocationStartCol >= 1 And allocationStartCol <= UBound(arrGlobalData, 2) Then
                If arrGlobalData(searchRow, allocationStartCol) = "SpendCat>>" Then
                    spendRow = searchRow
                    Exit For
                End If
            End If
        Next searchRow

        ' Search for RevCat>> marker
        For searchRow = rowIndex - 1 To 1 Step -1
            If searchRow >= 1 And searchRow <= UBound(arrGlobalData, 1) And allocationStartCol >= 1 And allocationStartCol <= UBound(arrGlobalData, 2) Then
                If arrGlobalData(searchRow, allocationStartCol) = "RevCat>>" Then
                    revRow = searchRow
                    Exit For
                End If
            End If
        Next searchRow

        ' Only process if all required rows were found
        If ledgerRow > 0 And spendRow > 0 And revRow > 0 And allocationStartCol > 0 Then

            ' Process allocation: multiply forecast amount by each allocation percentage

            ' First, find the actual allocation range dynamically from the markers
            If forecastColMapping.Exists(colIndex) Then
                Dim targetForecastCol As Long
                targetForecastCol = forecastColMapping(colIndex)

                ' Get the forecast amount for this column
                Dim forecastAmount As Double
                forecastAmount = arrGlobalData(rowIndex, colIndex)

                ' Find the allocation end marker to determine the actual range
                Dim allocationEndCol As Long
                allocationEndCol = 0
                For j = allocationStartCol + 1 To UBound(arrGlobalData, 2)
                    If arrGlobalData(rowIndex, j) = "<<Allocation End" Then
                        allocationEndCol = j - 1 ' End column is one before the marker
                        Exit For
                    End If
                Next j

                ' If no end marker found, use a reasonable default
                If allocationEndCol = 0 Then allocationEndCol = allocationStartCol + 10

                ' Loop through allocation percentage columns dynamically
                For allocationCol = allocationStartCol + 1 To allocationEndCol
                    If allocationCol <= UBound(arrGlobalData, 2) And ledgerRow <= UBound(arrGlobalData, 1) And spendRow <= UBound(arrGlobalData, 1) And revRow <= UBound(arrGlobalData, 1) Then

                        ' Read categories from the marker rows, but from the allocation column (not marker column)
                        ledgerAccount = Left(arrGlobalData(ledgerRow, allocationCol), 6)
                        spendCat = arrGlobalData(spendRow, allocationCol)
                        revCat = arrGlobalData(revRow, allocationCol)

                        ' Read allocation percentage from current allocation column with error handling
                        Dim allocationPercentage As Double
                        allocationPercentage = 0 ' Default to 0

                        ' Handle potential type mismatch by checking if the value is numeric
                        If IsNumeric(arrGlobalData(rowIndex, allocationCol)) And arrGlobalData(rowIndex, allocationCol) <> "" Then
                            allocationPercentage = CDbl(arrGlobalData(rowIndex, allocationCol))
                        End If

                        ' Only process if we have valid categories and non-zero allocation percentage

                        ' Allow processing even if spendCat is empty, as long as we have ledger and rev categories
                        If ledgerAccount <> "" And revCat <> "" And allocationPercentage <> 0 Then
                            lookupValue = company & CostCenter & customerSegment & venue & ledgerAccount & revCat & spendCat & usGaap
                            updateValue = -(forecastAmount * allocationPercentage)
                            Call UpdateActualDictionaryValue(globalActualDictForecastCols, lookupValue, updateValue, targetForecastCol, company, CostCenter, customerSegment, venue, ledgerAccount, revCat, spendCat, usGaap, maxRowActual, forecastUpdatedCols)
                        Else
                        End If
                    Else
                    End If
                Next allocationCol
            End If
        Else
            Debug.Print "Warning: Could not find all required allocation rows for rowIndex " & rowIndex & " (allocationStartCol=" & allocationStartCol & ", ledgerRow=" & ledgerRow & ", spendRow=" & spendRow & ", revRow=" & revRow & ")"
        End If
    End If

End Sub

Private Sub ProcessActualAdjustmentTransaction(arrGlobalData As Variant, globalActualDictForecastCols As Object, rowIndex As Long, colIndex As Long, forecastColMapping As Object, forecastUpdatedCols As Object)

    ' Process Adjustment transaction type for actuals - updates existing rows only, never creates new rows
    Dim company As String, CostCenter As String, customerSegment As String
    Dim venue As String, ledgerAccount As String, revCat As String
    Dim spendCat As String, usGaap As String, updateValue As Variant
    Dim targetCol As Long

    ' Extract metadata from the row - some may be empty for partial matching
    company = arrGlobalData(rowIndex, 2)     ' Column B
    CostCenter = arrGlobalData(rowIndex, 3)  ' Column C
    customerSegment = arrGlobalData(rowIndex, 4) ' Column D
    venue = arrGlobalData(rowIndex, 5)       ' Column E
    ledgerAccount = Left(arrGlobalData(rowIndex, 6), 6) ' Column F
    revCat = arrGlobalData(rowIndex, 7)      ' Column G
    spendCat = arrGlobalData(rowIndex, 8)    ' Column H
    usGaap = arrGlobalData(rowIndex, 9)      ' Column I
    updateValue = arrGlobalData(rowIndex, colIndex)

    ' Map forecast column to target column dynamically
    If forecastColMapping.Exists(colIndex) Then
        targetCol = forecastColMapping(colIndex)
    Else

        ' Fallback if mapping not found (shouldn't happen)
        targetCol = 10 ' Default to J
    End If

    ' Find all matching existing entries in globalActualDictForecastCols and update them
    Dim key As Variant, rowData As Variant
    For Each key In globalActualDictForecastCols.keys
        rowData = globalActualDictForecastCols(key)

        ' Check if this row matches the adjustment criteria
        Dim isMatch As Boolean
        isMatch = True

        ' Only check fields that have values in the adjustment row
        If company <> "" And rowData(3) <> company Then isMatch = False
        If CostCenter <> "" And rowData(4) <> CostCenter Then isMatch = False
        If customerSegment <> "" And rowData(5) <> customerSegment Then isMatch = False
        If venue <> "" And rowData(6) <> venue Then isMatch = False
        If ledgerAccount <> "" And rowData(7) <> ledgerAccount Then isMatch = False
        If revCat <> "" And rowData(8) <> revCat Then isMatch = False
        If spendCat <> "" And rowData(9) <> spendCat Then isMatch = False
        If usGaap <> "" And rowData(10) <> usGaap Then isMatch = False

        ' If match found, update the target column
        If isMatch Then
            rowData(targetCol + 1) = updateValue ' +1 because row(1) is row number, data starts at row(2)
            globalActualDictForecastCols(key) = rowData

            ' Track this column as updated
            Dim updateKey As String
            updateKey = key & "|" & targetCol
            forecastUpdatedCols(updateKey) = True
            Debug.Print "Actual Adjustment applied to key: " & Left(key, 30) & "... | TargetCol=" & targetCol & " | Value=" & updateValue
        End If
    Next key

End Sub


Private Sub UpdateActualDictionaryValue(globalActualDictForecastCols As Object, lookupValue As String, updateValue As Variant, targetCol As Long, company As String, CostCenter As String, customerSegment As String, venue As String, ledgerAccount As String, revCat As String, spendCat As String, usGaap As String, ByRef maxRowActual As Long, forecastUpdatedCols As Object)

    ' Update or create dictionary entry for actual data
    If globalActualDictForecastCols.Exists(lookupValue) Then

        ' Update existing entry
        Dim existingRow As Variant
        existingRow = globalActualDictForecastCols(lookupValue)
        existingRow(targetCol + 1) = updateValue ' +1 because row(1) is row number, data starts at row(2)
        globalActualDictForecastCols(lookupValue) = existingRow

        ' Track this column as updated
        Dim updateKey As String
        updateKey = lookupValue & "|" & targetCol
        forecastUpdatedCols(updateKey) = True

        ' Debug output for actual updates

        'Static actualMatchCount As Long

        'If actualMatchCount < 10 Then

            'actualMatchCount = actualMatchCount + 1

            'Debug.Print "ACTUAL MATCH " & actualMatchCount & ": Key=" & Left(lookupValue, 30) & "... | Row=" & existingRow(1) & " | TargetCol=" & targetCol & " | Value=" & updateValue

        'End If
    Else

        ' Create new entry
        If updateValue <> 0 Then
            maxRowActual = maxRowActual + 1
            Dim newRow(1 To 23) As Variant ' Row + 22 columns (A to V)
            newRow(1) = maxRowActual ' Row number
            newRow(2) = lookupValue  ' Column A
            newRow(3) = company      ' Column B
            newRow(4) = CostCenter   ' Column C
            newRow(5) = customerSegment ' Column D
            newRow(6) = venue        ' Column E
            newRow(7) = ledgerAccount ' Column F
            newRow(8) = revCat       ' Column G
            newRow(9) = spendCat     ' Column H
            newRow(10) = usGaap      ' Column I

            ' Initialize J to V (columns 10-22) with empty values to preserve existing data
            For i = 11 To 23
                newRow(i) = "" ' Use empty string instead of 0 to avoid overwriting existing data
            Next i
            newRow(targetCol + 1) = updateValue ' Only set the specific forecast column
            globalActualDictForecastCols(lookupValue) = newRow

            ' Track this column as updated
            updateKey = lookupValue & "|" & targetCol
            forecastUpdatedCols(updateKey) = True

            ' Debug output for first 10 new actual rows

            'Static actualNewRowCount As Long

            'If actualNewRowCount < 10 Then

                'actualNewRowCount = actualNewRowCount + 1

                'Debug.Print "ACTUAL NEW " & actualNewRowCount & ": Key=" & Left(lookupValue, 30) & "... | Row=" & maxRowActual & " | TargetCol=" & targetCol & " | Value=" & updateValue & " | Company=" & company

            'End If
        End If
    End If

End Sub
Private Sub WriteBulkActualData(globalActualDict As Object, maxRowActual As Long, forecastUpdatedCols As Object, actualColStart As Long, actualColEnd As Long, forecastColStart As Long, forecastColEnd As Long)

    ' Create bulk array for all data
    Dim dictSize As Long
    dictSize = globalActualDict.Count
    If dictSize = 0 Then Exit Sub
    
    ' Prepare bulk data array (rows x columns)
    Dim bulkDataArray() As Variant
    ReDim bulkDataArray(1 To dictSize, 1 To 21) ' A to U columns
    
    ' Fill array with dictionary data
    Dim i As Long, j As Long
    i = 0
    For Each key In globalActualDict.keys
        i = i + 1
        Dim rowData As Variant
        rowData = globalActualDict(key)
        
        ' Copy data to bulk array (columns A to U = indices 2 to 22 in rowData)
        For j = 1 To 21
            bulkDataArray(i, j) = rowData(j + 1)
        Next j
        
        ' Debug print first 20 lines by column during bulk paste
        If i <= 20 Then
            Dim lineDebug As String
            lineDebug = "Bulk paste line " & i & ": "
            For j = 1 To 21
                lineDebug = lineDebug & Chr(64 + j) & "=" & bulkDataArray(i, j) & " "
            Next j
            Debug.Print lineDebug
        End If
    Next key
    
    ' Determine target range starting from row 12
    Dim targetRange As Range
    Set targetRange = globalActualFYSheet.Range("A12").Resize(dictSize, 21)
    
    ' Bulk paste all data
    targetRange.value = bulkDataArray

End Sub
    


Public Sub UpdateActualFYWithAllocations(filter5YP As String, sourceSheet As Worksheet)

    ' Updated Aug 2025 - For updating Actual-FY sheet with forecast allocations

    ' Based on UpdateBudgetFYWithAllocations but targets wsActual only

    ' Uses "Column Start Forecast>>" markers and only updates Forecast columns
    Call declareGlobal
    Dim wsActual As Worksheet
    Dim wsExclusion As Worksheet
    Dim lastRowActual As Long, lastRowExclusion As Long
    Dim dictDept As Object
    Dim i As Long, j As Long, k As Long
    Dim ytdAverage As Variant, department As String
    Dim allocArr As Variant
    Dim processedItems As Long, totalItems As Long
    Dim startTime As Double, elapsedTime As Double
    startTime = Timer
    Set wsActual = ThisWorkbook.Sheets("Actual-FY")
    Set wsExclusion = ThisWorkbook.Sheets("Exclusions")
    sourceSheet.Calculate
    wsActual.Calculate
    wsExclusion.Calculate

    ' --- Load exclusionArr from UpdateExclusion sheet ---
    lastRowExclusion = wsExclusion.Cells(wsExclusion.Rows.Count, "A").End(xlUp).row
    Dim tempArrExcl() As Variant
    Dim exclusionArr() As Variant
    Dim exclusionCount As Long
    exclusionCount = 0
    For i = 2 To lastRowExclusion
        If wsExclusion.Cells(i, 1).value = filter5YP Then
            exclusionCount = exclusionCount + 1
            ReDim Preserve tempArrExcl(1 To exclusionCount)
            Dim rowArr(1 To 9) As Variant
            For j = 1 To 9
                If IsEmpty(wsExclusion.Cells(i, j).value) Or wsExclusion.Cells(i, j).value = "" Then
                    rowArr(j) = ""
                Else
                    rowArr(j) = wsExclusion.Cells(i, j).value
                End If
            Next j
            tempArrExcl(exclusionCount) = rowArr
        End If
    Next i
    If exclusionCount > 0 Then
        ReDim exclusionArr(1 To exclusionCount, 1 To 9)
        For i = 1 To exclusionCount
            For j = 1 To 9
                exclusionArr(i, j) = tempArrExcl(i)(j)
            Next j
        Next i
    Else
        ReDim exclusionArr(1 To 1, 1 To 9)
        For j = 1 To 9
            exclusionArr(1, j) = ""
        Next j
    End If

    ' --- Step 1: Build allocation dictionary from sourceSheet using Forecast markers ---
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).row
    Set dictDept = CreateObject("Scripting.Dictionary")
    totalItems = lastRowSource - 9
    processedItems = 0
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Building Forecast Allocation..."
    Progress_Bar.percentDone 0
    For i = 10 To lastRowSource
        If sourceSheet.Cells(i, "A").value = "Allocated" Then
            ' Check if row should be excluded (DO NOT REFORECAST in column AI)
            If UCase(Trim(CStr(sourceSheet.Cells(i, "AI").value))) = "DO NOT REFORECAST" Then
                Debug.Print "Excluded row " & i & " from UpdateActualFYWithAllocations due to DO NOT REFORECAST in column AI"
                GoTo NextRowActual
            End If
            
            department = sourceSheet.Cells(i, 11).value
            For j = 1 To sourceSheet.Columns.Count
                If sourceSheet.Cells(i, j).value = "Column Start Forecast>>" Then
                    Exit For
                End If
            Next j
            If j <= sourceSheet.Columns.Count Then
                Dim tempArr(1 To 12) As Double
                For k = 0 To 11
                    tempArr(k + 1) = sourceSheet.Cells(i, j + k + 1).value
                Next k
                dictDept(department) = tempArr
            End If
        End If
        processedItems = processedItems + 1
        If processedItems Mod 100 = 0 Or processedItems = totalItems Then
            Progress_Bar.percentDone processedItems / totalItems
            Progress_Bar.setStatusText "Building Forecast Allocation: " & processedItems & " of " & totalItems
            DoEvents
        End If
NextRowActual:
    Next i

    ' --- Step 2: Detect forecast columns from sourceSheet row 7 (Aj7 to Au7) ---
    Dim forecastSourceData As Variant
    forecastSourceData = sourceSheet.Range("AJ7:AU7").value ' Columns 36-47 in sourceSheet
    Dim forecastCols() As Boolean
    ReDim forecastCols(10 To 21) ' J=10 to U=21 in wsActual
    For i = 1 To UBound(forecastSourceData, 2)
        If forecastSourceData(1, i) = "Forecast" Then
            Dim actualCol As Long
            actualCol = 9 + i ' AJ=36 maps to J=10, so AJ=10, AK=11, etc.
            If actualCol >= 10 And actualCol <= 21 Then
                forecastCols(actualCol) = True
                Debug.Print "Detected forecast column in wsActual: " & Chr(64 + actualCol) & " (column " & actualCol & ")"
            End If
        End If
    Next i

    ' --- Step 3: Update wsActual using allocation dictionary, skip excluded rows, only update forecast columns ---
    Dim lastRowActualFY As Long
    lastRowActualFY = wsActual.Cells(wsActual.Rows.Count, "W").End(xlUp).row
    totalItems = lastRowActualFY - 1
    processedItems = 0
    Progress_Bar.setStatusText "Populating Actual-FY forecast columns..."
    Progress_Bar.percentDone 0
    For i = 2 To lastRowActualFY

        ' Gather candidate values from Actual-FY row
        Dim candidate(1 To 9) As Variant
        For k = 1 To 9
            candidate(k) = wsActual.Cells(i, k).value ' A=1, ..., I=9
        Next k

        ' Check exclusionArr for multi-field AND match
        Dim isExcluded As Boolean
        isExcluded = False
        For x = 1 To exclusionCount
            Dim allMatch As Boolean
            allMatch = True
            For y = 2 To 9 ' Skip column 1 (Type)
                If exclusionArr(x, y) <> "" Then
                    If candidate(y) <> exclusionArr(x, y) Then
                        allMatch = False
                        Exit For
                    End If
                End If
            Next y
            If allMatch Then
                isExcluded = True
                Exit For
            End If
        Next x
        If Not isExcluded Then
            If wsActual.Cells(i, "W").value = filter5YP Then
                ytdAverage = wsActual.Cells(i, "V").value
                If IsEmpty(ytdAverage) Or Not IsNumeric(ytdAverage) Then ytdAverage = 0
                department = wsActual.Cells(i, "Y").value
                If dictDept.Exists(department) Then
                    allocArr = dictDept(department)
                    For j = 10 To 21 ' Columns J to U

                        ' Only update if this column is detected as forecast
                        If forecastCols(j) Then
                            wsActual.Cells(i, j).value = Round((1 + allocArr(j - 9)) * ytdAverage, 6)
                        End If
                    Next j
                End If
            End If
        End If
        processedItems = processedItems + 1
        If processedItems Mod 100 = 0 Or processedItems = totalItems Then
            Progress_Bar.percentDone processedItems / totalItems
            Progress_Bar.setStatusText "Populating Actual-FY forecast: " & processedItems & " of " & totalItems
            DoEvents
        End If
    Next i
    sourceSheet.Calculate
    wsActual.Calculate
    wsExclusion.Calculate

    'Set as values the formulas
    Call SetGlobalFYSheetValues
    elapsedTime = Timer - startTime
    Unload Progress_Bar
    MsgBox "Actual-FY sheet forecast columns have been populated successfully!" & vbCrLf & _
           "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds", vbInformation

End Sub

Public Sub UpdateActualFYWithAllocationsNoDept(filterArr As Variant, sourceSheet As Worksheet, typeTrans As String)

    ' Updated Aug 2025 - For updating Actual-FY sheet with forecast allocations (NoDept version)

    ' Based on UpdateBudgetFYWithAllocationsNoDept but targets wsActual only

    ' Uses "Column Start Forecast>>" markers and only updates Forecast columns
    Call declareGlobal
    Dim wsActual As Worksheet
    Dim wsExclusion As Worksheet
    Dim lastRowActual As Long, lastRowSource As Long, lastRowExclusion As Long
    Dim dictAlloc As Object
    Dim i As Long, j As Long, k As Long, filterIdx As Long
    Dim ytdAverage As Variant, department As String
    Dim processedItems As Long, totalItems As Long
    Dim startTime As Double, elapsedTime As Double
    Dim rowIndex As Long
    Dim colStart As Long
    Dim isExcluded As Boolean, allMatch As Boolean
    startTime = Timer

    ' Set worksheet references for performance
    Set wsActual = ThisWorkbook.Sheets("Actual-FY")
    Set wsExclusion = ThisWorkbook.Sheets("Exclusions")

    ' Calculate all sheets to ensure formulas are up to date
    sourceSheet.Calculate
    wsActual.Calculate
    wsExclusion.Calculate

    ' --- Load exclusionArr from UpdateExclusion sheet for matching type ---
    lastRowExclusion = wsExclusion.Cells(wsExclusion.Rows.Count, "A").End(xlUp).row
    Dim tempArr() As Variant
    Dim exclusionArr() As Variant
    Dim exclusionCount As Long
    exclusionCount = 0

    ' Collect matching rows into tempArr as array of arrays
    For i = 2 To lastRowExclusion
        If wsExclusion.Cells(i, 1).value = typeTrans Then
            exclusionCount = exclusionCount + 1
            ReDim Preserve tempArr(1 To exclusionCount)
            Dim rowArr(1 To 9) As Variant
            For j = 1 To 9
                If IsEmpty(wsExclusion.Cells(i, j).value) Or wsExclusion.Cells(i, j).value = "" Then
                    rowArr(j) = ""
                Else
                    rowArr(j) = wsExclusion.Cells(i, j).value
                End If
            Next j
            tempArr(exclusionCount) = rowArr
        End If
    Next i

    ' Copy tempArr to a 2D array exclusionArr
    If exclusionCount > 0 Then
        ReDim exclusionArr(1 To exclusionCount, 1 To 9)
        For i = 1 To exclusionCount
            For j = 1 To 9
                exclusionArr(i, j) = tempArr(i)(j)
            Next j
        Next i
    Else
        ReDim exclusionArr(1 To 1, 1 To 9)
        For j = 1 To 9
            exclusionArr(1, j) = ""
        Next j
    End If

    ' --- Build exclusionLookupArr for filterArr removal ---
    Dim exclusionLookupArr() As String
    If exclusionCount > 0 Then
        ReDim exclusionLookupArr(1 To exclusionCount)
        For i = 1 To exclusionCount
            Dim lookupStr As String
            lookupStr = ""
            For j = 2 To 9 ' Skip Type
                If exclusionArr(i, j) <> "" Then
                    lookupStr = lookupStr & "|" & exclusionArr(i, j)
                End If
            Next j
            exclusionLookupArr(i) = lookupStr
        Next i
    Else
        ReDim exclusionLookupArr(1 To 1)
        exclusionLookupArr(1) = ""
    End If

    ' --- Remove matching values from filterArr ---
    Dim newFilterArr() As Variant
    Dim newCount As Long
    newCount = 0
    For i = LBound(filterArr) To UBound(filterArr)
        Dim isExcluded1 As Boolean
        isExcluded1 = False
        For j = 1 To exclusionCount
            If filterArr(i) = exclusionLookupArr(j) Then
                isExcluded1 = True
                Exit For
            End If
        Next j
        If Not isExcluded1 Then
            newCount = newCount + 1
            ReDim Preserve newFilterArr(1 To newCount)
            newFilterArr(newCount) = filterArr(i)
        End If
    Next i
    filterArr = newFilterArr

    ' Debug.Print first 5 rows of exclusionArr
    For i = 1 To Application.Min(5, exclusionCount)
        Dim debugStr As String
        debugStr = ""
        For j = 1 To 9
            debugStr = debugStr & exclusionArr(i, j) & ";"
        Next j
        Debug.Print "ExclusionArr Row " & i & ": " & debugStr
    Next i

    ' Initialize progress tracking for 2 main steps
    totalItems = 2
    processedItems = 0
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Step 1: Building Forecast Allocation..."
    Progress_Bar.percentDone processedItems / totalItems

    ' --- Step 1: Build allocation dictionary from sourceSheet using Forecast markers ---
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).row
    Set dictAlloc = CreateObject("Scripting.Dictionary")
    Debug.Print "=== UpdateActualFYWithAllocationsNoDept - Building Allocation Dictionary ==="
    Debug.Print "lastRowSource: " & lastRowSource
    Debug.Print "typeTrans: " & typeTrans
    Debug.Print "filterArr count: " & (UBound(filterArr) - LBound(filterArr) + 1)
    For i = LBound(filterArr) To UBound(filterArr)
        Debug.Print "  filterArr(" & i & "): " & filterArr(i)
    Next i
    For i = 10 To lastRowSource
        If sourceSheet.Cells(i, "A").value = "Allocated" Then
            ' Check if row should be excluded (DO NOT REFORECAST in column AI)
            If UCase(Trim(CStr(sourceSheet.Cells(i, "AI").value))) = "DO NOT REFORECAST" Then
                Debug.Print "Excluded row " & i & " from UpdateActualFYWithAllocationsNoDept due to DO NOT REFORECAST in column AI"
                GoTo NextRowAllocNoDept
            End If

            ' NoDept version: Use the filter value directly, not column K (department)

            ' The allocation should match the filterArr values (like "OPEX", "PAYROLL", etc.)
            Dim allocationType As String
            allocationType = sourceSheet.Cells(i, 11).value ' Column K contains the allocation type
            Debug.Print "Row " & i & ": Found Allocated row, Column K value: '" & allocationType & "'"

            ' Check if this allocation type matches any of our filter values
            Dim matchesFilter As Boolean
            matchesFilter = False
            For filterIdx = LBound(filterArr) To UBound(filterArr)
                If allocationType = filterArr(filterIdx) Then
                    matchesFilter = True
                    Debug.Print "  Matches filter: " & filterArr(filterIdx)
                    Exit For
                End If
            Next filterIdx
            If matchesFilter Then

                ' Find forecast columns
                colStart = 0
                For j = 1 To sourceSheet.Columns.Count
                    If sourceSheet.Cells(i, j).value = "Column Start Forecast>>" Then
                        colStart = j + 1
                        Debug.Print "  Found forecast start at column: " & colStart
                        Exit For
                    End If
                Next j
                If colStart > 0 Then
                    Dim tempArrAlloc(1 To 12) As Double
                    For k = 0 To 11
                        tempArrAlloc(k + 1) = sourceSheet.Cells(i, colStart + k).value
                        If k < 3 Then ' Debug first 3 values
                            Debug.Print "    tempArrAlloc(" & (k + 1) & "): " & tempArrAlloc(k + 1)
                        End If
                    Next k
                    dictAlloc(allocationType) = tempArrAlloc
                    Debug.Print "  Added to dictAlloc with key: '" & allocationType & "'"
                Else
                    Debug.Print "  ERROR: No forecast start marker found"
                End If
            Else
                Debug.Print "  Skipped - allocation type '" & allocationType & "' does not match any filter"
            End If
        End If
        If i Mod 100 = 0 Then DoEvents
NextRowAllocNoDept:
    Next i
    Debug.Print "Final dictAlloc count: " & dictAlloc.Count
    For Each key In dictAlloc.keys
        Debug.Print "  dictAlloc key: '" & key & "'"
    Next key
    processedItems = processedItems + 1
    Progress_Bar.percentDone processedItems / totalItems
    Progress_Bar.setStatusText "Step 2: Detecting forecast columns and populating Actual-FY..."

    ' --- Step 2: Detect forecast columns from sourceSheet row 7 (Aj7 to Au7) ---
    Dim forecastSourceData As Variant
    forecastSourceData = sourceSheet.Range("AJ7:AU7").value ' Columns 36-47 in sourceSheet
    Dim forecastCols() As Boolean
    ReDim forecastCols(10 To 21) ' J=10 to U=21 in wsActual
    Debug.Print "=== UpdateActualFYWithAllocationsNoDept - Forecast Column Detection ==="
    Debug.Print "sourceSheet.Range('AJ7:AU7') analysis:"
    Dim forecastColCount As Long
    forecastColCount = 0
    For i = 1 To UBound(forecastSourceData, 2)
        Dim sourceCol As Long
        sourceCol = 35 + i ' AJ=36, AK=37, etc.
        Debug.Print "  SourceSheet col " & Chr(65 + sourceCol - 1) & sourceCol & " (index " & i & "): '" & forecastSourceData(1, i) & "'"
        If forecastSourceData(1, i) = "Forecast" Then
            Dim actualCol As Long
            actualCol = 9 + i ' AJ=36 maps to J=10, so AJ=10, AK=11, etc.
            If actualCol >= 10 And actualCol <= 21 Then
                forecastCols(actualCol) = True
                forecastColCount = forecastColCount + 1
                Debug.Print "    -> Mapped to wsActual column " & Chr(64 + actualCol) & " (" & actualCol & ") - FORECAST"
            End If
        End If
    Next i
    Debug.Print "Total forecast columns detected: " & forecastColCount
    If forecastColCount = 0 Then
        Debug.Print "WARNING: No forecast columns detected! This may be why no updates are happening."
    End If

    ' --- Step 3: Update wsActual using allocation dictionary, skip excluded rows, only update forecast columns ---
    Dim lastRowActualFY As Long
    Dim allocArr As Variant
    lastRowActualFY = wsActual.Cells(wsActual.Rows.Count, "W").End(xlUp).row
    Debug.Print "=== UpdateActualFYWithAllocationsNoDept - Updating Rows ==="
    Debug.Print "lastRowActualFY: " & lastRowActualFY
    Debug.Print "exclusionCount: " & exclusionCount
    Dim rowsProcessed As Long, rowsUpdated As Long
    rowsProcessed = 0
    rowsUpdated = 0
    For i = 2 To lastRowActualFY

        ' Gather candidate values from Actual-FY row
        Dim candidate(1 To 9) As Variant
        For k = 1 To 9
            candidate(k) = wsActual.Cells(i, k).value ' A=1, ..., I=9
        Next k

        ' Check exclusionArr for multi-field AND match
        isExcluded = False
        For x = 1 To exclusionCount
            allMatch = True
            For y = 2 To 9 ' Skip column 1 (Type)
                If exclusionArr(x, y) <> "" Then
                    If candidate(y) <> exclusionArr(x, y) Then
                        allMatch = False
                        Exit For
                    End If
                End If
            Next y
            If allMatch Then
                isExcluded = True
                Exit For
            End If
        Next x
        If Not isExcluded Then
            Dim currentRowValue As String
            currentRowValue = wsActual.Cells(i, "W").value ' Column W - 5 YP plan mapping

            ' Debug first 10 rows
            If rowsProcessed < 10 Then
                Debug.Print "Row " & i & ": Column W value = '" & currentRowValue & "'"
            End If
            For j = LBound(filterArr) To UBound(filterArr)
                If currentRowValue = filterArr(j) Then
                    ytdAverage = wsActual.Cells(i, "V").value
                    If IsEmpty(ytdAverage) Or Not IsNumeric(ytdAverage) Then ytdAverage = 0

                    ' Use filterArr(j) as the lookup key for allocation
                    department = filterArr(j)
                    If rowsProcessed < 10 Then
                        Debug.Print "  Matched filter: " & filterArr(j) & ", ytdAverage: " & ytdAverage
                        Debug.Print "  Looking for dictAlloc key: '" & department & "'"
                        Debug.Print "  dictAlloc.Exists('" & department & "'): " & dictAlloc.Exists(department)
                    End If
                    If dictAlloc.Exists(department) Then
                        allocArr = dictAlloc(department)
                        Dim columnsUpdated As Long
                        columnsUpdated = 0
                        For colStart = 10 To 21 ' Columns J to U

                            ' Only update if this column is detected as forecast
                            If forecastCols(colStart) Then
                                Dim newValue As Double
                                newValue = Round((1 + allocArr(colStart - 9)) * ytdAverage, 6)
                                wsActual.Cells(i, colStart).value = newValue
                                columnsUpdated = columnsUpdated + 1
                                If rowsProcessed < 5 And columnsUpdated <= 3 Then
                                    Debug.Print "    Updated col " & Chr(64 + colStart) & " (" & colStart & ") = " & newValue & " [formula: (1 + " & allocArr(colStart - 9) & ") * " & ytdAverage & "]"
                                End If
                            End If
                        Next colStart
                        If rowsProcessed < 10 Then
                            Debug.Print "  Updated " & columnsUpdated & " forecast columns for row " & i
                        End If
                        rowsUpdated = rowsUpdated + 1
                    Else
                        If rowsProcessed < 10 Then
                            Debug.Print "  ERROR: dictAlloc does not contain key '" & department & "'"
                        End If
                    End If
                    Exit For
                End If
            Next j
        Else
            If rowsProcessed < 5 Then
                Debug.Print "Row " & i & ": Excluded from processing"
            End If
        End If
        rowsProcessed = rowsProcessed + 1
        If i Mod 100 = 0 Then DoEvents
    Next i
    Debug.Print "=== Final Results ==="
    Debug.Print "Total rows processed: " & rowsProcessed
    Debug.Print "Total rows updated: " & rowsUpdated
    processedItems = processedItems + 1
    Progress_Bar.percentDone processedItems / totalItems
    Progress_Bar.setStatusText "Done!"
    sourceSheet.Calculate
    wsActual.Calculate
    wsExclusion.Calculate

    'Set as values the formulas
    Call SetGlobalFYSheetValues
    elapsedTime = Timer - startTime
    Unload Progress_Bar
    MsgBox "Actual-FY sheet forecast columns have been populated successfully!" & vbCrLf & _
           "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds", vbInformation

End Sub








