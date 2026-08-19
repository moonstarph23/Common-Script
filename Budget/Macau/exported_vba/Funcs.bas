Attribute VB_Name = "Funcs"
Public Sub ForceFullCalculation()

    On Error Resume Next

    ' Normal calculation
    Call Application.Calculate

    ' Extended calculation
    Call Application.CalculateFull

    ' Full Calculation

    'Call Application.CalculateFullRebuild

End Sub

Sub SetUpDataValidation()

    Dim ws As Worksheet
    Dim cell As Range
    Dim validationColumn As Long
    Dim validationRange As Range
    Dim colLetter As String
    Dim firstRow As Long
    Dim lastRow As Long

    ' Set the first and last row for the data validation range
    firstRow = 10
    lastRow = 100

    ' Set the worksheet object. Replace "Sheet1" with the name of your sheet.
    Set ws = ThisWorkbook.Sheets("(A) Opex")

    ' Loop from row 13 to 300 in column Q
    For Each cell In ws.Range("Q13:Q300")

        ' Calculate the validation column offset based on the current cell row
        validationColumn = 39 + cell.row - 12 ' AN is the 40th column, offset starts from row 13

        ' Use Cells to specify the range for validation
        Set validationRange = ws.Cells(firstRow, validationColumn).Resize(lastRow - firstRow + 1)

        ' Get column letter of validation range for debug purposes
        colLetter = Split(ws.Cells(, validationColumn).Address, "$\")(1)

        ' Clear any existing validations
        cell.Validation.Delete

        ' Add new validation
        With cell.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 operator:=xlBetween, Formula1:="=" & colLetter & firstRow & ":" & colLetter & lastRow
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Next cell
    MsgBox "Data validation setup complete!"

End Sub

Sub FilterActualFY()

    Dim wsCriteria As Worksheet
    Dim wsData As Worksheet
    Dim lastRowCriteria As Long
    Dim critRange As Range
    Dim cell As Range
    Dim filterArray As Variant
    Dim dict As Object
    Dim i As Long

    ' Set references to the worksheets
    Set wsCriteria = ThisWorkbook.Sheets("Sheet1")
    Set wsData = ThisWorkbook.Sheets("Actual-FY")

    ' Find the last row in Sheet1 Column A
    lastRowCriteria = wsCriteria.Cells(wsCriteria.Rows.Count, "A").End(xlUp).row
    Set critRange = wsCriteria.Range("A1:A" & lastRowCriteria)

    ' Initialize dictionary
    Set dict = CreateObject("Scripting.Dictionary")

    ' Collect unique filter values from Sheet1 Column A
    For Each cell In critRange
        If Not dict.Exists(cell.value) And Not IsEmpty(cell.value) Then
            dict.Add cell.value, Nothing
        End If
    Next cell

    ' Only proceed if there are items in the dictionary
    If dict.Count > 0 Then
        filterArray = dict.keys
    Else
        MsgBox "No data to filter."
        Exit Sub
    End If

    ' Make sure the data range has a filter applied
    If Not wsData.AutoFilterMode Then
        wsData.Range("F1:F" & wsData.Cells(wsData.Rows.Count, "F").End(xlUp).row).AutoFilter
    End If

    ' Set the filter on Column F in Actual-FY for the collected values
    wsData.AutoFilterMode = False
    wsData.Range("F1:F" & wsData.Cells(wsData.Rows.Count, "F").End(xlUp).row).AutoFilter Field:=1, Criteria1:=filterArray, operator:=xlFilterValues

End Sub

Sub AddXMLReference()

    Dim vbProj As Object
    Dim chkRef As Object
    Dim refAlreadySet As Boolean

    ' Reference to the VBA Project
    Set vbProj = ThisWorkbook.VBProject

    ' Check if the Microsoft XML reference is already set
    refAlreadySet = False
    For Each chkRef In vbProj.References
        If chkRef.Name = "MSXML2" Then
            refAlreadySet = True
            Exit For
        End If
    Next chkRef

    ' Add the reference if not already set
    If Not refAlreadySet Then
        vbProj.References.AddFromGuid "{F5078F18-C551-11D3-89B9-0000F81FE221}", 6, 0
        MsgBox "Microsoft XML, v6.0 reference has been added.", vbInformation
    Else

        'MsgBox "Microsoft XML, v6.0 reference is already set.", vbInformation
    End If

End Sub

Sub choosePL()

    viewPL.Show

End Sub

Sub openTOC()

    tableContents.Show

End Sub

Sub addItemsForm()

    addItems.Show

End Sub

Public Sub SetGlobalFYSheetValues()

    ' Call global variables at the top as per coding instructions
    Call declareGlobal

    ' Determine the last used row in column F of the GlobalFY sheet
    Dim wsGlobalFY As Worksheet
    Set wsGlobalFY = globalBudgetFYSheet
    Dim lastRow As Long
    lastRow = wsGlobalFY.Cells(wsGlobalFY.Rows.Count, "F").End(xlUp).row

    ' Convert formulas to values in columns W12:Y[lastRow] and AC12:AJ[lastRow]
    With wsGlobalFY
        .Range("W12:Y" & lastRow).value = .Range("W12:Y" & lastRow).value
        .Range("AB12:AJ" & lastRow).value = .Range("AB12:AJ" & lastRow).value
    End With

End Sub

Public Sub groupData(sheetToGroup As String)

    ' Use global variables and call them at the top
    Call declareGlobal

    ' Disable calculation and screen updating for performance
    Application.Calculation = xlManual
    Application.ScreenUpdating = False

    ' Set reference to the target worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetToGroup)

    ' Determine the last used row in column F
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' Load data from A12:U[lastRow] into an array
    Dim dataArr As Variant
    dataArr = ws.Range("A12:U" & lastRow).value

    ' Create a dictionary to group data by column A
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long
    For i = 1 To UBound(dataArr, 1)
        Dim key As Variant
        key = dataArr(i, 1) ' Column A
        If Not dict.Exists(key) Then

            ' Store the first row as the base
            dict(key) = dataArr(i, 1)
            For j = 2 To 21
                dict(key & "_col" & j) = dataArr(i, j)
            Next j
        Else

            ' For columns J to U (columns 10 to 21), sum the values
            For j = 10 To 21
                dict(key & "_col" & j) = dict(key & "_col" & j) + dataArr(i, j)
            Next j
        End If
    Next i

    ' Clear all data from A12:AR[lastRow]
    ws.Range("A12:AR" & lastRow).ClearContents

    ' Paste grouped data back starting at A12
    Dim groupCount As Long
    groupCount = dict.Count / 21 ' Each group has 21 columns
    If groupCount > 0 Then
        Dim groupedArr() As Variant
        ReDim groupedArr(1 To groupCount, 1 To 21)
        Dim idx As Long
        idx = 1
        For Each key In dict.keys
            If InStr(key, "_col") = 0 Then
                groupedArr(idx, 1) = dict(key)
                For j = 2 To 21
                    groupedArr(idx, j) = dict(key & "_col" & j)
                Next j
                idx = idx + 1
            End If
        Next key
        ws.Range("A12:U" & 11 + groupCount).value = groupedArr
    End If

    ' Determine lastRow again after grouping
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row

    ' Copy formulas from V10:AJ10 to each row V12:AJ[lastRow] by row
    Dim r As Long
    For r = 12 To lastRow
        ws.Range("V" & r & ":AJ" & r).FormulaR1C1 = ws.Range("V10:AJ10").FormulaR1C1
    Next r
    ws.Calculate

    ' Convert columns V to AJ to values except columns V and Z
    Dim colNum As Long
    For colNum = ws.Range("V1").Column To ws.Range("AJ1").Column
        If colNum <> ws.Range("V1").Column And colNum <> ws.Range("Z1").Column Then
            ws.Range(ws.Cells(12, colNum), ws.Cells(lastRow, colNum)).value = ws.Range(ws.Cells(12, colNum), ws.Cells(lastRow, colNum)).value
        End If
    Next colNum

    ' Re-enable calculation and screen updating
    Application.ScreenUpdating = True

End Sub

Sub groupActual()

    ' Use global variables and call them at the top
    Call declareGlobal
    Call groupData(globalActualFYSheet.Name)

End Sub

Sub groupBudgetFY()

    ' Use global variables and call them at the top
    Call declareGlobal
    Call groupData(globalBudgetFYSheet.Name)

End Sub

