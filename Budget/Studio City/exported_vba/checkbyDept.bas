Attribute VB_Name = "checkbyDept"
Function getActualFYSummaryDict() As Object
    Call declareGlobal
    Dim ws As Worksheet
    Set ws = globalActualFYSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Z").End(xlUp).row
    Dim summaryDeptDict As Object
    Set summaryDeptDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim summary As String, dept As String
    Dim amount As Double
    For i = 12 To lastRow
        summary = ws.Cells(i, "X").value
        dept = ws.Cells(i, "Y").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        Dim columnFValue As String
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            Dim firstChar As String
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If summary <> "" And summary <> "0" Then
            Dim key As String
            key = dept & "|" & summary
            If summaryDeptDict.Exists(key) Then
                summaryDeptDict(key) = summaryDeptDict(key) + amount
            Else
                summaryDeptDict.Add key, amount
            End If
        End If
    Next i

    ' Debug print first 10 items for BLANK DEPARTMENT only
    Dim debugCount As Long
    debugCount = 0
    Dim k As Variant
    For Each k In summaryDeptDict.keys
        If Left(k, 15) = "BLANK DEPARTMENT" Then
            Debug.Print "SummaryDict Key: " & k & " | Amount: " & summaryDeptDict(k)
            debugCount = debugCount + 1
            If debugCount >= 10 Then Exit For
        End If
    Next k

    ' Calculate totals for message box
    Dim totalRevenue As Double, totalExpenses As Double, ebitda As Double
    Dim corporateRecharge As Double, brandingFee As Double, sbc As Double
    Dim totalNonOperating As Double, netIncome As Double
    totalRevenue = 0
    totalExpenses = 0
    ebitda = 0
    corporateRecharge = 0
    brandingFee = 0
    sbc = 0
    totalNonOperating = 0
    netIncome = 0

    ' Aggregate totals (sum across all departments)
    For Each k In summaryDeptDict.keys
        If InStr(1, k, "Total Revenue", vbTextCompare) > 0 Then totalRevenue = totalRevenue + summaryDeptDict(k)
        If InStr(1, k, "Total Expenses", vbTextCompare) > 0 Then totalExpenses = totalExpenses + summaryDeptDict(k)
        If InStr(1, k, "Corporate Recharge", vbTextCompare) > 0 Then corporateRecharge = corporateRecharge + summaryDeptDict(k)
        If InStr(1, k, "Branding Fee", vbTextCompare) > 0 Then brandingFee = brandingFee + summaryDeptDict(k)
        If InStr(1, k, "SBC", vbTextCompare) > 0 Then sbc = sbc + summaryDeptDict(k)
        If InStr(1, k, "Total Non-Operating Costs", vbTextCompare) > 0 Then totalNonOperating = totalNonOperating + summaryDeptDict(k)
    Next k

    ' Revised EBITDA and Net Income calculation
    ebitda = totalRevenue + totalExpenses
    netIncome = ebitda + corporateRecharge + brandingFee + sbc + totalNonOperating

    ' Build message string
    Dim msg As String
    msg = "Total Revenue: " & Format(totalRevenue, "#,##0.00") & vbCrLf
    msg = msg & "Total Expenses: " & Format(totalExpenses, "#,##0.00") & vbCrLf
    msg = msg & "EBITDA: " & Format(ebitda, "#,##0.00") & vbCrLf & vbCrLf
    msg = msg & "Corporate Recharge: " & Format(corporateRecharge, "#,##0.00") & vbCrLf
    msg = msg & "Branding Fee: " & Format(brandingFee, "#,##0.00") & vbCrLf
    msg = msg & "SBC: " & Format(sbc, "#,##0.00") & vbCrLf
    msg = msg & "Total Non-Operating Costs: " & Format(totalNonOperating, "#,##0.00") & vbCrLf & vbCrLf
    msg = msg & "Net Income: " & Format(netIncome, "#,##0.00") & vbCrLf & vbCrLf
    'MsgBox msg, vbInformation, "Summary Totals"

    ' --- Create new dictionary getCompanyTotal, group by Company (AJ) and Summary (X) ---
    Dim companyTotalDict As Object
    Set companyTotalDict = CreateObject("Scripting.Dictionary")
    Dim company As String
    For i = 12 To lastRow
        summary = ws.Cells(i, "X").value
        company = ws.Cells(i, "AJ").value
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If summary <> "" And summary <> "0" Then
            Dim compKey As String
            compKey = company & "|" & summary
            If companyTotalDict.Exists(compKey) Then
                companyTotalDict(compKey) = companyTotalDict(compKey) + amount
            Else
                companyTotalDict.Add compKey, amount
            End If
        End If
    Next i

    ' --- Append companyTotalDict to summaryDeptDict ---
    For Each k In companyTotalDict.keys
        If summaryDeptDict.Exists(k) Then
            summaryDeptDict(k) = summaryDeptDict(k) + companyTotalDict(k)
        Else
            summaryDeptDict.Add k, companyTotalDict(k)
        End If
    Next k
    Set getActualFYSummaryDict = summaryDeptDict
End Function
Function getActualFYFiveYearMappingDict() As Object
    Call declareGlobal
    Dim ws As Worksheet
    Set ws = globalActualFYSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Z").End(xlUp).row
    Dim fiveYearMappingDict As Object
    Set fiveYearMappingDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim fiveyear As String, summary As String, mapping As String, dept As String, company As String
    Dim amount As Double

    ' Build fiveYearMappingDict (by department)
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        summary = ws.Cells(i, "X").value
        mapping = fiveyear & "|" & summary
        dept = ws.Cells(i, "Y").value
        company = ws.Cells(i, "AJ").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        Dim columnFValue As String
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            Dim firstChar As String
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If mapping <> "" And mapping <> "0" Then
            Dim key As String
            key = dept & "|" & mapping
            If fiveYearMappingDict.Exists(key) Then
                fiveYearMappingDict(key) = fiveYearMappingDict(key) + amount
            Else
                fiveYearMappingDict.Add key, amount
            End If
        End If
    Next i

    ' --- Append company totals grouped by company|mapping ---
    Dim companyTotalDict As Object
    Set companyTotalDict = CreateObject("Scripting.Dictionary")
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        summary = ws.Cells(i, "X").value
        mapping = fiveyear & "|" & summary
        company = ws.Cells(i, "AJ").value
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If mapping <> "" And mapping <> "0" Then
            Dim compKey As String
            compKey = company & "|" & mapping
            If companyTotalDict.Exists(compKey) Then
                companyTotalDict(compKey) = companyTotalDict(compKey) + amount
            Else
                companyTotalDict.Add compKey, amount
            End If
        End If
    Next i

    ' Append companyTotalDict to fiveYearMappingDict
    Dim k As Variant
    For Each k In companyTotalDict.keys
        If fiveYearMappingDict.Exists(k) Then
            fiveYearMappingDict(k) = fiveYearMappingDict(k) + companyTotalDict(k)
        Else
            fiveYearMappingDict.Add k, companyTotalDict(k)
        End If
    Next k

    ' --- Add direct DEPT|FIVEYEAR keys for individual line items ---
    ' This allows access to individual line items like "Cost of Sales" without the summary prefix
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        dept = ws.Cells(i, "Y").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If fiveyear <> "" And fiveyear <> "0" Then
            Dim directKey As String
            directKey = dept & "|" & fiveyear
            If fiveYearMappingDict.Exists(directKey) Then
                fiveYearMappingDict(directKey) = fiveYearMappingDict(directKey) + amount
            Else
                fiveYearMappingDict.Add directKey, amount
            End If
        End If
    Next i
    
    ' Add direct COMPANY|FIVEYEAR keys too
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        company = ws.Cells(i, "AJ").value
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If fiveyear <> "" And fiveyear <> "0" Then
            Dim compDirectKey As String
            compDirectKey = company & "|" & fiveyear
            If fiveYearMappingDict.Exists(compDirectKey) Then
                fiveYearMappingDict(compDirectKey) = fiveYearMappingDict(compDirectKey) + amount
            Else
                fiveYearMappingDict.Add compDirectKey, amount
            End If
        End If
    Next i

    ' --- Add Total Variable Cost, Total Opex, and Total Staff Cost keys for each department ---
    Dim deptList As Object
    Set deptList = CreateObject("Scripting.Dictionary")

    ' Collect all unique departments
    For Each k In fiveYearMappingDict.keys
        dept = Trim(Split(k, "|")(0))
        If Not deptList.Exists(dept) Then deptList.Add dept, True
    Next k
    Dim varCostKeys As Variant
    varCostKeys = Array("Cost of Sales", "Gaming Taxes & Premium", "Gaming Commissions", "Complimentary", "Bad Debt Expenses")
    Dim opexKeys As Variant
    opexKeys = Array("Marketing Expenses", "Licenses and Fees", "Maintenance", "Supplies", "Uniforms/ Laundry", _
                     "Utilities", "Professional Fees", "Travel & Entertainment", "Communications", "Rent", "Other Admin")
    Dim d As Variant, sumVarCost As Double, sumOpex As Double, sumStaffCost As Double, subKey As Variant
    For Each d In deptList.keys
        sumVarCost = 0
        sumOpex = 0
        sumStaffCost = 0

        ' Sum for Total Variable Cost (Only Total Expenses in summary)
        For Each subKey In varCostKeys
            k = d & "|" & subKey
            If fiveYearMappingDict.Exists(k & "|Total Expenses") Then
                sumVarCost = sumVarCost + fiveYearMappingDict(k & "|Total Expenses")
            End If
        Next subKey
        fiveYearMappingDict(d & "|Total Variable Cost") = sumVarCost

        ' Sum for Total Opex (Only Total Expenses in summary)
        For Each subKey In opexKeys
            k = d & "|" & subKey
            If fiveYearMappingDict.Exists(k & "|Total Expenses") Then
                sumOpex = sumOpex + fiveYearMappingDict(k & "|Total Expenses")
            End If
        Next subKey
        fiveYearMappingDict(d & "|Total Opex") = sumOpex

        ' Sum for Total Staff Cost (keys that start with "Staff Costs - " and summary is "Total Expenses")
        For Each k In fiveYearMappingDict.keys
            If Left(k, Len(d & "|Staff Costs -")) = d & "|Staff Costs -" And Right(k, Len("|Total Expenses")) = "|Total Expenses" Then
                sumStaffCost = sumStaffCost + fiveYearMappingDict(k)
            End If
        Next k
        fiveYearMappingDict(d & "|Total Staff Cost") = sumStaffCost
    Next d
    Set getActualFYFiveYearMappingDict = fiveYearMappingDict
End Function



Sub populateSetupActualTotals()
    ' Usage: Call populateSetupTotals
    ' This version uses Z and AZ as the default columns.
    Call declareGlobal

    Dim startCol As String
    Dim endCol As String
    startCol = "AT"
    endCol = "BD"

    Dim wsSetup As Worksheet
    Set wsSetup = globalsetupSheet

    Dim wsActualFY As Worksheet
    Set wsActualFY = globalActualFYSheet
    
    If wsSetup.AutoFilterMode Then wsSetup.AutoFilterMode = False
    If wsActualFY.AutoFilterMode Then wsActualFY.AutoFilterMode = False
    
    Dim lastRow As Long
    lastRow = wsSetup.Cells(wsSetup.Rows.Count, "Y").End(xlUp).row

    ' Refresh Actual-FY sheet ONCE before getting data
    wsActualFY.Calculate
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False


    Dim summaryTotalDict As Object
    Set summaryTotalDict = getActualFYSummaryDict()

    Dim fiveYearMappingDict As Object
    Set fiveYearMappingDict = getActualFYFiveYearMappingDict()

    ' Build a dictionary to map header names to columns (Row 8, columns startCol to endCol)
    Dim headerDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    Dim col As Long
    For col = wsSetup.Range(startCol & "8").Column To wsSetup.Range(endCol & "8").Column
        Dim headerName As String
        headerName = Trim(wsSetup.Cells(8, col).value)
        If headerName <> "" Then
            headerDict(headerName) = col
        End If
    Next col

    ' List of keys that should NOT be overridden by fiveYearMappingDict
    Dim doNotOverrideKeys As Variant
    doNotOverrideKeys = Array( _
        "Statistics", _
        "Total Expenses", _
        "Total Non-Operating Costs", _
        "Total Revenue", _
        "Gaming Tax: Theo Tax Adjustment @2.85%", _
        "Gaming Tax: Theo Tax Adjustment @2.8%", _
        "Corporate Recharge", _
        "Branding Fee", _
        "SBC", _
        "Owner's Rental", _
        "Dividend Income" _
    )

    ' Progress bar setup
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Populating Setup Totals..."
    Progress_Bar.percentDone 0

    Dim i As Long, keyName As Variant, dictKey As String, dept As String

    For i = 9 To lastRow
        dept = wsSetup.Cells(i, "D").value

        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"

        If dept <> "" And dept <> "0" Then
            ' Plot summaryTotalDict
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                With wsSetup.Cells(i, headerDict(keyName))
                    If summaryTotalDict.Exists(dictKey) Then
                        .value = summaryTotalDict(dictKey)
                    ElseIf wsSetup.Cells(8, headerDict(keyName)).value <> "" And Not .HasFormula Then
                        .value = 0
                    ElseIf Not .HasFormula Then
                        .ClearContents
                    End If
                End With
            Next keyName
        Else
            ' Clear only if not a formula for all columns in the range
            For Each keyName In headerDict.keys
                If Not wsSetup.Cells(i, headerDict(keyName)).HasFormula Then
                    wsSetup.Cells(i, headerDict(keyName)).ClearContents
                End If
            Next keyName
        End If

        If (i - 8) Mod 10 = 0 Or i = lastRow Then
            Progress_Bar.percentDone (i - 8) / (lastRow - 8)
        End If
    Next i

    ' Plot fiveYearMappingDict after summaryTotalDict, but do not override protected keys
    For i = 9 To lastRow
        dept = wsSetup.Cells(i, "D").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"

        If dept <> "" And dept <> "0" Then
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                If fiveYearMappingDict.Exists(dictKey) Then
                    ' Only plot if not in doNotOverrideKeys
                    If IsError(Application.Match(keyName, doNotOverrideKeys, 0)) Then
                        wsSetup.Cells(i, headerDict(keyName)).value = fiveYearMappingDict(dictKey)
                    End If
                End If
            Next keyName
        End If
    Next i

    Progress_Bar.percentDone 1
    Progress_Bar.setStatusText "Done"
    Sleep 500
    Unload Progress_Bar

    ' Refresh Setup sheet ONCE after populating all
    wsSetup.Calculate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Function getBudgetFYSummaryDict() As Object
    Call declareGlobal
    Dim ws As Worksheet
    Set ws = globalBudgetFYSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Z").End(xlUp).row
    Dim summaryDeptDict As Object
    Set summaryDeptDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim summary As String, dept As String
    Dim amount As Double
    For i = 12 To lastRow
        summary = ws.Cells(i, "X").value
        dept = ws.Cells(i, "Y").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        Dim columnFValue As String
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            Dim firstChar As String
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If summary <> "" And summary <> "0" Then
            Dim key As String
            key = dept & "|" & summary
            If summaryDeptDict.Exists(key) Then
                summaryDeptDict(key) = summaryDeptDict(key) + amount
            Else
                summaryDeptDict.Add key, amount
            End If
        End If
    Next i

    ' Debug print first 10 items for BLANK DEPARTMENT only
    Dim debugCount As Long
    debugCount = 0
    Dim k As Variant
    For Each k In summaryDeptDict.keys
        If Left(k, 15) = "BLANK DEPARTMENT" Then
            Debug.Print "SummaryDict Key: " & k & " | Amount: " & summaryDeptDict(k)
            debugCount = debugCount + 1
            If debugCount >= 10 Then Exit For
        End If
    Next k

    ' Calculate totals for message box
    Dim totalRevenue As Double, totalExpenses As Double, ebitda As Double
    Dim corporateRecharge As Double, brandingFee As Double, sbc As Double
    Dim totalNonOperating As Double, netIncome As Double
    totalRevenue = 0
    totalExpenses = 0
    ebitda = 0
    corporateRecharge = 0
    brandingFee = 0
    sbc = 0
    totalNonOperating = 0
    netIncome = 0

    ' Aggregate totals (sum across all departments)
    For Each k In summaryDeptDict.keys
        If InStr(1, k, "Total Revenue", vbTextCompare) > 0 Then totalRevenue = totalRevenue + summaryDeptDict(k)
        If InStr(1, k, "Total Expenses", vbTextCompare) > 0 Then totalExpenses = totalExpenses + summaryDeptDict(k)
        If InStr(1, k, "Corporate Recharge", vbTextCompare) > 0 Then corporateRecharge = corporateRecharge + summaryDeptDict(k)
        If InStr(1, k, "Branding Fee", vbTextCompare) > 0 Then brandingFee = brandingFee + summaryDeptDict(k)
        If InStr(1, k, "SBC", vbTextCompare) > 0 Then sbc = sbc + summaryDeptDict(k)
        If InStr(1, k, "Total Non-Operating Costs", vbTextCompare) > 0 Then totalNonOperating = totalNonOperating + summaryDeptDict(k)
    Next k

    ' Revised EBITDA and Net Income calculation
    ebitda = totalRevenue + totalExpenses
    netIncome = ebitda + corporateRecharge + brandingFee + sbc + totalNonOperating

    ' Build message string
    Dim msg As String
    msg = "Total Revenue: " & Format(totalRevenue, "#,##0.00") & vbCrLf
    msg = msg & "Total Expenses: " & Format(totalExpenses, "#,##0.00") & vbCrLf
    msg = msg & "EBITDA: " & Format(ebitda, "#,##0.00") & vbCrLf & vbCrLf
    msg = msg & "Corporate Recharge: " & Format(corporateRecharge, "#,##0.00") & vbCrLf
    msg = msg & "Branding Fee: " & Format(brandingFee, "#,##0.00") & vbCrLf
    msg = msg & "SBC: " & Format(sbc, "#,##0.00") & vbCrLf
    msg = msg & "Total Non-Operating Costs: " & Format(totalNonOperating, "#,##0.00") & vbCrLf & vbCrLf
    msg = msg & "Net Income: " & Format(netIncome, "#,##0.00") & vbCrLf & vbCrLf
    'MsgBox msg, vbInformation, "Summary Totals"

    ' --- Create new dictionary getCompanyTotal, group by Company (AJ) and Summary (X) ---
    Dim companyTotalDict As Object
    Set companyTotalDict = CreateObject("Scripting.Dictionary")
    Dim company As String
    For i = 12 To lastRow
        summary = ws.Cells(i, "X").value
        company = ws.Cells(i, "AJ").value
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If summary <> "" And summary <> "0" Then
            Dim compKey As String
            compKey = company & "|" & summary
            If companyTotalDict.Exists(compKey) Then
                companyTotalDict(compKey) = companyTotalDict(compKey) + amount
            Else
                companyTotalDict.Add compKey, amount
            End If
        End If
    Next i

    ' --- Append companyTotalDict to summaryDeptDict ---
    For Each k In companyTotalDict.keys
        If summaryDeptDict.Exists(k) Then
            summaryDeptDict(k) = summaryDeptDict(k) + companyTotalDict(k)
        Else
            summaryDeptDict.Add k, companyTotalDict(k)
        End If
    Next k
    Set getBudgetFYSummaryDict = summaryDeptDict
End Function
Function getBudgetFYFiveYearMappingDict() As Object
    Call declareGlobal
    Dim ws As Worksheet
    Set ws = globalBudgetFYSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Z").End(xlUp).row
    Dim fiveYearMappingDict As Object
    Set fiveYearMappingDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim fiveyear As String, summary As String, mapping As String, dept As String, company As String
    Dim amount As Double

    ' Build fiveYearMappingDict (by department)
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        summary = ws.Cells(i, "X").value
        mapping = fiveyear & "|" & summary
        dept = ws.Cells(i, "Y").value
        company = ws.Cells(i, "AJ").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        Dim columnFValue As String
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            Dim firstChar As String
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If mapping <> "" And mapping <> "0" Then
            Dim key As String
            key = dept & "|" & mapping
            If fiveYearMappingDict.Exists(key) Then
                fiveYearMappingDict(key) = fiveYearMappingDict(key) + amount
            Else
                fiveYearMappingDict.Add key, amount
            End If
        End If
    Next i

    ' --- Append company totals grouped by company|mapping ---
    Dim companyTotalDict As Object
    Set companyTotalDict = CreateObject("Scripting.Dictionary")
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        summary = ws.Cells(i, "X").value
        mapping = fiveyear & "|" & summary
        company = ws.Cells(i, "AJ").value
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If mapping <> "" And mapping <> "0" Then
            Dim compKey As String
            compKey = company & "|" & mapping
            If companyTotalDict.Exists(compKey) Then
                companyTotalDict(compKey) = companyTotalDict(compKey) + amount
            Else
                companyTotalDict.Add compKey, amount
            End If
        End If
    Next i

    ' Append companyTotalDict to fiveYearMappingDict
    Dim k As Variant
    For Each k In companyTotalDict.keys
        If fiveYearMappingDict.Exists(k) Then
            fiveYearMappingDict(k) = fiveYearMappingDict(k) + companyTotalDict(k)
        Else
            fiveYearMappingDict.Add k, companyTotalDict(k)
        End If
    Next k


    ' --- Add direct DEPT|FIVEYEAR keys for individual line items ---
    ' This allows access to individual line items like "Cost of Sales" without the summary prefix
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        dept = ws.Cells(i, "Y").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If fiveyear <> "" And fiveyear <> "0" Then
            Dim directKey As String
            directKey = dept & "|" & fiveyear
            If fiveYearMappingDict.Exists(directKey) Then
                fiveYearMappingDict(directKey) = fiveYearMappingDict(directKey) + amount
            Else
                fiveYearMappingDict.Add directKey, amount
            End If
        End If
    Next i
    
    ' Add direct COMPANY|FIVEYEAR keys too
    For i = 12 To lastRow
        fiveyear = ws.Cells(i, "W").value
        company = ws.Cells(i, "AJ").value
        If Trim(company) = "" Or Trim(company) = "0" Then company = "BLANK COMPANY"
        amount = ws.Cells(i, "Z").value
        
        ' Invert sign if column F starts with 1-8
        columnFValue = Trim(CStr(ws.Cells(i, "F").value))
        If Len(columnFValue) > 0 Then
            firstChar = Left(columnFValue, 1)
            If firstChar >= "1" And firstChar <= "8" Then
                amount = -amount
            End If
        End If
        
        If fiveyear <> "" And fiveyear <> "0" Then
            Dim compDirectKey As String
            compDirectKey = company & "|" & fiveyear
            If fiveYearMappingDict.Exists(compDirectKey) Then
                fiveYearMappingDict(compDirectKey) = fiveYearMappingDict(compDirectKey) + amount
            Else
                fiveYearMappingDict.Add compDirectKey, amount
            End If
        End If
    Next i

    ' --- Add Total Variable Cost, Total Opex, and Total Staff Cost keys for each department ---
    Dim deptList As Object
    Set deptList = CreateObject("Scripting.Dictionary")

    ' Collect all unique departments
    For Each k In fiveYearMappingDict.keys
        dept = Trim(Split(k, "|")(0))
        If Not deptList.Exists(dept) Then deptList.Add dept, True
    Next k
    Dim varCostKeys As Variant
    varCostKeys = Array("Cost of Sales", "Gaming Taxes & Premium", "Gaming Commissions", "Complimentary", "Bad Debt Expenses")
    Dim opexKeys As Variant
    opexKeys = Array("Marketing Expenses", "Licenses and Fees", "Maintenance", "Supplies", "Uniforms/ Laundry", _
                     "Utilities", "Professional Fees", "Travel & Entertainment", "Communications", "Rent", "Other Admin")
    Dim d As Variant, sumVarCost As Double, sumOpex As Double, sumStaffCost As Double, subKey As Variant
    For Each d In deptList.keys
        sumVarCost = 0
        sumOpex = 0
        sumStaffCost = 0

        ' Sum for Total Variable Cost (Only Total Expenses in summary)
        For Each subKey In varCostKeys
            k = d & "|" & subKey
            If fiveYearMappingDict.Exists(k & "|Total Expenses") Then
                sumVarCost = sumVarCost + fiveYearMappingDict(k & "|Total Expenses")
            End If
        Next subKey
        fiveYearMappingDict(d & "|Total Variable Cost") = sumVarCost

        ' Sum for Total Opex (Only Total Expenses in summary)
        For Each subKey In opexKeys
            k = d & "|" & subKey
            If fiveYearMappingDict.Exists(k & "|Total Expenses") Then
                sumOpex = sumOpex + fiveYearMappingDict(k & "|Total Expenses")
            End If
        Next subKey
        fiveYearMappingDict(d & "|Total Opex") = sumOpex

        ' Sum for Total Staff Cost (keys that start with "Staff Costs - " and summary is "Total Expenses")
        For Each k In fiveYearMappingDict.keys
            If Left(k, Len(d & "|Staff Costs -")) = d & "|Staff Costs -" And Right(k, Len("|Total Expenses")) = "|Total Expenses" Then
                sumStaffCost = sumStaffCost + fiveYearMappingDict(k)
            End If
        Next k
        fiveYearMappingDict(d & "|Total Staff Cost") = sumStaffCost
    Next d
    Set getBudgetFYFiveYearMappingDict = fiveYearMappingDict
End Function

Sub populateSetupBudgetTotals()
    ' Usage: Call populateSetupBudgetTotals
    ' This version uses AS and T as the default columns.
    Call declareGlobal

    Dim startCol As String
    Dim endCol As String
    startCol = "Z"
    endCol = "AI"

    Dim wsSetup As Worksheet
    Set wsSetup = globalsetupSheet

    Dim wsBudgetFY As Worksheet
    Set wsBudgetFY = globalBudgetFYSheet

    If wsSetup.AutoFilterMode Then wsSetup.AutoFilterMode = False
    If wsBudgetFY.AutoFilterMode Then wsBudgetFY.AutoFilterMode = False

    Dim lastRow As Long
    lastRow = wsSetup.Cells(wsSetup.Rows.Count, "Y").End(xlUp).row

    ' Refresh Budget-FY sheet ONCE before getting data
    wsBudgetFY.Calculate
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim summaryTotalDict As Object
    Set summaryTotalDict = getBudgetFYSummaryDict()

    Dim fiveYearMappingDict As Object
    Set fiveYearMappingDict = getBudgetFYFiveYearMappingDict()

    ' Build a dictionary to map header names to columns (Row 8, columns startCol to endCol)
    Dim headerDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    Dim col As Long
    For col = wsSetup.Range(startCol & "8").Column To wsSetup.Range(endCol & "8").Column
        Dim headerName As String
        headerName = Trim(wsSetup.Cells(8, col).value)
        If headerName <> "" Then
            headerDict(headerName) = col
        End If
    Next col

    ' List of keys that should NOT be overridden by fiveYearMappingDict
    Dim doNotOverrideKeys As Variant
    doNotOverrideKeys = Array( _
        "Statistics", _
        "Total Expenses", _
        "Total Non-Operating Costs", _
        "Total Revenue", _
        "Gaming Tax: Theo Tax Adjustment @2.85%", _
        "Gaming Tax: Theo Tax Adjustment @2.8%", _
        "Corporate Recharge", _
        "Branding Fee", _
        "SBC", _
        "Owner's Rental", _
        "Dividend Income" _
    )

    ' Progress bar setup
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Populating Setup Totals..."
    Progress_Bar.percentDone 0

    Dim i As Long, keyName As Variant, dictKey As String, dept As String

    For i = 9 To lastRow
        dept = wsSetup.Cells(i, "D").value

        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"

        If dept <> "" And dept <> "0" Then
            ' Plot summaryTotalDict
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                With wsSetup.Cells(i, headerDict(keyName))
                    If summaryTotalDict.Exists(dictKey) Then
                        .value = summaryTotalDict(dictKey)
                    ElseIf wsSetup.Cells(8, headerDict(keyName)).value <> "" And Not .HasFormula Then
                        .value = 0
                    ElseIf Not .HasFormula Then
                        .ClearContents
                    End If
                End With
            Next keyName
        Else
            ' Clear only if not a formula for all columns in the range
            For Each keyName In headerDict.keys
                If Not wsSetup.Cells(i, headerDict(keyName)).HasFormula Then
                    wsSetup.Cells(i, headerDict(keyName)).ClearContents
                End If
            Next keyName
        End If

        If (i - 8) Mod 10 = 0 Or i = lastRow Then
            Progress_Bar.percentDone (i - 8) / (lastRow - 8)
        End If
    Next i

    ' Plot fiveYearMappingDict after summaryTotalDict, but do not override protected keys
    For i = 9 To lastRow
        dept = wsSetup.Cells(i, "D").value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"

        If dept <> "" And dept <> "0" Then
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                If fiveYearMappingDict.Exists(dictKey) Then
                    ' Only plot if not in doNotOverrideKeys
                    If IsError(Application.Match(keyName, doNotOverrideKeys, 0)) Then
                        wsSetup.Cells(i, headerDict(keyName)).value = fiveYearMappingDict(dictKey)
                    End If
                End If
            Next keyName
        End If
    Next i

    Progress_Bar.percentDone 1
    Progress_Bar.setStatusText "Done"
    Sleep 500
    Unload Progress_Bar

    ' Refresh Setup sheet ONCE after populating all
    wsSetup.Calculate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Sub populateWSActualTotals()

    ' Usage: Call populateWSActualTotals
    ' Uses ActiveSheet instead of globalsetupSheet
    Call declareGlobal
    Dim startCol As String
    Dim endCol As String
    Dim startRow As Long
    Dim deptCol As String
    startCol = "Y"
    endCol = "AG"
    startRow = 12 ' Header row - can be changed as needed
    deptCol = "K" ' Department column - can be changed as needed
    Dim wsSetup As Worksheet
    Set wsSetup = ActiveSheet
    Dim wsActualFY As Worksheet
    Set wsActualFY = globalActualFYSheet
    If wsSetup.AutoFilterMode Then wsSetup.AutoFilterMode = False
    If wsActualFY.AutoFilterMode Then wsActualFY.AutoFilterMode = False
    Dim lastRow As Long
    lastRow = wsSetup.Cells(wsSetup.Rows.Count, deptCol).End(xlUp).row

    ' Refresh Actual-FY sheet ONCE before getting data
    wsActualFY.Calculate
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim summaryTotalDict As Object
    Set summaryTotalDict = getActualFYSummaryDict()
    Dim fiveYearMappingDict As Object
    Set fiveYearMappingDict = getActualFYFiveYearMappingDict()

    ' Build a dictionary to map header names to columns (Row startRow, columns startCol to endCol)
    Dim headerDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    Dim col As Long
    For col = wsSetup.Range(startCol & startRow).Column To wsSetup.Range(endCol & startRow).Column
        Dim headerName As String
        headerName = Trim(wsSetup.Cells(startRow, col).value)
        If headerName <> "" Then
            headerDict(headerName) = col
        End If
    Next col

    ' List of keys that should NOT be overridden by fiveYearMappingDict
    Dim doNotOverrideKeys As Variant
    doNotOverrideKeys = Array( _
        "Statistics", _
        "Total Expenses", _
        "Total Non-Operating Costs", _
        "Total Revenue", _
        "Gaming Tax: Theo Tax Adjustment @2.85%", _
        "Gaming Tax: Theo Tax Adjustment @2.8%", _
        "Corporate Recharge", _
        "Branding Fee", _
        "SBC", _
        "Owner's Rental", _
        "Dividend Income" _
    )

    ' Progress bar setup
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Populating Worksheet Totals..."
    Progress_Bar.percentDone 0
    Dim i As Long, keyName As Variant, dictKey As String, dept As String
    For i = startRow + 1 To lastRow
        dept = wsSetup.Cells(i, deptCol).value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        If dept <> "" And dept <> "0" Then

            ' Plot summaryTotalDict
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                With wsSetup.Cells(i, headerDict(keyName))
                    If Not .HasFormula Then ' Do not override formula rows
                        If summaryTotalDict.Exists(dictKey) Then
                            .value = summaryTotalDict(dictKey)
                        ElseIf wsSetup.Cells(startRow, headerDict(keyName)).value <> "" Then
                            .value = 0
                        Else
                            .ClearContents
                        End If
                    End If
                End With
            Next keyName
        Else

            ' Clear only if not a formula for all columns in the range
            For Each keyName In headerDict.keys
                If Not wsSetup.Cells(i, headerDict(keyName)).HasFormula Then
                    wsSetup.Cells(i, headerDict(keyName)).ClearContents
                End If
            Next keyName
        End If
        If (i - startRow) Mod 10 = 0 Or i = lastRow Then
            Progress_Bar.percentDone (i - startRow) / (lastRow - startRow)
        End If
    Next i

    ' Plot fiveYearMappingDict after summaryTotalDict, but do not override protected keys
    For i = startRow + 1 To lastRow
        dept = wsSetup.Cells(i, deptCol).value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        If dept <> "" And dept <> "0" Then
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                If fiveYearMappingDict.Exists(dictKey) Then

                    ' Only plot if not in doNotOverrideKeys and not a formula
                    If IsError(Application.Match(keyName, doNotOverrideKeys, 0)) And Not wsSetup.Cells(i, headerDict(keyName)).HasFormula Then
                        wsSetup.Cells(i, headerDict(keyName)).value = fiveYearMappingDict(dictKey)
                    End If
                End If
            Next keyName
        End If
    Next i
    Progress_Bar.percentDone 1
    Progress_Bar.setStatusText "Done"
    Sleep 500
    Unload Progress_Bar

    ' Refresh active sheet ONCE after populating all
    wsSetup.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub populateWSBudgetTotals()

    ' Usage: Call populateWSBudgetTotals
    ' Uses ActiveSheet instead of globalsetupSheet
    Call declareGlobal
    Dim startCol As String
    Dim endCol As String
    Dim startRow As Long
    Dim deptCol As String
    startCol = "L"
    endCol = "T"
    startRow = 12 ' Header row - can be changed as needed
    deptCol = "K" ' Department column - can be changed as needed
    Dim wsSetup As Worksheet
    Set wsSetup = ActiveSheet
    Dim wsBudgetFY As Worksheet
    Set wsBudgetFY = globalBudgetFYSheet
    If wsSetup.AutoFilterMode Then wsSetup.AutoFilterMode = False
    If wsBudgetFY.AutoFilterMode Then wsBudgetFY.AutoFilterMode = False
    Dim lastRow As Long
    lastRow = wsSetup.Cells(wsSetup.Rows.Count, deptCol).End(xlUp).row

    ' Refresh Budget-FY sheet ONCE before getting data
    wsBudgetFY.Calculate
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim summaryTotalDict As Object
    Set summaryTotalDict = getBudgetFYSummaryDict()
    Dim fiveYearMappingDict As Object
    Set fiveYearMappingDict = getBudgetFYFiveYearMappingDict()

    ' Build a dictionary to map header names to columns (Row startRow, columns startCol to endCol)
    Dim headerDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    Dim col As Long
    For col = wsSetup.Range(startCol & startRow).Column To wsSetup.Range(endCol & startRow).Column
        Dim headerName As String
        headerName = Trim(wsSetup.Cells(startRow, col).value)
        If headerName <> "" Then
            headerDict(headerName) = col
        End If
    Next col

    ' List of keys that should NOT be overridden by fiveYearMappingDict
    Dim doNotOverrideKeys As Variant
    doNotOverrideKeys = Array( _
        "Statistics", _
        "Total Expenses", _
        "Total Non-Operating Costs", _
        "Total Revenue", _
        "Gaming Tax: Theo Tax Adjustment @2.85%", _
        "Gaming Tax: Theo Tax Adjustment @2.8%", _
        "Corporate Recharge", _
        "Branding Fee", _
        "SBC", _
        "Owner's Rental", _
        "Dividend Income" _
    )

    ' Progress bar setup
    Progress_Bar.Show False
    Progress_Bar.setStatusText "Populating Worksheet Totals..."
    Progress_Bar.percentDone 0
    Dim i As Long, keyName As Variant, dictKey As String, dept As String
    For i = startRow + 1 To lastRow
        dept = wsSetup.Cells(i, deptCol).value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        If dept <> "" And dept <> "0" Then

            ' Plot summaryTotalDict
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                With wsSetup.Cells(i, headerDict(keyName))
                    If Not .HasFormula Then ' Do not override formula rows
                        If summaryTotalDict.Exists(dictKey) Then
                            .value = summaryTotalDict(dictKey)
                        ElseIf wsSetup.Cells(startRow, headerDict(keyName)).value <> "" Then
                            .value = 0
                        Else
                            .ClearContents
                        End If
                    End If
                End With
            Next keyName
        Else

            ' Clear only if not a formula for all columns in the range
            For Each keyName In headerDict.keys
                If Not wsSetup.Cells(i, headerDict(keyName)).HasFormula Then
                    wsSetup.Cells(i, headerDict(keyName)).ClearContents
                End If
            Next keyName
        End If
        If (i - startRow) Mod 10 = 0 Or i = lastRow Then
            Progress_Bar.percentDone (i - startRow) / (lastRow - startRow)
        End If
    Next i

    ' Plot fiveYearMappingDict after summaryTotalDict, but do not override protected keys
    For i = startRow + 1 To lastRow
        dept = wsSetup.Cells(i, deptCol).value
        If Trim(dept) = "" Or Trim(dept) = "0" Then dept = "BLANK DEPARTMENT"
        If dept <> "" And dept <> "0" Then
            For Each keyName In headerDict.keys
                dictKey = dept & "|" & keyName
                If fiveYearMappingDict.Exists(dictKey) Then

                    ' Only plot if not in doNotOverrideKeys and not a formula
                    If IsError(Application.Match(keyName, doNotOverrideKeys, 0)) And Not wsSetup.Cells(i, headerDict(keyName)).HasFormula Then
                        wsSetup.Cells(i, headerDict(keyName)).value = fiveYearMappingDict(dictKey)
                    End If
                End If
            Next keyName
        End If
    Next i
    Progress_Bar.percentDone 1
    Progress_Bar.setStatusText "Done"
    Sleep 500
    Unload Progress_Bar

    ' Refresh active sheet ONCE after populating all
    wsSetup.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub UpdateKPI()

    Call populateWSBudgetTotals
    Call populateWSActualTotals

End Sub







