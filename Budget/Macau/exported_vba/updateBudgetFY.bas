Attribute VB_Name = "updateBudgetFY"
Sub updateAssumptionsExcel()
    'On Error GoTo ErrorHandler
    
        Call declareGlobal
        
        ' Turn off screen updating to prevent Excel from refreshing the screen
        Application.ScreenUpdating = False
        
        ' Turn off automatic calculations to avoid recalculating formulas during the macro
        Application.Calculation = xlCalculationManual
        
        ' Turn off event triggers to prevent other macros from running due to changes
        Application.EnableEvents = False
        
        ' Optional: Disable alerts (such as confirmation prompts)
        Application.DisplayAlerts = False
        
    
        '###STEP 2 UPDATE THE DATA OF BUDGET-FY SHEET###
    
        Call newRows(ActiveSheet)
        'Call newRowsActual(ActiveSheet)
        
        ' Turn everything back on after the macro finishes
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
         Application.DisplayStatusBar = True
        
        Exit Sub ' Exit the subroutine if no error occurs
        

        Application.Calculation = xlCalculationManual

    
ErrorHandler:
    ' Display error message
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    ' Turn everything back on after the macro finishes
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
    
End Sub
Sub updateAssumptionsForecastExcel()
    'On Error GoTo ErrorHandler
    
        Call declareGlobal
        
        ' Turn off screen updating to prevent Excel from refreshing the screen
        Application.ScreenUpdating = False
        
        ' Turn off automatic calculations to avoid recalculating formulas during the macro
        Application.Calculation = xlCalculationManual
        
        ' Turn off event triggers to prevent other macros from running due to changes
        Application.EnableEvents = False
        
        ' Optional: Disable alerts (such as confirmation prompts)
        Application.DisplayAlerts = False
        
    
        '###STEP 2 UPDATE THE DATA OF BUDGET-FY SHEET###
    
        'Call newRows(ActiveSheet)
        Call newRowsActual(ActiveSheet)
        
        ' Turn everything back on after the macro finishes
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
         Application.DisplayStatusBar = True
        
        Exit Sub ' Exit the subroutine if no error occurs
        

        Application.Calculation = xlCalculationManual

    
ErrorHandler:
    ' Display error message
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    ' Turn everything back on after the macro finishes
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
    
End Sub

Sub updateCompAssumptions()
    On Error GoTo ErrorHandler
    
    Call declareGlobal
    
     ' Turn off screen updating to prevent Excel from refreshing the screen
    Application.ScreenUpdating = False
    
    ' Turn off automatic calculations to avoid recalculating formulas during the macro
    Application.Calculation = xlCalculationManual
    
    ' Turn off event triggers to prevent other macros from running due to changes
    Application.EnableEvents = False
    
    ' Optional: Disable alerts (such as confirmation prompts)
    Application.DisplayAlerts = False
    
    Call UpdateBudgetFYWithAllocations("Complimentary", globalcompSheet)
    'Call UpdateActualFYWithAllocations("Complimentary", globalcompSheet)
    
        ' Turn everything back on after the macro finishes
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = True
        Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
        Exit Sub ' Exit the subroutine if no error occurs
    
    

    
ErrorHandler:
    ' Display error message
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    ' Turn everything back on after the macro finishes
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
End Sub
Sub updateCompForecastAssumptions()
    On Error GoTo ErrorHandler
    
    Call declareGlobal
    
     ' Turn off screen updating to prevent Excel from refreshing the screen
    Application.ScreenUpdating = False
    
    ' Turn off automatic calculations to avoid recalculating formulas during the macro
    Application.Calculation = xlCalculationManual
    
    ' Turn off event triggers to prevent other macros from running due to changes
    Application.EnableEvents = False
    
    ' Optional: Disable alerts (such as confirmation prompts)
    Application.DisplayAlerts = False
    
    'Call UpdateBudgetFYWithAllocations("Complimentary", globalcompSheet)
    Call UpdateActualFYWithAllocations("Complimentary", globalcompSheet)
    
    Call newRowsActual(ActiveSheet)
    
        ' Turn everything back on after the macro finishes
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = True
        Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
        Exit Sub ' Exit the subroutine if no error occurs
    
    

    
ErrorHandler:
    ' Display error message
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    ' Turn everything back on after the macro finishes
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
End Sub

Public Sub updatePayrollAssumptions()
    Call declareGlobal

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim Categories As Variant
    Categories = Array( _
        "Staff Costs - Allowance", "Staff Costs - Basic Salary", "Staff Costs - Benefit", _
        "Staff Costs - Bonus and Incentives", "Staff Costs - Outsourced Labour", _
        "Staff Costs - Overtime", "Staff Costs - Penalties", "Staff Costs - Relocation & Others", _
        "Staff Costs - Staff Cost Recharge", "Staff Costs - Staff Dining recharge", _
        "Staff Costs - Staff Dining", "Staff Costs - Staff Benefit recharge", _
        "Staff Costs - Employee Event & Gift", "Staff Costs - Employee Training & Development", _
        "Staff Costs - Employee Transportation", "Staff Costs - Training" _
    )

    Call UpdateBudgetFYWithAllocationsNoDept(Categories, globalPayrollSheet, "Payroll")
    'Call UpdateActualFYWithAllocationsNoDept(Categories, globalPayrollSheet, "Payroll")
    'Payroll parameter is for exclusion
    
    globalBudgetFYSheet.Calculate
    globalPayrollSheet.Calculate

    MsgBox "Budget-FY sheet has been populated successfully!", vbInformation

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
End Sub
Public Sub updatePayrollForecastAssumptions()
    Call declareGlobal

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim Categories As Variant
    Categories = Array( _
        "Staff Costs - Allowance", "Staff Costs - Basic Salary", "Staff Costs - Benefit", _
        "Staff Costs - Bonus and Incentives", "Staff Costs - Outsourced Labour", _
        "Staff Costs - Overtime", "Staff Costs - Penalties", "Staff Costs - Relocation & Others", _
        "Staff Costs - Staff Cost Recharge", "Staff Costs - Staff Dining recharge", _
        "Staff Costs - Staff Dining", "Staff Costs - Staff Benefit recharge", _
        "Staff Costs - Employee Event & Gift", "Staff Costs - Employee Training & Development", _
        "Staff Costs - Employee Transportation", "Staff Costs - Training" _
    )

    'Call UpdateBudgetFYWithAllocationsNoDept(Categories, globalPayrollSheet, "Payroll")
    Call UpdateActualFYWithAllocationsNoDept(Categories, globalPayrollSheet, "Payroll")
    'Payroll parameter is for exclusion
    Call newRowsActual(ActiveSheet)
    
    globalBudgetFYSheet.Calculate
    globalPayrollSheet.Calculate

    MsgBox "Budget-FY sheet has been populated successfully!", vbInformation

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
End Sub
Public Sub updateOpexAssumptions()
    Call declareGlobal

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim Categories As Variant
    Categories = Array( _
        "Communications", "Maintenance", "Marketing Expenses", "Other Admin", _
        "Professional Fees", "Rent", "Supplies", "Travel & Entertainment", _
        "Utilities", "Licenses and Fees", "Uniforms/ Laundry" _
    )
    

    Call UpdateBudgetFYWithAllocationsNoDept(Categories, globalOpexSheet, "Opex")
    'Call UpdateActualFYWithAllocationsNoDept(Categories, globalOpexSheet, "Opex")
    

    globalBudgetFYSheet.Calculate
    globalOpexSheet.Calculate

    MsgBox "Budget-FY sheet has been populated successfully!", vbInformation

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
End Sub
Public Sub updateOpexForecastAssumptions()
    Call declareGlobal

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim Categories As Variant
    Categories = Array( _
        "Communications", "Maintenance", "Marketing Expenses", "Other Admin", _
        "Professional Fees", "Rent", "Supplies", "Travel & Entertainment", _
        "Utilities", "Licenses and Fees", "Uniforms/ Laundry" _
    )
    

    'Call UpdateBudgetFYWithAllocationsNoDept(Categories, globalOpexSheet, "Opex")
    Call UpdateActualFYWithAllocationsNoDept(Categories, globalOpexSheet, "Opex")
    Call newRowsActual(ActiveSheet)
    
    globalBudgetFYSheet.Calculate
    globalOpexSheet.Calculate

    MsgBox "Budget-FY sheet has been populated successfully!", vbInformation

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
     Application.DisplayStatusBar = True
End Sub

Sub updateBelowAssumptions()
    Call declareGlobal
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Dim Categories As Variant
    Categories = Array( _
        "Corporate Recharge", "Branding Fee", "Table Fee", "SBC", "Inter-company Rent", "Owner's Rental", "Profit Share", "Gaming Capex Sharing", "Depreciation & Amortisation", "Pre-Opening Cost", "Interest Income/Expenses", "Interest Expenses & Financing Fees", "Development Costs", "Other Non-Operating Income/Expenses", "Tax Provision", "Employee LTI Scheme Costs" _
    )
    Call UpdateBudgetFYWithAllocationsNoDept(Categories, globalBelowEbitda, "Below Ebitda")
    globalBudgetFYSheet.Calculate
    globalBelowEbitda.Calculate
    MsgBox "Budget-FY sheet has been populated successfully!", vbInformation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
End Sub
















