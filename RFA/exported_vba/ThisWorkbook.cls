VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Public ctrString As Long
Private Sub Workbook_Open()
    Dim startupErrorDescription As String

    On Error GoTo OpenError

    Application.Visible = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ThisWorkbook.Windows(1).Visible = False

    MainWindow.Show
    ctrString = 1
    Exit Sub

OpenError:
    startupErrorDescription = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Visible = True
    On Error Resume Next
    ThisWorkbook.Windows(1).Visible = True
    On Error GoTo 0
    MsgBox "The RFA application could not start: " & _
        startupErrorDescription, _
        vbCritical, "RFA Startup"
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call StopTimer
End Sub
Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    Call StopTimer
    Call SetTimer
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _
  ByVal Target As Excel.Range)
    Call StopTimer
    Call SetTimer
End Sub
