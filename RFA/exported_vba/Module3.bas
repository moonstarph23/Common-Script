Attribute VB_Name = "Module3"
Dim DownTime As Date
Public RfaWorkbookIsClosing As Boolean

Public Sub SetTimer()
    DownTime = Now + TimeValue("05:10:00")
    Application.OnTime Earliesttime:=DownTime, Procedure:="ShutDown", Schedule:=True
End Sub

Public Sub StopTimer()
    On Error Resume Next
    Application.OnTime Earliesttime:=DownTime, _
      Procedure:="ShutDown", Schedule:=False
 End Sub

Public Sub ShutDown()
    Call CloseRfaWorkbook(False)
End Sub

Public Sub CloseRfaWorkbook(Optional ByVal saveChanges As Boolean = False)
    Dim closeErrorDescription As String

    If RfaWorkbookIsClosing Then Exit Sub

    RfaWorkbookIsClosing = True
    Call StopTimer

    On Error GoTo CloseError
    ThisWorkbook.Close SaveChanges:=saveChanges
    Exit Sub

CloseError:
    closeErrorDescription = Err.Description
    RfaWorkbookIsClosing = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Visible = True
    MsgBox "The RFA workbook could not close: " & closeErrorDescription, _
        vbExclamation, "RFA Close"
End Sub


    


