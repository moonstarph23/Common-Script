Attribute VB_Name = "Module3"
Dim DownTime As Date

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
    Application.DisplayAlerts = False
    ThisWorkbook.Application.Quit
    ThisWorkbook.Close savechanges = False
    
End Sub



    


