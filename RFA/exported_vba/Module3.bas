Attribute VB_Name = "Module3"
Dim DownTime As Date

Public Sub SetTimer()
    DownTime = Now + TimeValue("05:10:00")
    Application.OnTime Earliesttime:=DownTime, Procedure:="ShutDown", Schedule:=True
    'Application.OnTime EarliestTime:=DownTime, _
    '  Procedure = "ShutDown", Schedule:=True
End Sub

Public Sub StopTimer()
    On Error Resume Next
    Application.OnTime Earliesttime:=DownTime, _
      Procedure:="ShutDown", Schedule:=False
 End Sub

Public Sub ShutDown()
    Application.DisplayAlerts = False
    'With ThisWorkbook
    '    .Saved = False
    '    .Close
    'End With
    ThisWorkbook.Application.Quit
    ThisWorkbook.Close savechanges = False
    
End Sub


'Public Sub SetTimerToRefresh()
'Dim NewDate As String
'Dim NewTime As String

'    DownTime = Now + TimeValue("00:00:01")
'    Application.OnTime Earliesttime:=DownTime
    
'    NewTime = Format(Time, "HH:MM")
'    NewTime = Replace(NewTime, ":", "")
'    If TimeNo(Time) = TimeNo("18:46:00") Then
'        If ctrString = 1 Then
'            Call StopTimerToRefresh
'            ctrString = 0
'        End If
'    ElseIf TimeNo(Time) > TimeNo("18:47:00") Then
'        ctr = 1
 '   End If
'End Sub

'Public Sub StopTimerToRefresh()
'Dim msg
'    On Error Resume Next
'    Application.OnTime Earliesttime:=DownTime , , False
'    msg = MsgBox("A")
'    Call SetTimerToRefresh
'End Sub

'Public Function TimeNo(Time As String) As Long
'TimeNo = CLng(Replace(Format(Time, "hhnnss"), ":", ""))
'End Function
