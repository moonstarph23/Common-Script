Attribute VB_Name = "progressCode"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub showProgress()
    Dim totalFiles As Long
    
    totalFiles = 39
    Progress_Bar.Show False
    For i = 1 To totalFiles
        Progress_Bar.setStatusText "Currently downloading: " & i
        Progress_Bar.percentDone i / totalFiles
        Sleep 100
    Next i
    Progress_Bar.setStatusText "Done"
    Progress_Bar.percentDone 1
    Sleep 1000
    Unload Progress_Bar
End Sub

