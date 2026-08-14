Attribute VB_Name = "Module2"
Option Explicit

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&


'API functions
Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                                       Alias "GetWindowLongA" _
                                      (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" _
                                       Alias "SetWindowLongA" _
                                       (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" _
                                      (ByVal hwnd As Long, _
                                       ByVal hWndInsertAfter As Long, _
                                       ByVal x As Long, _
                                       ByVal Y As Long, _
                                       ByVal cx As Long, _
                                       ByVal cy As Long, _
                                       ByVal wFlags As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" _
                                    Alias "FindWindowA" _
                                    (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32.dll" _
                                         () As Long
Private Declare PtrSafe Function SendMessage Lib "user32" _
                                     Alias "SendMessageA" _
                                     (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
                                     (ByVal hwnd As Long) As Long

Public Sub AppTasklist(myForm)

'Add this userform into the Task bar
    Dim WStyle As Long
    Dim Result As Long
    Dim hwnd As Long
    
    hwnd = FindWindow(vbNullString, MainWindow.Caption)
    WStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    WStyle = WStyle Or WS_EX_APPWINDOW
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_HIDEWINDOW)
    Result = SetWindowLong(hwnd, GWL_EXSTYLE, WStyle)
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_SHOWWINDOW)

End Sub


