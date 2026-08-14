Attribute VB_Name = "Module1"
Private Const GWL_STYLE As Long = -16

Public Const MIN_BOX As Long = &H20000
Public Const MAX_BOX As Long = &H10000

Const SC_CLOSE As Long = &HF060
Const SC_MAXIMIZE As Long = &HF030
Const SC_MINIMIZE As Long = &HF020
Const SC_RESTORE As Long = &HF120

 Private Declare PtrSafe Function GetWindowLong _
   Lib "user32.dll" _
    Alias "GetWindowLongA" _
     (ByVal hwnd As Long, _
      ByVal nIndex As Long) As Long
               
 Private Declare PtrSafe Function SetWindowLong _
  Lib "user32.dll" _
   Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
     
'Redraw the Icons on the Window's Title Bar
 Private Declare PtrSafe Function DrawMenuBar _
  Lib "user32.dll" _
   (ByVal hwnd As Long) As Long

'Returns the Window Handle of the Window accepting input
 Private Declare PtrSafe Function GetForegroundWindow _
  Lib "user32.dll" () As Long

Public Sub AddToForm(ByVal Box_Type As Long)

 Dim BitMask As Long
 Dim Window_Handle As Long
 Dim WindowStyle As Long
 Dim Ret As Long

   If Box_Type = MIN_BOX Or Box_Type = MAX_BOX Then
      Window_Handle = GetForegroundWindow()
  
       WindowStyle = GetWindowLong(Window_Handle, GWL_STYLE)
       BitMask = WindowStyle Or Box_Type
  
      Ret = SetWindowLong(Window_Handle, GWL_STYLE, BitMask)
      Ret = DrawMenuBar(Window_Handle)
   End If

End Sub

Sub tst()
Application.Visible = True

End Sub
