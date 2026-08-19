VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progress_Bar
   OleObjectBlob   =   "Progress_Bar.frx":0000
   Caption         =   "Please wait..."
   ClientHeight    =   840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   8
End
Attribute VB_Name = "Progress_Bar"
Attribute VB_Base = "0{2AC7468E-8C3A-4825-8F82-49BBD42B80D0}{D15445D4-A67D-404D-9984-378E71D3584F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub ProgressBar_Click()

End Sub

Private Sub ProgressContainer_Click()

End Sub

Private Sub StatusText_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Set progress bar width to 0 and update status text
    ProgressBar.Width = 0
    StatusText.Caption = "Loading..."
    
    ' Position the form in the middle of the screen
    Dim excelApp As Object
    Dim excelWindow As Object
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim formWidth As Long
    Dim formHeight As Long
    Dim leftPos As Long
    Dim topPos As Long
    
    ' Get the Excel application and active window
    Set excelApp = Application
    Set excelWindow = excelApp.ActiveWindow
    
    ' Get the screen dimensions
    screenWidth = Application.Width
    screenHeight = Application.Height
    
    ' Get the form dimensions
    formWidth = Me.Width
    formHeight = Me.Height
    
    ' Calculate the left and top positions to center the form
    leftPos = (screenWidth - formWidth) / 2
    topPos = (screenHeight - formHeight) / 2
    
    ' Set the form's position
    Me.StartUpPosition = 0 ' 0 = Manual position
    Me.Left = excelWindow.Left + leftPos
    Me.Top = excelWindow.Top + topPos
End Sub

Public Sub setStatusText(ByVal s As String)
    StatusText.Caption = s
    DoEvents
End Sub

Public Sub percentDone(ByVal percent As Double)
    If percent > 1 Then percent = percent / 100
    ProgressBar.Width = ProgressContainer.Width * percent
    DoEvents
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True
End Sub
