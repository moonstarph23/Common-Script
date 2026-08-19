VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWait
   OleObjectBlob   =   "frmWait.frx":0000
   Caption         =   "Please wait..."
   ClientHeight    =   840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   3
End
Attribute VB_Name = "frmWait"
Attribute VB_Base = "0{D1871946-CCF1-49A3-8BBD-9557C93AA593}{7B568BFC-DC83-48F1-A58D-85E5A2070846}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub UserForm_Initialize()
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
