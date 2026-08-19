VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tableContents
   OleObjectBlob   =   "tableContents.frx":0000
   Caption         =   "Table of Contents"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   90
End
Attribute VB_Name = "tableContents"
Attribute VB_Base = "0{89A50FD0-0E87-4E72-8083-9B295E85F9DD}{B082EEC1-F828-4F0C-B7B6-36805A9D29E4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub Label11_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub updateButton_Click()
    Call declareGlobal


    
    
    Unload Me
    

End Sub

Private Sub exitButton_Click()
    Unload Me
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub Treeview1_NodeClick(ByVal Node As MSComctlLib.Node)

Call declareGlobal

Dim wsDef As Worksheet

Set wsDef = ThisWorkbook.Sheets("TOC-Parameters")

Dim LastRowID As Long
LastRowID = wsDef.Cells(wsDef.Rows.Count, "A").End(xlUp).row

For Each ID In Range(wsDef.Cells(2, 1), wsDef.Cells(LastRowID, "A"))

If ID.value = Node.Text Then


Label13.Caption = ID.Offset(, 2).value


    Dim ws As Worksheet
    Dim sheetFound As Boolean
    sheetFound = False
    
    ' Check if a worksheet with the name of the node exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = ID.Offset(, 1).value Then
            ws.Activate
            sheetFound = True
            Exit For
        End If
    Next ws


End If

Next ID





End Sub



Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim lastCol As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TOC")

    ' Clear all existing nodes before adding new ones
    TreeView1.Nodes.Clear

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    TreeView1.Visible = False

    For i = 1 To lastCol
        Dim parentKey As String
        parentKey = Trim(ws.Cells(1, i).value)  ' Ensure no leading/trailing spaces
        TreeView1.Nodes.Add key:=parentKey, Text:=parentKey
        Call FillChildNodes(i, parentKey)
    Next i

    TreeView1.Visible = True
    
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
    
    'Default
    'TextBox5.Text = "All"
    
    ' Expand all nodes
    For Each Node In TreeView1.Nodes
        Node.Expanded = True
    Next Node

End Sub

Sub FillChildNodes(ByVal col As Integer, ByVal parentKey As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TOC")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).row

    For rowIndex = 2 To lastRow
        Dim departmentValue As String
        departmentValue = Trim(ws.Cells(rowIndex, col).value)  ' Trim spaces
        If departmentValue <> "" Then
            Dim childKey As String
            childKey = parentKey & "_" & departmentValue
            TreeView1.Nodes.Add key:=childKey, _
                                Relative:=parentKey, _
                                Relationship:=tvwChild, _
                                Text:=departmentValue
            'Call FillSubChildNodes(childKey, departmentValue)
        End If
    Next rowIndex
End Sub

Sub FillSubChildNodes(ByVal childKey As String, ByVal departmentValue As String)
    Dim ws As Worksheet
    Set ws = Worksheets("P&L-Child")
    Dim childColumn As Integer
    Dim lastRow As Long
    Dim found As Boolean

    ' Find the column for the child name in the first row
    For childColumn = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, childColumn).value = departmentValue Then
            found = True
            Exit For
        End If
    Next childColumn

    ' Proceed only if the child name was found in the first row
    If found Then
        lastRow = ws.Cells(ws.Rows.Count, childColumn).End(xlUp).row
        ' Loop from row 2 to last row in the found column to add sub-child nodes
        For rowIndex = 2 To lastRow
            Dim subChildValue As String
            subChildValue = ws.Cells(rowIndex, childColumn).value
            If subChildValue <> "" Then  ' Ensure the cell is not empty
                TreeView1.Nodes.Add key:=childKey & "_" & subChildValue, _
                                    Relative:=childKey, _
                                    Relationship:=tvwChild, _
                                    Text:=subChildValue
            End If
        Next rowIndex
    End If
End Sub

