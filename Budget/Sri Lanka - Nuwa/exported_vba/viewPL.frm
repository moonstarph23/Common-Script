VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} viewPL
   OleObjectBlob   =   "viewPL.frx":0000
   Caption         =   "View P&L"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   85
End
Attribute VB_Name = "viewPL"
Attribute VB_Base = "0{2BCA0D66-063E-4731-9703-C7155FFA2CDF}{1A61CEEF-5AC8-407B-98F3-ACBEB2CDA80F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub updateButton_Click()
    Call declareGlobal

    If TextBox5.Text = "All" Then
        globalMasterSheet.Range("G1").value = "<>*""""*"
    Else
        globalMasterSheet.Range("G1").value = TextBox5.Text
    End If

    If TextBox2.Text = "All" Then
        globalMasterSheet.Range("G2").value = "<>*""""*"
    Else
        globalMasterSheet.Range("G2").value = TextBox2.Text
    End If
    
    If TextBox3.Text = "All" Then
        globalMasterSheet.Range("G3").value = "<>*""""*"
    Else
        globalMasterSheet.Range("G3").value = TextBox3.Text
    End If
    
    If TextBox4.Text = "All" Then
        globalMasterSheet.Range("G4").value = "<>*""""*"
    Else
        globalMasterSheet.Range("G4").value = TextBox4.Text
    End If
    
    If TextBox1.Text = "All" Then
        globalMasterSheet.Range("I1").value = "<>*""""*"
    Else
        globalMasterSheet.Range("I1").value = TextBox1.Text
    End If
    
    If TextBox6.Text = "Overall" Then
        globalMasterSheet.Range("I1").value = "Overall"
    End If
    
    If TextBox6.Text = "Summary by Department" Then
        globalMasterSheet.Range("I1").value = "Overall"
    End If
    
    If TextBox6.Text = "Summary by Company" Then
        globalMasterSheet.Range("I1").value = "Overall"
    End If
    globalMasterSheet.Calculate
    globalMasterSheetUSD.Calculate
    
    Unload Me
    
    If TextBox6.Text = "Individual P&L" Then
        globalMasterSheet.Range("A2").value = "updatePnL"
        Call updatePnL(25)
    ElseIf TextBox6.Text = "Summary" Then
        globalMasterSheet.Range("A2").value = "updatePnLAF"
        Call updatePnL(32)
    ElseIf TextBox6.Text = "Summary2" Then
        globalMasterSheet.Range("A2").value = "updatePnLAG"
        Call updatePnL(33)
    ElseIf TextBox6.Text = "Summary3" Then
        globalMasterSheet.Range("A2").value = "updatePnLAH"
        Call updatePnL(34)
    ElseIf TextBox6.Text = "Summary4" Then
        globalMasterSheet.Range("A2").value = "updatePnLAI"
        Call updatePnL(35)
    ElseIf TextBox6.Text = "SummaryByCompany" Then
        globalMasterSheet.Range("A2").value = "updatePnLAJ"
        Call updatePnL(36)
    ElseIf TextBox6.Text = "Overall" Then
        globalMasterSheet.Range("A2").value = "updatePnLALL"
        Call updatePnL(28)
    End If
    
    Range("A1").value = TextBox6.Text
End Sub
Private Sub exitButton_Click()
    Unload Me
End Sub
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub
Private Sub Treeview1_NodeClick(ByVal Node As MSComctlLib.Node)

Call declareGlobal

Dim wsDef As Worksheet

Set wsDef = ThisWorkbook.Sheets("P&L-Parameters")

Dim LastRowID As Long
LastRowID = wsDef.Cells(wsDef.Rows.Count, "A").End(xlUp).row

For Each ID In Range(wsDef.Cells(2, 1), wsDef.Cells(LastRowID, "A"))

    If ID.value = Node.Text Then
    
        TextBox5.Text = ID.Offset(, 1).value
        TextBox2.Text = ID.Offset(, 2).value 'Cost Center
        TextBox3.Text = ID.Offset(, 3).value
        TextBox4.Text = ID.Offset(, 4).value
        TextBox1.Text = ID.Offset(, 0).value
        TextBox6.Text = ID.Offset(, 5).value
    
    End If

Next ID

End Sub
Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim lastCol As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("P&L-Category1")

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
    TextBox2.Text = "All"
    TextBox3.Text = "All"
    TextBox4.Text = "All"
    TextBox5.Text = "All"
    TextBox6.Text = "Overall"
    
End Sub
Sub FillChildNodes(ByVal col As Integer, ByVal parentKey As String)
    Dim ws As Worksheet
    Set ws = Sheet1
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
            Call FillSubChildNodes(childKey, departmentValue)
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
            ' Call FillSubSubChildNodes for each sub-child
                Call FillSubSubChildNodes(childKey & "_" & subChildValue, subChildValue)
            End If
        Next rowIndex
    End If
End Sub

Sub FillSubSubChildNodes(ByVal childKey As String, ByVal departmentValue As String)
    Dim ws As Worksheet
    Set ws = Worksheets("P&L-Child-Child")
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

                ' Call FillSubSubChildNodes for each sub-child
                Call FillSubSubSubChildNodes(childKey & "_" & subChildValue, subChildValue)
            End If
        Next rowIndex
    End If
End Sub

Sub FillSubSubSubChildNodes(ByVal subSubChildKey As String, ByVal subSubChildValue As String)
    Dim ws As Worksheet
    Set ws = Worksheets("P&L-Child-Child-Child") ' Reference the "P&L-Child-Child-Child" worksheet
    Dim subSubChildColumn As Integer
    Dim lastRow As Long
    Dim found As Boolean

    ' Find the column for the sub-sub-child name in the first row
    For subSubChildColumn = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, subSubChildColumn).value = subSubChildValue Then
            found = True
            Exit For
        End If
    Next subSubChildColumn

    ' Proceed only if the sub-sub-child name was found in the first row
    If found Then
        lastRow = ws.Cells(ws.Rows.Count, subSubChildColumn).End(xlUp).row
        ' Loop from row 2 to last row in the found column to add sub-sub-sub-child nodes
        For rowIndex = 2 To lastRow
            Dim subSubSubChildValue As String
            subSubSubChildValue = ws.Cells(rowIndex, subSubChildColumn).value
            If subSubSubChildValue <> "" Then  ' Ensure the cell is not empty
                TreeView1.Nodes.Add key:=subSubChildKey & "_" & subSubSubChildValue, _
                                    Relative:=subSubChildKey, _
                                    Relationship:=tvwChild, _
                                    Text:=subSubSubChildValue
            End If
        Next rowIndex
    End If
End Sub




