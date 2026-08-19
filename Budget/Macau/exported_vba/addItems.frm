VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addItems
   OleObjectBlob   =   "addItems.frx":0000
   Caption         =   "Add New Items"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   92
End
Attribute VB_Name = "addItems"
Attribute VB_Base = "0{051A7C88-012A-48B1-87F7-FBB90C767B80}{AA0C995C-13AB-43B9-A042-1EDEE68EADED}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub clearButton_Click()

    TextBox5.Text = "" 'Company
    TextBox2.Text = "" 'Cost Center
    TextBox3.Text = "" 'Customer Segment
    TextBox4.Text = "" 'Venue
    TextBox6.Text = "" 'row number
    
    Me.ComboBox1.value = ""
    Me.ComboBox2.Clear
    Me.ComboBox3.Clear

End Sub

Private Sub updateButton_Click()
    Call declareGlobal
    
    Dim rowNumber As Long
    Dim comboBoxValue As String
    
    ' Get the value from TextBox6 and assign it as rowNumber
    rowNumber = CLng(Me.TextBox6.value)
    
    ' Populate the ActiveSheet's cell at (rowNumber, 2) with the value from ComboBox1
    If rowNumber > 0 And Len(Me.ComboBox1.value) > 0 Then
        ActiveSheet.Cells(rowNumber, 1).value = "Amount"
        ActiveSheet.Cells(rowNumber, 2).value = TextBox5.Text 'Company
        ActiveSheet.Cells(rowNumber, 3).value = TextBox2.Text 'Cost Center
        ActiveSheet.Cells(rowNumber, 4).value = TextBox3.Text 'Customer Segment
        ActiveSheet.Cells(rowNumber, 5).value = TextBox4.Text 'Venue
        ActiveSheet.Cells(rowNumber, 6).value = Left(Me.ComboBox1.value, 6) 'Ledger
        ActiveSheet.Cells(rowNumber, 7).value = Me.ComboBox2.value 'Rev Cat
        ActiveSheet.Cells(rowNumber, 8).value = Me.ComboBox3.value 'Spend Cat
        ActiveSheet.Cells(rowNumber, 11).value = "Put Description Here" 'Spend Cat
        'Debug.Print "Assigned to ActiveSheet.Cells(" & rowNumber & ", 2): " & comboBoxValue
        
        ' Construct the formula for summing columns 12 to 23
        Formula = "=SUM(L" & rowNumber & ":W" & rowNumber & ")"
        
        ' Place the formula in ActiveSheet's cell at (rowNumber, 24)
        ActiveSheet.Cells(rowNumber, 24).Formula = Formula
        
        
        ' Construct the formula for summing columns 28 to 39
        Formula2 = "=SUM(AB" & rowNumber & ":AM" & rowNumber & ")"
        
        ' Place the formula in ActiveSheet's cell at (rowNumber, 24)
        ActiveSheet.Cells(rowNumber, 40).Formula = Formula2
        
        ' Format cells in columns 12 to 24
        With ActiveSheet.Range(ActiveSheet.Cells(rowNumber, 12), ActiveSheet.Cells(rowNumber, 24))
            .NumberFormat = "#,##0;(#,##0);""-"";""-"""
            .Font.Name = "Calibri"
            .Font.Size = 11
        End With
        
        ' Format cells in columns 28 to 40
        With ActiveSheet.Range(ActiveSheet.Cells(rowNumber, 28), ActiveSheet.Cells(rowNumber, 40))
            .NumberFormat = "#,##0;(#,##0);""-"";""-"""
            .Font.Name = "Calibri"
            .Font.Size = 11
        End With
        
           ' Format cells in columns 1 to 9
        With ActiveSheet.Range(ActiveSheet.Cells(rowNumber, 1), ActiveSheet.Cells(rowNumber, 9))
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Italic = True
        End With
        
        With ActiveSheet.Range(ActiveSheet.Cells(rowNumber, 1), ActiveSheet.Cells(rowNumber, 9)).Interior
             .Color = RGB(255, 242, 204)
        End With
        
        
        ' Format cell (rowNumber, 11)
        With ActiveSheet.Cells(rowNumber, 11)
            .Font.Italic = True
            .IndentLevel = 2
            .Font.Name = "Calibri"
            .Font.Size = 11
        End With
        
        ' Format cell (rowNumber, 24)
        With ActiveSheet.Cells(rowNumber, 24).Interior
            .Color = RGB(242, 242, 242)
        End With
        
        ' Format cell (rowNumber, 40)
        With ActiveSheet.Cells(rowNumber, 40).Interior
            .Color = RGB(242, 242, 242)
        End With
        
        ' Check the conditions and assign the appropriate formula
        If ActiveSheet.Cells(rowNumber, 8).value <> "" And ActiveSheet.Cells(rowNumber, 7).value = "" Then
            Formula = "=-SUMIFS('Actual-FY'!J:J,'Actual-FY'!$B:$B,$B" & rowNumber & ",'Actual-FY'!$C:$C,$C" & rowNumber & ",'Actual-FY'!$D:$D,$D" & rowNumber & ",'Actual-FY'!$E:$E,$E" & rowNumber & ",'Actual-FY'!$F:$F,$F" & rowNumber & ",'Actual-FY'!$H:$H,$H" & rowNumber & ")"
        ElseIf ActiveSheet.Cells(rowNumber, 8).value = "" And ActiveSheet.Cells(rowNumber, 7).value <> "" Then
            Formula = "=-SUMIFS('Actual-FY'!J:J,'Actual-FY'!$B:$B,$B" & rowNumber & ",'Actual-FY'!$C:$C,$C" & rowNumber & ",'Actual-FY'!$D:$D,$D" & rowNumber & ",'Actual-FY'!$E:$E,$E" & rowNumber & ",'Actual-FY'!$F:$F,$F" & rowNumber & ",'Actual-FY'!$G:$G,$G" & rowNumber & ")"
        ElseIf ActiveSheet.Cells(rowNumber, 8).value <> "" And ActiveSheet.Cells(rowNumber, 7).value <> "" Then
            Formula = "=-SUMIFS('Actual-FY'!J:J,'Actual-FY'!$B:$B,$B" & rowNumber & ",'Actual-FY'!$C:$C,$C" & rowNumber & ",'Actual-FY'!$D:$D,$D" & rowNumber & ",'Actual-FY'!$E:$E,$E" & rowNumber & ",'Actual-FY'!$F:$F,$F" & rowNumber & ",'Actual-FY'!$G:$G,$G" & rowNumber & ",'Actual-FY'!$H:$H,$H" & rowNumber & ")"
        End If
        
        ' Place the formula in ActiveSheet's cell at (rowNumber, 28)
        If Formula <> "" Then
            ActiveSheet.Cells(rowNumber, 28).Formula = Formula
            'Debug.Print "Formula assigned to ActiveSheet.Cells(" & rowNumber & ", 28): " & Formula
            
            ' Copy the formula from cell(rowNumber, 28) to cells(rowNumber, 29) to cells(rowNumber, 39)
            ActiveSheet.Cells(rowNumber, 28).Copy Destination:=ActiveSheet.Range(ActiveSheet.Cells(rowNumber, 29), ActiveSheet.Cells(rowNumber, 39))
        End If
        
        MsgBox "Please provide USGAAP in column I"
    Else
        MsgBox "Please ensure that TextBox6 contains a valid row number and ComboBox1 has a selected value.", vbExclamation
    End If

    ActiveSheet.Calculate
    
    'Unload Me
    

End Sub

Private Sub exitButton_Click()
    Unload Me
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub Treeview1_NodeClick(ByVal Node As MSComctlLib.Node)

Call declareGlobal

Dim wsDef As Worksheet

Set wsDef = ThisWorkbook.Sheets("addItems-Parameters")

Dim LastRowID As Long
LastRowID = wsDef.Cells(wsDef.Rows.Count, "A").End(xlUp).row

For Each ID In Range(wsDef.Cells(2, 1), wsDef.Cells(LastRowID, "A"))

If ID.value = Node.Text Then

TextBox5.Text = ID.Offset(, 1).value
TextBox2.Text = ID.Offset(, 2).value 'Cost Center
TextBox3.Text = ID.Offset(, 3).value
TextBox4.Text = ID.Offset(, 4).value
'Label6.Caption = ID.Offset(, 4).Value
'TextBox1.Text = ID.Offset(, 5).Value



End If

Next ID

End Sub



Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim lastCol As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("addItems-Category1")

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
    
    'ComboBox1 -Ledger

    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim i2 As Long
    Dim value As String
    Dim uniqueValues As Collection
    
    ' Set the worksheet
    Set ws2 = ThisWorkbook.Sheets("addItems")
    
    ' Find the last row with data in column A (1st column)
    lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).row
    
    ' Initialize a collection to store unique values
    Set uniqueValues = New Collection
    
    ' Loop through each cell in column A starting from row 2
    On Error Resume Next ' Ignore errors due to duplicate values in collection
    For i2 = 2 To lastRow
        value = ws2.Cells(i2, 1).value
        If Len(value) > 0 Then
            uniqueValues.Add value, CStr(value) ' Add value to collection
        End If
    Next i2
    On Error GoTo 0 ' Resume normal error handling
    
    ' Transfer collection to array for sorting
    ReDim tempArray(1 To uniqueValues.Count)
    For i2 = 1 To uniqueValues.Count
        tempArray(i2) = uniqueValues(i2)
    Next i2
    
    ' Sort the array using a simple bubble sort algorithm
    For i = 1 To UBound(tempArray) - 1
        For j = i + 1 To UBound(tempArray)
            If tempArray(i) > tempArray(j) Then
                temp = tempArray(i)
                tempArray(i) = tempArray(j)
                tempArray(j) = temp
            End If
        Next j
    Next i
    
    ' Add the sorted unique values to ComboBox1
    For i2 = 1 To UBound(tempArray)
        Me.ComboBox1.AddItem tempArray(i2)
    Next i2

    
    
    
End Sub

Sub FillChildNodes(ByVal col As Integer, ByVal parentKey As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("addItems-Category1")
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
    Set ws = Worksheets("addItems-Child")
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

Private Sub ComboBox1_Change()
    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim i2 As Long
    Dim selectedValue As String
    Dim value As String
    
    ' Clear previous items from ComboBox2 and ComboBox3
    Me.ComboBox2.Clear
    Me.ComboBox3.Clear
    
    ' Set the worksheet
    Set ws2 = ThisWorkbook.Sheets("addItems")
    
    ' Find the last row with data in column A (1st column)
    lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).row
    
    ' Get the selected value from ComboBox1
    selectedValue = Me.ComboBox1.value
    
    ' Loop through each cell in column A starting from row 2 and filter based on ComboBox1 selection
    For i2 = 2 To lastRow
        If ws2.Cells(i2, 1).value = selectedValue Then
            ' Add value to ComboBox2 from column 2
            value = ws2.Cells(i2, 2).value
            If Len(value) > 0 Then
                Me.ComboBox2.AddItem value
                'Debug.Print "Added to ComboBox2: " & value ' Print the value added to ComboBox2
            End If
            
            ' Add value to ComboBox3 from column 3
            value = ws2.Cells(i2, 3).value
            If Len(value) > 0 Then
                Me.ComboBox3.AddItem value
                'Debug.Print "Added to ComboBox3: " & value ' Print the value added to ComboBox3
            End If
        End If
    Next i2
End Sub
