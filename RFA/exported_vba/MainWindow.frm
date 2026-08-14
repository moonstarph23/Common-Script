Attribute VB_Name = "MainWindow"
Attribute VB_Base = "0{D9FF0D3E-8CE4-4324-A6B7-F934406198D4}{4BB0A439-0A87-417C-80BE-41445910C4FF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim NewDatabaseLoc1 As String
Dim NewDatabaseLoc2 As String

Dim NewColumn As Long
Dim NewCycle As String
Dim NewShiftToday As String

Dim NewShiftRightNow As String
Dim VarListItem As Long
Dim mIsLoadingData As Boolean
Dim mHasValidRosterData As Boolean
'Dim NewDateRightNow As String

Private Sub cmdInsert_Click()
 Dim li As ListItem
 Dim msg
'    If Len(Trim(txtNewGroup.Text)) = 0 Then
'        msg = MsgBox("New Group must not Empty!", vbOKOnly + vbExclamation, "EISys Information!")
'        txtNewGroup.SetFocus
'        Exit Sub
'    End If
    
'    If Len(Trim(txtFullName.Text)) = 0 Then
'        msg = MsgBox("Full Name must not Empty!", vbOKOnly + vbExclamation, "EISys Information!")
'        txtFullName.SetFocus
'        Exit Sub
'    End If
   
'    Set li = lvListTrainee.ListItems.Add(, , txtIDNumber.Text)
'    li.SubItems(1) = txtFullName.Text
    
'    txtIDNumber.Text = ""
'    txtFullName.Text = ""
'    txtIDNumber.SetFocus
    'li.SubItems(2) = mcolRecords(i).Current
End Sub

Private Sub cmdInsertFile_Click()

End Sub

Private Sub cmdRemove_Click()
    Dim res As Integer
     
    If Not lvListTrainee.SelectedItem Is Nothing Then
        res = MsgBox("This item will be permanently Deleted. Continue?", vbYesNo + vbDefaultButton2, "Delete")
        
        If res = vbNo Then
            Exit Sub
        Else
            lvListTrainee.ListItems.Remove (lvListTrainee.SelectedItem.Index)
        End If
    ElseIf lvListTrainee.ListItems.Count = 0 Then
        MsgBox "Nothing to Delete...", vbExclamation, "EISys Information!"
    Else
        MsgBox "Please Highlight the Item to Delete...", vbInformation, "EISys Information!"
    End If
End Sub

Private Sub cmbDate_Change()
    If mIsLoadingData Then Exit Sub
    Call StopTimer
    Call SetTimer
    Call CallShiftStarts
End Sub

Private Sub cmbPosition_Change()
    If mIsLoadingData Then Exit Sub
    Call StopTimer
    Call SetTimer
    Call CallShiftStarts
End Sub

Private Sub cmbShift_Change()
    If mIsLoadingData Then Exit Sub
    Call StopTimer
    Call SetTimer
    Call CallShiftStarts
End Sub

Private Sub CallShiftStarts()
Dim NewColumnS As Integer
Dim NewCycleS As String
Dim iCtr As Integer
Dim NewBFound As Boolean
Dim NewStatus As String

    For i = 1 To lvDate.ListItems.Count
        If lvDate.ListItems(i).Text = cmbDate.Text Then
            NewColumnS = lvDate.ListItems(i).SubItems(1)
            NewCycleS = lvDate.ListItems(i).SubItems(2)
            i = lvDate.ListItems.Count
        End If
    Next i

    lvListofAvailableShift.ListItems.Clear
    
    If NewCycleS = "First Cycle" Then
        Call LookForRecordFirstForShifting
    ElseIf NewCycleS = "Second Cycle" Then
        Call LookForRecordSecondForShifting
    End If

    For j = 1 To lvListofAvailableShift.ListItems.Count
        lvListofAvailableShift.ListItems(j).Checked = True
    Next j
End Sub

Private Sub cmdClose_Click()
    'lvListTrainee.ListItems.Clear
    'cmbPosition.Clear
    'cmbDate.Clear
    'cmbShift.Clear
    'frmTrack.Visible = False
    'ThisWorkbook.Saved = False
    Application.DisplayAlerts = False
    ThisWorkbook.Application.Quit
    ThisWorkbook.Close savechanges:=False
    'End
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim z As Long
Dim ctr As Long
Dim NewStatus As String
Dim NewDate As Date
Dim NewRemarks As String
Dim NewShift As String

    Call StopTimer
    lvListofAbsent.ListItems.Clear
    
    For i = 1 To lvListTrainee.ListItems.Count
    
        'if lvlisttrainee.ListItems(i)
        If lvListTrainee.ListItems(i).SubItems(3) = "First Cycle" Then
            lvAll.ListItems.Clear
        
            For j = 1 To lvFirstCycle.ListItems.Count
                If InStr(1, lvFirstCycle.ListItems(j).Text, lvListTrainee.ListItems(i).Text, vbTextCompare) Then
                    lvFirstCycle.ListItems(j).Selected = True
                    lvFirstCycle.ListItems(j).EnsureVisible
                    j = lvFirstCycle.ListItems.Count
                Else
                End If
            Next j
        
            k = Val(lvListTrainee.ListItems(i).SubItems(2)) - 1
            l = 1
            NewShift = lvListTrainee.ListItems(i).SubItems(4) 'lvFirstCycle.SelectedItem.SubItems(k + 1)
            While k > 1
                NewStatus = Replace(lvFirstCycle.SelectedItem.SubItems(k), "-", "")
                
                If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                    l = l + 1
                ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                    Set li = lvAll.ListItems.Add(, , lvFirstCycle.SelectedItem.Text)
                    li.SubItems(1) = lvFirstCycle.SelectedItem.SubItems(1)
                    li.SubItems(2) = lvListTrainee.ListItems(i).SubItems(4) 'lvFirstCycle.SelectedItem.SubItems(4)
                    NewDate = DateAdd("d", -l, cmbDate.Text)
                    li.SubItems(3) = NewDate & "-(" & NewStatus & ")"
                    'li.SubItems(2) = lvListTrainee.ListItems(i).SubItems(4)
                    'k = 0
                    l = l + 1
                ElseIf IsNumeric(NewStatus) Then
                    k = 0
                End If
                k = k - 1
            Wend
                        
            z = lvAll.ListItems.Count
            ctr = 1
            While z > 0
                If ctr = 1 Then
                    NewRemarks = lvAll.ListItems(z).SubItems(3)
                    ctr = 0
                ElseIf z = 1 Then
                    NewRemarks = NewRemarks & " " & lvAll.ListItems(z).SubItems(3)
                Else
                    NewRemarks = NewRemarks & " " & lvAll.ListItems(z).SubItems(3)
                End If
                z = z - 1
            Wend
                                    
            Set li = lvListofAbsent.ListItems.Add(, , lvFirstCycle.SelectedItem.Text)
            li.SubItems(1) = lvFirstCycle.SelectedItem.SubItems(1)
            li.SubItems(2) = NewShift
            li.SubItems(3) = NewRemarks
            
        ElseIf lvListTrainee.ListItems(i).SubItems(3) = "Second Cycle" Then
            lvAll.ListItems.Clear
        
            For j = 1 To lvSecondCycle.ListItems.Count
                If InStr(1, lvSecondCycle.ListItems(j).Text, lvListTrainee.ListItems(i).Text, vbTextCompare) Then
                    lvSecondCycle.ListItems(j).Selected = True
                    lvSecondCycle.ListItems(j).EnsureVisible
                    j = lvSecondCycle.ListItems.Count
                Else
                End If
            Next j
        
            k = Val(lvListTrainee.ListItems(i).SubItems(2)) - 1
            l = 1
            NewShift = lvListTrainee.ListItems(i).SubItems(4) 'lvFirstCycle.SelectedItem.SubItems(k + 1)
            While k > 0
                If k = 1 Then
                    For a = 1 To lvFirstCycle.ListItems.Count
                        If InStr(1, lvFirstCycle.ListItems(a).Text, lvListTrainee.ListItems(i).Text, vbTextCompare) Then
                            lvFirstCycle.ListItems(a).Selected = True
                            lvFirstCycle.ListItems(a).EnsureVisible
                        Else
                        End If
                    Next a
                
                    d = 15 'Val(lvListTrainee.ListItems(i).SubItems(2)) - 1
                    b = l
                    NewShift = lvListTrainee.ListItems(i).SubItems(4) 'lvFirstCycle.SelectedItem.SubItems(k + 1)
                    
                    While d > 1
                        NewStatus = Replace(lvFirstCycle.SelectedItem.SubItems(d), "-", "")
                        
                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                            b = b + 1
                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                            Set li = lvAll.ListItems.Add(, , lvFirstCycle.SelectedItem.Text)
                            li.SubItems(1) = lvFirstCycle.SelectedItem.SubItems(1)
                            li.SubItems(2) = lvListTrainee.ListItems(i).SubItems(4) 'lvFirstCycle.SelectedItem.SubItems(4)
                            NewDate = DateAdd("d", -b, cmbDate.Text)
                            li.SubItems(3) = NewDate & "-(" & NewStatus & ")"
                            'k = 0
                            b = b + 1
                        ElseIf IsNumeric(NewStatus) Then
                            d = 0
                        End If
                        d = d - 1
                    Wend
                                
                    'c = lvAll.ListItems.Count
                    'Ctr = 1
                    'While c > 0
                    '    If Ctr = 1 Then
                    '        NewRemarks = lvAll.ListItems(c).SubItems(3)
                    '        Ctr = 0
                    '    ElseIf c = 1 Then
                    '        NewRemarks = NewRemarks & "and " & lvAll.ListItems(c).SubItems(3)
                    '    Else
                    '        NewRemarks = NewRemarks & ", " & lvAll.ListItems(c).SubItems(3)
                    '    End If
                    '    c = c - 1
                    'Wend
                                            
                    'Set li = lvListofAbsent.ListItems.Add(, , lvFirstCycle.SelectedItem.Text)
                    'li.SubItems(1) = lvFirstCycle.SelectedItem.SubItems(1)
                    'li.SubItems(2) = NewShift
                    'li.SubItems(3) = NewRemarks
                
                Else
                    NewStatus = Replace(lvSecondCycle.SelectedItem.SubItems(k), "-", "")
                    
                    If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                        l = l + 1
                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                        Set li = lvAll.ListItems.Add(, , lvSecondCycle.SelectedItem.Text)
                        li.SubItems(1) = lvSecondCycle.SelectedItem.SubItems(1)
                        li.SubItems(2) = lvListTrainee.ListItems(i).SubItems(4) 'lvSecondCycle.SelectedItem.SubItems(4)
                        NewDate = DateAdd("d", -l, cmbDate.Text)
                        li.SubItems(3) = NewDate & "-(" & NewStatus & ")"
                        'k = 0
                        l = l + 1
                    ElseIf IsNumeric(NewStatus) Then
                        k = 0
                    End If
                End If
                k = k - 1
            Wend
                        
            z = lvAll.ListItems.Count
            ctr = 1
            While z > 0
                If ctr = 1 Then
                    NewRemarks = lvAll.ListItems(z).SubItems(3)
                    ctr = 0
                ElseIf z = 1 Then
                    NewRemarks = NewRemarks & " " & lvAll.ListItems(z).SubItems(3)
                Else
                    NewRemarks = NewRemarks & " " & lvAll.ListItems(z).SubItems(3)
                End If
                z = z - 1
            Wend
                                    
            Set li = lvListofAbsent.ListItems.Add(, , lvSecondCycle.SelectedItem.Text)
            li.SubItems(1) = lvSecondCycle.SelectedItem.SubItems(1)
            li.SubItems(2) = NewShift
            li.SubItems(3) = NewRemarks
        
        End If
    Next i
    
    If lvListofAbsent.ListItems.Count > 0 Then
        Call SaveToWorksheet
    End If
    Call SetTimer
    
End Sub

Private Sub SaveToWorksheet()
On Error GoTo ErrHandler
Dim i As Long
Dim ws As Worksheet
Dim rng1 As Range
Dim NewID As String
Dim NewName As String
Dim Rand As Long

   ThisWorkbook.Windows(1).Visible = True
    Set ws = Worksheets("Sheet1")
    Set rng1 = ws.Cells(Rows.Count, "A").End(xlUp)
    
    Cells(2, 7).Value = Format(Date, "MM/DD/YYYY")
    Cells(3, 7).Value = NewShiftRightNow
    
    Rand = 7
    'MsgBox (ws.Cells(Rand, 1).Value)
    Do While ws.Cells(Rand, 1).Value <> ""
        ws.Rows(Rand) = ""
        
        'ws.Rows(Rand + 4, 1).RowHeight = 12.75
        
        With ws.Cells(Rand + 4, 1).Borders
            .LineStyle = xlNone
        End With
    
        With ws.Cells(Rand + 4, 2).Borders
            .LineStyle = xlNone
        End With
    
        With ws.Cells(Rand + 4, 3).Borders
            .LineStyle = xlNone
        End With
    
        With ws.Cells(Rand + 4, 4).Borders
            .LineStyle = xlNone
        End With
        
        With ws.Cells(Rand + 4, 5).Borders
            .LineStyle = xlNone
        End With
        
        With ws.Cells(Rand + 4, 6).Borders
            .LineStyle = xlNone
        End With
        
        With ws.Cells(Rand + 4, 7).Borders
            .LineStyle = xlNone
        End With
        
        'ws.Range("E" & Rand + 4).WrapText = True
        'ws.Range("A" & Rand + 4 & ":E" & Rand + 4).EntireRow.AutoFit
        
        Rand = Rand + 1
    
    Loop
    

    For i = 1 To lvListofAbsent.ListItems.Count
        

        NewID = Mid(lvListofAbsent.ListItems(i).Text, 2, 7)
        NewName = Mid(lvListofAbsent.ListItems(i).Text, 11, 45)
    
        'rng1.Offset(i, 0) = NewID
        'rng1.Offset(i, 1) = NewName
        'rng1.Offset(i, 2) = lvListofAbsent.ListItems(i).SubItems(1)
        'rng1.Offset(i, 3) = lvListofAbsent.ListItems(i).SubItems(3)
        ws.Cells(i + 6, 1) = NewID
        ws.Cells(i + 6, 2) = NewName
        ws.Cells(i + 6, 3) = lvListofAbsent.ListItems(i).SubItems(1)
        ws.Cells(i + 6, 4) = lvListofAbsent.ListItems(i).SubItems(2)
        
        If InStr(1, lvListofAbsent.ListItems(i).SubItems(3), "MAT", vbTextCompare) Then
            ws.Cells(i + 6, 5) = "Came from MAT"
            ws.Rows(i + 6).RowHeight = 30
        Else
            If Len(lvListofAbsent.ListItems(i).SubItems(3)) < 20 Then
                'ws.Cells(i + 4, 5) = vbNewLine & vbNewLine & lvListofAbsent.ListItems(i).SubItems(3) & vbNewLine & vbNewLine
                'ws.Cells(i + 6, 5) = vbNewLine & lvListofAbsent.ListItems(i).SubItems(3) & vbNewLine
                ws.Cells(i + 6, 5) = lvListofAbsent.ListItems(i).SubItems(3)
                
                'ws.Range("E" & Rand + 4).EntireRow.AutoFit
                'ws.Range(i + 6).RowHeight = "50"
                'ws.Range("E" & Rand + 4).EntireRow.AutoFit
                ws.Rows(i + 6).RowHeight = 30
                'ws.Rows("1:" & i + 6).EntireRow.RowHeight = 100
            'ElseIf Len(lvListofAbsent.ListItems(i).SubItems(3)) < 40 Then
            '    ws.Cells(i + 6, 5) = lvListofAbsent.ListItems(i).SubItems(3)
            'ElseIf Len(lvListofAbsent.ListItems(i).SubItems(3)) > 40 Then
            Else
                ws.Cells(i + 6, 5) = lvListofAbsent.ListItems(i).SubItems(3)
            End If
        End If
        
        With ws.Cells(i + 6, 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        With ws.Cells(i + 6, 2).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        With ws.Cells(i + 6, 3).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        With ws.Cells(i + 6, 4).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        With ws.Cells(i + 6, 5).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        With ws.Cells(i + 6, 6).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        With ws.Cells(i + 6, 7).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        'ws.Rows(i + 4).RowHeight = 45
        If Len(lvListofAbsent.ListItems(i).SubItems(3)) < 20 Then
            'ws.Range("E" & i + 4).WrapText = True
            'ws.Range("E" & i + 4).RowHeight = 40
        Else
            ws.Range("E" & i + 6).WrapText = True
            ws.Range("E" & i + 6).EntireRow.AutoFit
        End If
    
    Next i
    
   'ThisWorkbook.Application.Visible = True
    
    
    'Sheets("Sheet1").Visible = false
    
    'MainWindow.Hide
    'ws.PrintPreview
     
      Call cmSave_Click
      
    
    MainWindow.Show
    
    
    
    'ThisWorkbook.Windows(1).Visible = False
   ' ThisWorkbook.Application.Visible = False
    
    
    
    
    'Application.Visible = False
    'ThisWorkbook.Windows(1).Visible = False
    
    'Application.Windows(1).Visible = False
        'With Rng.Borders
        '.LineStyle = xlContinuous
        '.Color = vbRed
        '.Weight = xlThin
        'End With
    'ActiveWorkbook.Save
    'End If
Exit Sub
ErrHandler:
  msg = MsgBox("Cannot Proceed to Task! Call for an assistance.", vbExclamation, "ETS Guide!")
End Sub

Private Sub cmdTrackEmpSL_Click()
End Sub
Private Sub LoadDataFromExcel()
    Call LoadStartupDataWithRetry
End Sub
Private Sub ClearAll()
    lblid.Caption = ""
    lblname.Caption = ""
    lblposition.Caption = ""
    lblday1.Caption = ""
    lblday2.Caption = ""
    lblday3.Caption = ""
    lblday4.Caption = ""
    lblday5.Caption = ""
    lblday6.Caption = ""
    lblday7.Caption = ""
    lblday8.Caption = ""
    lblday9.Caption = ""
    lblday10.Caption = ""
    lblday11.Caption = ""
    lblday12.Caption = ""
    lblday13.Caption = ""
    lblday14.Caption = ""
    lblday15.Caption = ""
    lblday16.Caption = ""
    lblday17.Caption = ""
    lblday18.Caption = ""
    lblday19.Caption = ""
    lblday20.Caption = ""
    lblday21.Caption = ""
    lblday22.Caption = ""
    lblday23.Caption = ""
    lblday24.Caption = ""
    lblday25.Caption = ""
    lblday26.Caption = ""
    lblday27.Caption = ""
    lblday28.Caption = ""
   
    Call ClearShifts
    Call DefaultBackground
End Sub

Private Sub ClearShifts()
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl4.Caption = ""
    lbl5.Caption = ""
    lbl6.Caption = ""
    lbl7.Caption = ""
    lbl8.Caption = ""
    lbl9.Caption = ""
    lbl10.Caption = ""
    lbl11.Caption = ""
    lbl12.Caption = ""
    lbl13.Caption = ""
    lbl14.Caption = ""
    lbl15.Caption = ""
    lbl16.Caption = ""
    lbl17.Caption = ""
    lbl18.Caption = ""
    lbl19.Caption = ""
    lbl20.Caption = ""
    lbl21.Caption = ""
    lbl22.Caption = ""
    lbl23.Caption = ""
    lbl24.Caption = ""
    lbl25.Caption = ""
    lbl26.Caption = ""
    lbl27.Caption = ""
    lbl28.Caption = ""

End Sub


Private Sub DefaultBackground()
    frmdate1.BackColor = &H80000004
    frmdate2.BackColor = &H80000004
    frmDate3.BackColor = &H80000004
    frmDate4.BackColor = &H80000004
    frmDate5.BackColor = &H80000004
    frmDate6.BackColor = &H80000004
    frmDate7.BackColor = &H80000004
    frmDate8.BackColor = &H80000004
    frmDate9.BackColor = &H80000004
    frmDate10.BackColor = &H80000004
    frmDate11.BackColor = &H80000004
    frmDate12.BackColor = &H80000004
    frmDate13.BackColor = &H80000004
    frmDate14.BackColor = &H80000004
    frmDate15.BackColor = &H80000004
    frmDate16.BackColor = &H80000004
    frmDate17.BackColor = &H80000004
    frmDate18.BackColor = &H80000004
    frmDate19.BackColor = &H80000004
    frmDate20.BackColor = &H80000004
    frmDate21.BackColor = &H80000004
    frmDate22.BackColor = &H80000004
    frmDate23.BackColor = &H80000004
    frmDate24.BackColor = &H80000004
    frmDate25.BackColor = &H80000004
    frmDate26.BackColor = &H80000004
    frmDate27.BackColor = &H80000004
    frmDate28.BackColor = &H80000004
End Sub
Private Sub cmdRefresh_Click()
    Call LoadStartupDataWithRetry
End Sub
Private Sub cmdTrackNow_Click()
Dim i As Long
Dim msg

    If Not mHasValidRosterData Then
        MsgBox "Roster data has not been validated. Select Refresh to retry.", _
            vbExclamation, "RFA Data Loading"
        Exit Sub
    End If
'TRACK DATE
    
    Call StopTimer
    
    txtSearch.Text = ""
    Call ClearAll
    
    NewShiftRightNow = cmbShift.Text
    lvListTrainee.ListItems.Clear
    lvAll.ListItems.Clear
    lvListofAbsent.ListItems.Clear
    'cmbDate.Text = "09/18/2016"
    For i = 1 To lvDate.ListItems.Count
        If lvDate.ListItems(i).Text = cmbDate.Text Then
            NewColumn = lvDate.ListItems(i).SubItems(1)
            NewCycle = lvDate.ListItems(i).SubItems(2)
            'msg = MsgBox(NewColumn & "   " & NewCycle)
            i = lvDate.ListItems.Count
        End If
    Next i
    
    'If optAll.Value = True Then
    '    Call DisplayAllEmployeeFirstCycle
    '
    '    If lvListTrainee.ListItems.Count > 0 Then
    '        lvListTrainee.SelectedItem.Selected = True
    '        lvListTrainee.SelectedItem.EnsureVisible
    '        Call lvListTrainee_Click
    '        txtSearch.Enabled = True
    '        Call FillAll
    '        cmdPrint.Enabled = False
    '    Else
    '        cmdPrint.Enabled = False
    '        msg = MsgBox("No Record Found!", vbOKOnly + vbInformation, "ETS Confirmation!")
    '        cmdTrackNow.SetFocus
    '    End If
    'ElseIf optSL.Value = True Then
    
        If NewCycle = "First Cycle" Then
            Call FirstCycle
            
            If lvAll.ListItems.Count = 0 Then
                cmdPrint.Enabled = False
                msg = MsgBox("No Record Found!", vbOKOnly + vbInformation, "ETS Confirmation!")
                cmdTrackNow.SetFocus
            Else
                Call LookForRecordFirst
                
                If lvListTrainee.ListItems.Count > 0 Then
                    lvListTrainee.SelectedItem.Selected = True
                    lvListTrainee.SelectedItem.EnsureVisible
                    Call lvListTrainee_Click
                    txtSearch.Enabled = True
                    Call FillAll
                    cmdPrint.Enabled = True
                Else
                    cmdPrint.Enabled = False
                    msg = MsgBox("No Record Found!", vbOKOnly + vbInformation, "ETS Confirmation!")
                    cmdTrackNow.SetFocus
                End If
            End If
        ElseIf NewCycle = "Second Cycle" Then
            Call SecondCycle
            
            If lvAll.ListItems.Count = 0 Then
                cmdPrint.Enabled = False
                msg = MsgBox("No Record Found!", vbOKOnly + vbInformation, "ETS Confirmation!")
                cmdTrackNow.SetFocus
            Else
                Call LookForRecordSecond
                If lvListTrainee.ListItems.Count > 0 Then
                    lvListTrainee.SelectedItem.Selected = True
                    lvListTrainee.SelectedItem.EnsureVisible
                    Call lvListTrainee_Click
                    txtSearch.Enabled = True
                    Call FillAll
                    cmdPrint.Enabled = True
                Else
                    cmdPrint.Enabled = False
                    msg = MsgBox("No Record Found!", vbOKOnly + vbInformation, "ETS Confirmation!")
                    cmdTrackNow.SetFocus
                End If
                
            End If
        End If
    'End If
    Call SetTimer
    ThisWorkbook.Sheets("Sheet1").Activate
    Application.Visible = False
    Range("x1") = cmbDate.Value
    
End Sub

Private Sub DisplayAllEmployeeFirstCycle()
Dim i As Long
Dim j As Long
Dim x As Long
Dim NewStatus As String
Dim NewShiftFrom As Long
Dim NewShiftTo As Long
Dim msg
Dim bFound As Boolean
    
    For i = 1 To lvFirstCycle.ListItems.Count
        If lvFirstCycle.ListItems(i).Text = "-" Then
        Else
            Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
            li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
            li.SubItems(2) = NewColumn
            li.SubItems(3) = NewCycle
            li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
        End If
    Next i
    
    For i = 1 To lvSecondCycle.ListItems.Count
        For j = 1 To lvAll.ListItems.Count
            If lvAll.ListItems(j).Text = lvSecondCycle.ListItems(i).Text Then
                j = lvAll.ListItems.Count
                bFound = True
            Else
                bFound = False
                j = lvAll.ListItems.Count
            End If
        Next j
        
        If bFound = False Then
            If lvSecondCycle.ListItems(i).Text = "-" Then
            Else
                Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
                li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
                li.SubItems(2) = NewColumn
                li.SubItems(3) = NewCycle
                li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
            End If
        End If
    Next i
    
    
    If cmbShift.Text = "All Shift" Then
        NewShiftFrom = "0000"
        NewShiftTo = 2359
    ElseIf cmbShift.Text = "Morning Shift" Then
        NewShiftFrom = "0500"
        NewShiftTo = 1200
    ElseIf cmbShift.Text = "Day Shift" Then
        NewShiftFrom = 1159
        NewShiftTo = 1800
    ElseIf cmbShift.Text = "Late Day Shift" Then
        NewShiftFrom = 1759
        NewShiftTo = 1930
    ElseIf cmbShift.Text = "Night Shift" Then
        NewShiftFrom = 1929
        NewShiftTo = "2359"
    ElseIf cmbShift.Text = "Late Night Shift" Then
        NewShiftFrom = "0111"
        NewShiftTo = "0500"
    End If

    If cmbShift.Text = "All Shift" Then
        For j = 1 To 5
            If j = 1 Then
                NewShiftFrom = "0500"
                NewShiftTo = 1200
            ElseIf j = 2 Then
                NewShiftFrom = 1159
                NewShiftTo = 1800
            ElseIf j = 3 Then
                NewShiftFrom = 1759
                NewShiftTo = 1930
            ElseIf j = 4 Then
                NewShiftFrom = 1929
                NewShiftTo = "2359"
            ElseIf j = 5 Then
                NewShiftFrom = "0111"
                NewShiftTo = "0500"
            End If
            
        For i = 1 To lvAll.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
            NewStatus = Mid(lvAll.ListItems(i).SubItems(4), 1, 4)
            If IsNumeric(NewStatus) Then
                If cmbPosition.Text = "All" Then
                    If IsNumeric(NewStatus) Then
                        If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                            li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                            li.SubItems(2) = NewColumn
                            li.SubItems(3) = NewCycle
                            li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                        End If
                    End If
                Else
                    If cmbPosition.Text = "Dealer" Then
                        NewPosition = "DLR"
                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                        NewPosition = "SUP"
                    ElseIf cmbPosition.Text = "Pit Manager" Then
                        NewPosition = "PM"
                    End If
                
                    If IsNumeric(NewStatus) Then
                        If lvAll.ListItems(i).SubItems(1) = NewPosition Then
                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                                li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                                li.SubItems(2) = NewColumn
                                li.SubItems(3) = NewCycle
                                li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                            End If
                        End If
                    End If
                
                End If
            Else 'THIS IS FOR THE VL,SL,EL,EVERYTHING
                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                li.SubItems(2) = NewColumn
                li.SubItems(3) = NewCycle
                li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
            End If
        Next i
            
            
        Next j
    Else
        For i = 1 To lvAll.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
            NewStatus = Mid(lvAll.ListItems(i).SubItems(4), 1, 4)
            If IsNumeric(NewStatus) Then
                If cmbPosition.Text = "All" Then
                    If IsNumeric(NewStatus) Then
                        If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                            li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                            li.SubItems(2) = NewColumn
                            li.SubItems(3) = NewCycle
                            li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                        End If
                    End If
                Else
                    If cmbPosition.Text = "Dealer" Then
                        NewPosition = "DLR"
                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                        NewPosition = "SUP"
                    ElseIf cmbPosition.Text = "Pit Manager" Then
                        NewPosition = "PM"
                    End If
                
                    If IsNumeric(NewStatus) Then
                        If lvAll.ListItems(i).SubItems(1) = NewPosition Then
                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                                li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                                li.SubItems(2) = NewColumn
                                li.SubItems(3) = NewCycle
                                li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                            End If
                        End If
                    End If
                
                End If
            Else 'THIS IS FOR THE VL,SL,EL,EVERYTHING
                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                li.SubItems(2) = NewColumn
                li.SubItems(3) = NewCycle
                li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                
            End If
        Next i
    End If
    
    'THIS WILL INSERT ALL DATA FROM LVALL AND LVABSENT
        
End Sub

Private Sub FirstCycle()
Dim i As Long
Dim j As Long
Dim NewStatus As String
'Dim NewShiftFrom As Long
Dim NewShiftFrom As String
Dim NewShiftTo As Long
Dim msg
    
    'If cmbShift.Text = "All Shift" Then
    '    NewShiftFrom = "0559"
    '    NewShiftTo = 2359
    'ElseIf cmbShift.Text = "Morning Shift" Then
    '    NewShiftFrom = "0559"
    '    NewShiftTo = 1200
    'ElseIf cmbShift.Text = "Day Shift" Then
    '    NewShiftFrom = 1159
    '    NewShiftTo = 1800
    'ElseIf cmbShift.Text = "Night Shift" Then
    '    NewShiftFrom = 1759
    '    NewShiftTo = 2359
    'End If

    'If cmbShift.Text = "All Shift" Then
        For j = 1 To lvListofAvailableShift.ListItems.Count
        
        'For j = 1 To 3
        '    If j = 1 Then
        '        NewShiftFrom = "0559"
        '        NewShiftTo = 1200
        '    ElseIf j = 2 Then
        '        NewShiftFrom = 1159
        '        NewShiftTo = 1800
        '    ElseIf j = 3 Then
        '        NewShiftFrom = 1759
        '        NewShiftTo = 2359
        '    End If
        
            If lvListofAvailableShift.ListItems(j).Checked = True Then
                NewShiftFrom = lvListofAvailableShift.ListItems(j).Text
                NewShiftFrom = Format(NewShiftFrom, "0000")
                
                For i = 1 To lvFirstCycle.ListItems.Count
                    NewStatus = Mid(lvFirstCycle.ListItems(i).SubItems(NewColumn), 1, 4)
                    If IsNumeric(NewStatus) Then
                        If cmbPosition.Text = "All" Then
                            If IsNumeric(NewStatus) Then
                                'If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                If NewStatus = NewShiftFrom Then
                                    Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                                    li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                                    li.SubItems(2) = NewColumn
                                    li.SubItems(3) = NewCycle
                                    li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
                                End If
                            End If
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            ElseIf cmbPosition.Text = "Pit Manager" Then
                                NewPosition = "PM"
                            End If
                        
                            If IsNumeric(NewStatus) Then
                                If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                    'If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                    If NewStatus = NewShiftFrom Then
                                        Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                                        li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                                        li.SubItems(2) = NewColumn
                                        li.SubItems(3) = NewCycle
                                        li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
                                    End If
                                End If
                            End If
                        
                        End If
                    End If
                Next i
            End If
        Next j
            'NewStatus = Mid(lvFirstCycle.ListItems(i).SubItems(NewColumn), 1, 4)
            'If IsNumeric(NewStatus) Then
            '    If cmbPosition.Text = "All" Then
            '        If IsNumeric(NewStatus) Then
            '            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
            '                Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
            '                li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
            '                li.SubItems(2) = NewColumn
            '                li.SubItems(3) = NewCycle
            '                li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
            '            End If
            '        End If
            '    Else
            '        If cmbPosition.Text = "Dealer" Then
            '            NewPosition = "DLR"
            '        ElseIf cmbPosition.Text = "Pit Supervisor" Then
            '            NewPosition = "SUP"
            '        ElseIf cmbPosition.Text = "Pit Manager" Then
            '            NewPosition = "PM"
            '        End If
            '
            '        If IsNumeric(NewStatus) Then
            '            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
            '                If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
            '                    Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
            '                    li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
            '                    li.SubItems(2) = NewColumn
            '                    li.SubItems(3) = NewCycle
            '                    li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
            '                End If
            '            End If
            '        End If
            '
            '    End If
            'End If
        'Next i
            
            
        'Next j
    
    'Else
    '    For i = 1 To lvFirstCycle.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
    '        NewStatus = Mid(lvFirstCycle.ListItems(i).SubItems(NewColumn), 1, 4)
    '        If IsNumeric(NewStatus) Then
    '            If cmbPosition.Text = "All" Then
    '                If IsNumeric(NewStatus) Then
    '                    If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
    '                        Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
    '                        li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
    '                        li.SubItems(2) = NewColumn
    '                        li.SubItems(3) = NewCycle
    '                        li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
    '                    End If
    '                End If
    '            Else
    '                If cmbPosition.Text = "Dealer" Then
    '                    NewPosition = "DLR"
    '                ElseIf cmbPosition.Text = "Pit Supervisor" Then
    '                    NewPosition = "SUP"
    '                ElseIf cmbPosition.Text = "Pit Manager" Then
    '                    NewPosition = "PM"
    '                End If
    '
    '                If IsNumeric(NewStatus) Then
    '                    If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
    '                        If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
    '                            Set li = lvAll.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
    '                            li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
    '                            li.SubItems(2) = NewColumn
    '                            li.SubItems(3) = NewCycle
    '                            li.SubItems(4) = lvFirstCycle.ListItems(i).SubItems(NewColumn)
    '                        End If
    '                    End If
    '                End If
    '
    '            End If
    '        End If
    '    Next i
    'End If

End Sub

Private Sub SecondCycle()
Dim i As Long
Dim NewStatus As String
'Dim NewShiftFrom As Long
Dim NewShiftFrom As String
Dim NewShiftTo As Long
Dim msg

        For j = 1 To lvListofAvailableShift.ListItems.Count
        
            If lvListofAvailableShift.ListItems(j).Checked = True Then
                NewShiftFrom = lvListofAvailableShift.ListItems(j).Text
                'NewShiftFrom = Format(NewShiftFrom, "0000")
                
                For i = 1 To lvSecondCycle.ListItems.Count
                    NewStatus = Mid(lvSecondCycle.ListItems(i).SubItems(NewColumn), 1, 4)
                    If IsNumeric(NewStatus) = True Then
                        If cmbPosition.Text = "All" Then
                            If IsNumeric(NewStatus) Then
                                'If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                If NewStatus = NewShiftFrom Then
                                    
                                    Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
                                    li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
                                    li.SubItems(2) = NewColumn
                                    li.SubItems(3) = NewCycle
                                    li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
                                End If
                            End If
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            ElseIf cmbPosition.Text = "Pit Manager" Then
                                NewPosition = "PM"
                            End If
                        
                            If IsNumeric(NewStatus) = True Then
                                If lvSecondCycle.ListItems(i).SubItems(1) = NewPosition Then
                                    'If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                    If NewStatus = NewShiftFrom Then
                                        Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
                                        li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
                                        li.SubItems(2) = NewColumn
                                        li.SubItems(3) = NewCycle
                                        li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
                                    End If
                                End If
                            End If
                        
                        End If
                    End If
                Next i
            End If
        Next j

    'If cmbShift.Text = "All Shift" Then
    '    NewShiftFrom = "0559"
    '    NewShiftTo = 2359
    'ElseIf cmbShift.Text = "Morning Shift" Then
    '    NewShiftFrom = "0559"
    '    NewShiftTo = 1200
    'ElseIf cmbShift.Text = "Day Shift" Then
    '    NewShiftFrom = 1159
    '    NewShiftTo = 1800
    'ElseIf cmbShift.Text = "Night Shift" Then
    '    NewShiftFrom = 1759
    '    NewShiftTo = 2359
    'End If


    'If cmbShift.Text = "All Shift" Then
    '    For j = 1 To 3
    '        If j = 1 Then
    '            NewShiftFrom = "0559"
    '            NewShiftTo = 1200
    '        ElseIf j = 2 Then
    '            NewShiftFrom = 1159
    '            NewShiftTo = 1800
    '        ElseIf j = 3 Then
    '            NewShiftFrom = 1759
    '            NewShiftTo = 2359
    '        End If
            
    '        For i = 1 To lvSecondCycle.ListItems.Count
    '            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
    '            NewStatus = Mid(lvSecondCycle.ListItems(i).SubItems(NewColumn), 1, 4)
    '            If IsNumeric(NewStatus) Then
    '                If cmbPosition.Text = "All" Then
    '                    If IsNumeric(NewStatus) Then
    '                        If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
    '                            Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
    '                            li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
    '                            li.SubItems(2) = NewColumn
    '                            li.SubItems(3) = NewCycle
    '                            li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
    '                        End If
    '                    End If
    '                Else
    '                    If cmbPosition.Text = "Dealer" Then
    '                        NewPosition = "DLR"
    '                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
    '                        NewPosition = "SUP"
    '                    ElseIf cmbPosition.Text = "Pit Manager" Then
    '                        NewPosition = "PM"
    '                    End If
    '
    '                    If IsNumeric(NewStatus) Then
    '                        If lvSecondCycle.ListItems(i).SubItems(1) = NewPosition Then
    '                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
    '                                Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
    '                                li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
    '                                li.SubItems(2) = NewColumn
    '                                li.SubItems(3) = NewCycle
    '                                li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
    '                            End If
    '                        End If
    '                    End If
    '
    '                End If
    '            End If
    '        Next i
            
    '    Next j
    'Else
    
    '    For i = 1 To lvSecondCycle.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
    '        NewStatus = Mid(lvSecondCycle.ListItems(i).SubItems(NewColumn), 1, 4)
    '        If IsNumeric(NewStatus) Then
    '            If cmbPosition.Text = "All" Then
    '                If IsNumeric(NewStatus) Then
    '                    If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
    '                        Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
    '                        li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
    '                        li.SubItems(2) = NewColumn
    '                        li.SubItems(3) = NewCycle
    '                        li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
    '                    End If
    '                End If
    '            Else
    '                If cmbPosition.Text = "Dealer" Then
    '                    NewPosition = "DLR"
    '                ElseIf cmbPosition.Text = "Pit Supervisor" Then
    '                    NewPosition = "SUP"
    '                ElseIf cmbPosition.Text = "Pit Manager" Then
    '                    NewPosition = "PM"
    '                End If
                
    '                If IsNumeric(NewStatus) Then
    '                    If lvSecondCycle.ListItems(i).SubItems(1) = NewPosition Then
    '                        If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
    '                            Set li = lvAll.ListItems.Add(, , lvSecondCycle.ListItems(i).Text)
    '                            li.SubItems(1) = lvSecondCycle.ListItems(i).SubItems(1)
    '                            li.SubItems(2) = NewColumn
    '                            li.SubItems(3) = NewCycle
    '                            li.SubItems(4) = lvSecondCycle.ListItems(i).SubItems(NewColumn)
    '                        End If
    '                    End If
    '                End If
    '
    '            End If
    '        End If
    '    Next i
    'End If
End Sub

Private Sub LookForRecordFirst()
Dim i As Long
Dim j As Long
Dim k As Long
Dim x As Long
Dim SecondCtr As Long
Dim FirstCtr As Long
Dim NewStatus As String
Dim bFound As Boolean
Dim xbFound As Boolean
    lvListTrainee.ListItems.Clear
    For i = 1 To lvAll.ListItems.Count
        For j = 1 To lvFirstCycle.ListItems.Count
            If lvAll.ListItems(i).Text = lvFirstCycle.ListItems(j).Text Then
                
                SecondCtr = lvAll.ListItems(i).SubItems(2) - 1
                While SecondCtr > 1
                    NewStatus = Replace(lvFirstCycle.ListItems(j).SubItems(SecondCtr), "-", "")
                    If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "PLV" Then
                        bFound = False
                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                        If lvListTrainee.ListItems.Count = 0 Then
                            xbFound = False
                        Else
                            xbFound = False
                            For x = 1 To lvListTrainee.ListItems.Count
                                If lvAll.ListItems(i).Text = lvListTrainee.ListItems(x).Text Then
                                    xbFound = True
                                    x = lvListTrainee.ListItems.Count
                                End If
                            Next x
                        End If
                        
                        If xbFound = False Then
                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                            li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                            li.SubItems(2) = lvAll.ListItems(i).SubItems(2)
                            li.SubItems(3) = lvAll.ListItems(i).SubItems(3)
                            li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                        End If
                        
                        SecondCtr = 0
                    ElseIf IsNumeric(NewStatus) Then
                        bFound = True 'EMPLOYEE IS NOT ABSENT YESTERDAY
                        SecondCtr = 0
                    End If
                    SecondCtr = SecondCtr - 1
                Wend
                
                'THIS IS FOR FIRST CYCLE
                j = lvFirstCycle.ListItems.Count
            End If
        Next j
    Next i

End Sub

Private Sub LookForRecordSecond()
Dim i As Long
Dim j As Long
Dim k As Long
Dim SecondCtr As Long
Dim FirstCtr As Long
Dim NewStatus As String
Dim bFound As Boolean
Dim x As Long
Dim xbFound As Boolean
    lvListTrainee.ListItems.Clear
    For i = 1 To lvAll.ListItems.Count
        For j = 1 To lvSecondCycle.ListItems.Count
            If lvAll.ListItems(i).Text = lvSecondCycle.ListItems(j).Text Then
                'If lvAll.ListItems(i).SubItems(3) = "First Cycle" Then
                'ElseIf lvAll.ListItems(i).SubItems(3) = "Second Cycle" Then
                'End If
                'MsgBox (lvAll.ListItems(i).SubItems(4))
                SecondCtr = lvAll.ListItems(i).SubItems(2) - 1
                While SecondCtr > 1
                    NewStatus = Replace(lvSecondCycle.ListItems(j).SubItems(SecondCtr), "-", "")
                    If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                        bFound = False
                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                        If lvListTrainee.ListItems.Count = 0 Then
                            xbFound = False
                        Else
                            xbFound = False
                            For x = 1 To lvListTrainee.ListItems.Count
                                If lvAll.ListItems(i).Text = lvListTrainee.ListItems(x).Text Then
                                    xbFound = True
                                    x = lvListTrainee.ListItems.Count
                                End If
                            Next x
                        End If
                        
                        If xbFound = False Then
                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                            li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                            li.SubItems(2) = lvAll.ListItems(i).SubItems(2)
                            li.SubItems(3) = lvAll.ListItems(i).SubItems(3)
                            li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                        End If
                        SecondCtr = 0
                        bFound = True
                    ElseIf IsNumeric(NewStatus) Then
                        bFound = True 'EMPLOYEE IS NOT ABSENT YESTERDAY
                        SecondCtr = 0
                    End If
                    SecondCtr = SecondCtr - 1
                Wend
                
                'THIS IS FOR FIRST CYCLE
                If bFound = False Then
                    For k = 1 To lvFirstCycle.ListItems.Count
                        If lvAll.ListItems(i).Text = lvFirstCycle.ListItems(k).Text Then
                            'If lvAll.ListItems(i).SubItems(3) = "First Cycle" Then
                            'ElseIf lvAll.ListItems(i).SubItems(3) = "Second Cycle" Then
                            'End If
                            SecondCtr = 15
                            While SecondCtr > 1
                                NewStatus = Replace(lvFirstCycle.ListItems(k).SubItems(SecondCtr), "-", "")
                                If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                
                                ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                    If lvListTrainee.ListItems.Count = 0 Then
                                        xbFound = False
                                    Else
                                        xbFound = False
                                        For x = 1 To lvListTrainee.ListItems.Count
                                            If lvAll.ListItems(i).Text = lvListTrainee.ListItems(x).Text Then
                                                xbFound = True
                                                x = lvListTrainee.ListItems.Count
                                            End If
                                        Next x
                                    End If
                                    
                                    If xbFound = False Then
                                        Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).Text)
                                        li.SubItems(1) = lvAll.ListItems(i).SubItems(1)
                                        li.SubItems(2) = lvAll.ListItems(i).SubItems(2)
                                        li.SubItems(3) = lvAll.ListItems(i).SubItems(3)
                                        li.SubItems(4) = lvAll.ListItems(i).SubItems(4)
                                    End If
                                    SecondCtr = 0
                                ElseIf IsNumeric(NewStatus) Then
                                    SecondCtr = 0
                                End If
                                SecondCtr = SecondCtr - 1
                            Wend
                            k = lvFirstCycle.ListItems.Count
                        End If
                    Next k
                End If
                j = lvSecondCycle.ListItems.Count
            End If
        Next j
    Next i
End Sub

Private Sub LookForRecordFirstForShifting()
Dim NewColumnS As Integer
Dim NewCycleS As String
Dim iCtr As Integer
Dim NewBFound As Boolean
Dim NewStatus As String

    For i = 1 To lvDate.ListItems.Count
        If lvDate.ListItems(i).Text = cmbDate.Text Then
            NewColumnS = lvDate.ListItems(i).SubItems(1)
            NewCycleS = lvDate.ListItems(i).SubItems(2)
            'msg = MsgBox(NewColumn & "   " & NewCycle)
            i = lvDate.ListItems.Count
        End If
    Next i
    
    If cmbShift.Text = "All Shift" Then
        NewShiftFrom = "0000"
        NewShiftTo = 2359
    ElseIf cmbShift.Text = "Morning Shift" Then
        NewShiftFrom = "0500"
        NewShiftTo = 1200
    ElseIf cmbShift.Text = "Day Shift" Then
        NewShiftFrom = 1159
        NewShiftTo = 1800
    ElseIf cmbShift.Text = "Late Day Shift" Then
        NewShiftFrom = 1759
        NewShiftTo = 1930
    ElseIf cmbShift.Text = "Night Shift" Then
        NewShiftFrom = 1929
        NewShiftTo = "2359"
    ElseIf cmbShift.Text = "Late Night Shift" Then
        NewShiftFrom = "0111"
        NewShiftTo = "0500"
    End If

    If cmbShift.Text = "All Shift" Then
        For j = 1 To 5
            If j = 1 Then
                NewShiftFrom = "0500"
                NewShiftTo = 1200
            ElseIf j = 2 Then
                NewShiftFrom = 1159
                NewShiftTo = 1800
            ElseIf j = 3 Then
                NewShiftFrom = 1759
                NewShiftTo = 1930
            ElseIf j = 4 Then
                NewShiftFrom = 1929
                NewShiftTo = "2359"
            ElseIf j = 5 Then
                NewShiftFrom = "0111"
                NewShiftTo = "0500"
            End If
            
        For i = 1 To lvFirstCycle.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
            NewStatus = Mid(lvFirstCycle.ListItems(i).SubItems(NewColumnS), 1, 4)
            If IsNumeric(NewStatus) Then
                If cmbPosition.Text = "All" Then
                    If IsNumeric(NewStatus) Then
                        NewBFound = False
                        For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                            If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                NewBFound = True
                                iCtr = lvListofAvailableShift.ListItems.Count
                            End If
                        Next iCtr
                        
                        If NewBFound = False Then
                            Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                        End If
                        
                    End If
                Else
                    If cmbPosition.Text = "Dealer" Then
                        NewPosition = "DLR"
                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                        NewPosition = "SUP"
                    ElseIf cmbPosition.Text = "Pit Manager" Then
                        NewPosition = "PM"
                    End If
                
                    If IsNumeric(NewStatus) Then
                        If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                NewBFound = False
                                For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                                    If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                        NewBFound = True
                                        iCtr = lvListofAvailableShift.ListItems.Count
                                    End If
                                Next iCtr
                                
                                If NewBFound = False Then
                                    Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                                End If
                            End If
                        End If
                    End If
                
                End If
            End If
        Next i
            
        Next j
    Else
        For i = 1 To lvFirstCycle.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
            NewStatus = Mid(lvFirstCycle.ListItems(i).SubItems(NewColumnS), 1, 4)
            If IsNumeric(NewStatus) Then
                If cmbPosition.Text = "All" Then
                    If IsNumeric(NewStatus) Then
                        If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                            NewBFound = False
                            For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                                If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                    NewBFound = True
                                    iCtr = lvListofAvailableShift.ListItems.Count
                                End If
                            Next iCtr
                            
                            If NewBFound = False Then
                                Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                            End If
                        End If
                    End If
                Else
                    If cmbPosition.Text = "Dealer" Then
                        NewPosition = "DLR"
                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                        NewPosition = "SUP"
                    ElseIf cmbPosition.Text = "Pit Manager" Then
                        NewPosition = "PM"
                    End If
                
                    If IsNumeric(NewStatus) Then
                        If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                NewBFound = False
                                For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                                    If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                        NewBFound = True
                                        iCtr = lvListofAvailableShift.ListItems.Count
                                    End If
                                Next iCtr
                                
                                If NewBFound = False Then
                                    Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                                End If
                            End If
                        End If
                    End If
                
                End If
            End If
        Next i
    End If
    
End Sub

Private Sub LookForRecordSecondForShifting()
Dim NewColumnS As Integer
Dim NewCycleS As String
Dim iCtr As Integer
Dim NewBFound As Boolean
Dim NewStatus As String

    For i = 1 To lvDate.ListItems.Count
        If lvDate.ListItems(i).Text = cmbDate.Text Then
            NewColumnS = lvDate.ListItems(i).SubItems(1)
            NewCycleS = lvDate.ListItems(i).SubItems(2)
            'msg = MsgBox(NewColumn & "   " & NewCycle)
            i = lvDate.ListItems.Count
        End If
    Next i
    
    If cmbShift.Text = "All Shift" Then
        NewShiftFrom = "0000"
        NewShiftTo = 2359
    ElseIf cmbShift.Text = "Morning Shift" Then
        NewShiftFrom = "0500"
        NewShiftTo = 1200
    ElseIf cmbShift.Text = "Day Shift" Then
        NewShiftFrom = 1159
        NewShiftTo = 1800
    ElseIf cmbShift.Text = "Late Day Shift" Then
        NewShiftFrom = 1759
        NewShiftTo = 1930
    ElseIf cmbShift.Text = "Night Shift" Then
        NewShiftFrom = 1929
        NewShiftTo = "2359"
    ElseIf cmbShift.Text = "Late Night Shift" Then
        NewShiftFrom = "0111"
        NewShiftTo = "0500"
    End If

    If cmbShift.Text = "All Shift" Then
        For j = 1 To 5
            If j = 1 Then
                NewShiftFrom = "0500"
                NewShiftTo = 1200
            ElseIf j = 2 Then
                NewShiftFrom = 1159
                NewShiftTo = 1800
            ElseIf j = 3 Then
                NewShiftFrom = 1759
                NewShiftTo = 1930
            ElseIf j = 4 Then
                NewShiftFrom = 1929
                NewShiftTo = "2359"
            ElseIf j = 5 Then
                NewShiftFrom = "0111"
                NewShiftTo = "0500"
            End If
            
        For i = 1 To lvFirstCycle.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
            NewStatus = Mid(lvSecondCycle.ListItems(i).SubItems(NewColumnS), 1, 4)
            If IsNumeric(NewStatus) Then
                If cmbPosition.Text = "All" Then
                    If IsNumeric(NewStatus) Then
                        NewBFound = False
                        For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                            If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                NewBFound = True
                                iCtr = lvListofAvailableShift.ListItems.Count
                            End If
                        Next iCtr
                        
                        If NewBFound = False Then
                            Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                        End If
                        
                    End If
                Else
                    If cmbPosition.Text = "Dealer" Then
                        NewPosition = "DLR"
                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                        NewPosition = "SUP"
                    ElseIf cmbPosition.Text = "Pit Manager" Then
                        NewPosition = "PM"
                    End If
                
                    If IsNumeric(NewStatus) Then
                        If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                NewBFound = False
                                For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                                    If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                        NewBFound = True
                                        iCtr = lvListofAvailableShift.ListItems.Count
                                    End If
                                Next iCtr
                                
                                If NewBFound = False Then
                                    Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                                End If
                            End If
                        End If
                    End If
                
                End If
            End If
        Next i
            
        Next j
    Else
        
        For i = 1 To lvFirstCycle.ListItems.Count
            'NewStatus = Replace(lvSecondCycle.ListItems(i).SubItems(3), "-", "")
            If i > lvSecondCycle.ListItems.Count Then
            Else
                NewStatus = Mid(lvSecondCycle.ListItems(i).SubItems(NewColumnS), 1, 4)
                If IsNumeric(NewStatus) Then
                    If cmbPosition.Text = "All" Then
                        If IsNumeric(NewStatus) Then
                            If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                NewBFound = False
                                For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                                    If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                        NewBFound = True
                                        iCtr = lvListofAvailableShift.ListItems.Count
                                    End If
                                Next iCtr
                                
                                If NewBFound = False Then
                                    Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                                End If
                            End If
                        End If
                    Else
                        If cmbPosition.Text = "Dealer" Then
                            NewPosition = "DLR"
                        ElseIf cmbPosition.Text = "Pit Supervisor" Then
                            NewPosition = "SUP"
                        ElseIf cmbPosition.Text = "Pit Manager" Then
                            NewPosition = "PM"
                        End If
                    
                        If IsNumeric(NewStatus) Then
                            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                If NewStatus > NewShiftFrom And NewStatus < NewShiftTo Then
                                    NewBFound = False
                                    For iCtr = 1 To lvListofAvailableShift.ListItems.Count
                                        If lvListofAvailableShift.ListItems(iCtr).Text = NewStatus Then
                                            NewBFound = True
                                            iCtr = lvListofAvailableShift.ListItems.Count
                                        End If
                                    Next iCtr
                                    
                                    If NewBFound = False Then
                                        Set li = lvListofAvailableShift.ListItems.Add(, , NewStatus)
                                    End If
                                End If
                            End If
                        End If
                    
                    End If
                End If
            End If
        Next i
    End If
    
End Sub


Private Sub ThisIsForATest()
Dim i As Long
Dim x As Long
Dim Y As Long
Dim z As Long
Dim NewCt As Long
Dim bFound As Boolean
Dim NewTime As String
Dim NewStatus As String
Dim NewPosition As String

Dim NewNumber1 As String
Dim NewCycle1 As String
Dim NewNumber2 As String
Dim NewCycle2 As String

Dim ctr As Long
'Call ReadDataFromCloseFile

    If cmbDate.Text = "" Then
        Exit Sub
    End If
    Call ClearAll
    lvListTrainee.ListItems.Clear
        
    'lvAll.ListItems.Clear
    For x = 1 To lvFirstCycle.ListItems.Count
        lvAll.ListItems.Clear
        
        For i = 1 To lvDate.ListItems.Count
            If lvDate.ListItems(i).Text = cmbDate.Text Then
                NewNumber1 = lvDate.ListItems(i).SubItems(1)
                NewCycle1 = lvDate.ListItems(i).SubItems(2)
                If i < 15 Then
                    Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                    li.SubItems(1) = lvFirstCycle.ListItems(x).Text
                    li.SubItems(2) = lvFirstCycle.ListItems(x).SubItems(1)
                    li.SubItems(3) = lvFirstCycle.ListItems(x).SubItems(i + 1)
                    li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                Else
                    Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                    li.SubItems(1) = lvSecondCycle.ListItems(x).Text
                    li.SubItems(2) = lvSecondCycle.ListItems(x).SubItems(1)
                    li.SubItems(3) = lvSecondCycle.ListItems(x).SubItems(i - 13)
                    li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                End If
                i = lvDate.ListItems.Count + 1
            Else
                If i < 15 Then
                    Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                    li.SubItems(1) = lvFirstCycle.ListItems(x).Text
                    li.SubItems(2) = lvFirstCycle.ListItems(x).SubItems(1)
                    li.SubItems(3) = lvFirstCycle.ListItems(x).SubItems(i + 1)
                    li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                Else
                    For Y = 1 To lvSecondCycle.ListItems.Count
                        If lvFirstCycle.ListItems(x).Text = lvSecondCycle.ListItems(Y).Text Then
                            Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                            li.SubItems(1) = lvSecondCycle.ListItems(x).Text
                            li.SubItems(2) = lvSecondCycle.ListItems(x).SubItems(1)
                            li.SubItems(3) = lvSecondCycle.ListItems(x).SubItems(i - 13)
                            li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                            Y = lvSecondCycle.ListItems.Count
                        End If
                    Next Y
                End If
            End If
            
            
            
        Next i
        
        i = lvAll.ListItems.Count
            NewShiftHour = Replace(lvAll.ListItems(i).SubItems(3), "-", "")
            
            If Len(Trim(NewShiftHour)) = "" Then
            Else
                NewTime = Mid(lvAll.ListItems(i).SubItems(3), 1, 4)
                If IsNumeric(NewTime) = True Then
                    If cmbShift.Text = "Morning Shift" Then
                        If Int(NewTime) > "0559" And Int(NewTime) < 1200 Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            li.SubItems(2) = NewNumber1
                                            li.SubItems(3) = NewCycle1
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                                li.SubItems(2) = NewNumber1
                                                li.SubItems(3) = NewCycle1
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Day Shift" Then
                        If Int(NewTime) < 1800 Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            li.SubItems(2) = NewNumber1
                                            li.SubItems(3) = NewCycle1
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                                li.SubItems(2) = NewNumber1
                                                li.SubItems(3) = NewCycle1
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Late Day Shift" Then
                        If Int(NewTime) < 1930 Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            li.SubItems(2) = NewNumber1
                                            li.SubItems(3) = NewCycle1
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                                li.SubItems(2) = NewNumber1
                                                li.SubItems(3) = NewCycle1
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Night Shift" Then
                        If Int(NewTime) < "2359" Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            li.SubItems(2) = NewNumber1
                                            li.SubItems(3) = NewCycle1
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                                li.SubItems(2) = NewNumber1
                                                li.SubItems(3) = NewCycle1
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Late Night Shift" Then
                        If Int(NewTime) > "0111" And Int(NewTime) < "0500" Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            li.SubItems(2) = NewNumber1
                                            li.SubItems(3) = NewCycle1
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                                li.SubItems(2) = NewNumber1
                                                li.SubItems(3) = NewCycle1
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    
                    End If
                Else
                End If
            End If
            
            
        
        
        
    Next x
    
    NewCt = 0
    For x = 165 To 185 'lvSecondCycle.ListItems.Count
        If lvSecondCycle.ListItems(x).Text = "[P006131] CALENDAS, Normita" Then
            NewCt = x
        End If
        lvAll.ListItems.Clear
        For z = 1 To lvFirstCycle.ListItems.Count
            If InStr(1, lvFirstCycle.ListItems(z).Text, lvSecondCycle.ListItems(x).Text, vbTextCompare) Then
                bFound = True
                z = lvFirstCycle.ListItems.Count
            Else
                bFound = False
            End If
        Next z
        
        If bFound = False Then
            For i = 1 To lvDate.ListItems.Count
                If lvDate.ListItems(i).Text = cmbDate.Text Then
                    If i < 15 Then
                        
                        Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                        li.SubItems(1) = lvSecondCycle.ListItems(x).Text
                        li.SubItems(2) = lvSecondCycle.ListItems(x).SubItems(1)
                        li.SubItems(3) = "-"
                        li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                    Else
                        Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                        li.SubItems(1) = lvSecondCycle.ListItems(x).Text
                        li.SubItems(2) = lvSecondCycle.ListItems(x).SubItems(1)
                        li.SubItems(3) = lvSecondCycle.ListItems(x).SubItems(i - 13)
                        li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                    End If
                    i = lvDate.ListItems.Count + 1
                Else
                    If i < 15 Then
                        Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                        li.SubItems(1) = lvSecondCycle.ListItems(x).Text
                        li.SubItems(2) = lvSecondCycle.ListItems(x).SubItems(1)
                        li.SubItems(3) = "-"
                        li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                    Else
                        For Y = 1 To lvSecondCycle.ListItems.Count
                            'If lvFirstCycle.ListItems(x).Text = lvSecondCycle.ListItems(y).Text Then
                                Set li = lvAll.ListItems.Add(, , lvDate.ListItems(i).Text)
                                li.SubItems(1) = lvSecondCycle.ListItems(x).Text
                                li.SubItems(2) = lvSecondCycle.ListItems(x).SubItems(1)
                                li.SubItems(3) = lvSecondCycle.ListItems(x).SubItems(i - 13)
                                li.SubItems(4) = lvDate.ListItems(i).SubItems(2)
                                Y = lvSecondCycle.ListItems.Count
                            'End If
                        Next Y
                    End If
                End If
            Next i
            
        i = lvAll.ListItems.Count
            NewShiftHour = Replace(lvAll.ListItems(i).SubItems(3), "-", "")
            
            If Len(Trim(NewShiftHour)) = "" Then
            Else
                NewTime = Mid(lvAll.ListItems(i).SubItems(3), 1, 4)
                If IsNumeric(NewTime) = True Then
                    If cmbShift.Text = "Morning Shift" Then
                        If Int(NewTime) > "0559" And Int(NewTime) < 1200 Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Day Shift" Then
                        If Int(NewTime) < 1800 Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Late Day Shift" Then
                        If Int(NewTime) < 1930 Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Night Shift" Then
                        If Int(NewTime) < "2359" Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    ElseIf cmbShift.Text = "Late Night Shift" Then
                        If Int(NewTime) > "0111" And Int(NewTime) < "0500" Then
                            While i > 0
                            
                                If cmbPosition.Text = "All" Then
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                            li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                Else
                                    If cmbPosition.Text = "Dealer" Then
                                        NewPosition = "DLR"
                                    ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                        NewPosition = "SUP"
                                    ElseIf cmbPosition.Text = "Pit Manager" Then
                                        NewPosition = "PM"
                                    End If
                                    
                                    ctr = i - 1
                                    If ctr > 1 Then
                                        NewStatus = Replace(lvAll.ListItems(ctr).SubItems(3), "-", "")
                                        If UCase(NewStatus) = "REST" Or UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Or UCase(NewStatus) = "LVCOV100" Or UCase(NewStatus) = "LVCOV50" Or UCase(NewStatus) = "PLV" Then
                                            
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "MAT" Or UCase(NewStatus) = "PMLV" Or UCase(NewStatus) = "PCL" Then
                                            If NewPosition = lvAll.ListItems(i).SubItems(2) Then
                                                Set li = lvListTrainee.ListItems.Add(, , lvAll.ListItems(i).SubItems(1))
                                                li.SubItems(1) = lvAll.ListItems(i).SubItems(2)
                                            End If
                                            i = 0
                                        Else
                                            i = 0
                                        End If
                                    End If
                                
                                
                                End If
                                
                                i = i - 1
                            Wend
                            
                        End If
                    
                    End If
                Else
                End If
            End If
            
        End If
        
        
        
        
        
        
    Next x


    If lvListTrainee.ListItems.Count > 0 Then
        lvListTrainee.SelectedItem.Selected = True
        lvListTrainee.SelectedItem.EnsureVisible
        Call lvListTrainee_Click
        txtSearch.Enabled = True
    Else
        msg = MsgBox("No Record Found!", vbOKOnly + vbInformation, "ETS System")
    End If
    'THIS IS FOR SORTING THE DATE

    
        '.ColumnHeaders.Add , , "Date", 100
        '.ColumnHeaders.Add , , "Name", 100
        '.ColumnHeaders.Add , , "Name", 100
        '.ColumnHeaders.Add , , "Shift", 100
        '.ColumnHeaders.Add , , "Cycle", 150

    'For i = 1 To lvFirstCycle.ListItems.cout
        
    'Next i
    'msg = MsgBox(NewCt)
    Call FillAll

End Sub

Private Sub FillAll()
    Dim NewDate As Date
    
    NewDate = lvDate.ListItems(1).Text
    
    lblday1.Caption = NewDate
    lblday2.Caption = DateAdd("d", 1, NewDate)
    lblday3.Caption = DateAdd("d", 2, NewDate)
    lblday4.Caption = DateAdd("d", 3, NewDate)
    lblday5.Caption = DateAdd("d", 4, NewDate)
    lblday6.Caption = DateAdd("d", 5, NewDate)
    lblday7.Caption = DateAdd("d", 6, NewDate)
    lblday8.Caption = DateAdd("d", 7, NewDate)
    lblday9.Caption = DateAdd("d", 8, NewDate)
    lblday10.Caption = DateAdd("d", 9, NewDate)
    lblday11.Caption = DateAdd("d", 10, NewDate)
    lblday12.Caption = DateAdd("d", 11, NewDate)
    lblday13.Caption = DateAdd("d", 12, NewDate)
    lblday14.Caption = DateAdd("d", 13, NewDate)
    lblday15.Caption = DateAdd("d", 14, NewDate)
    lblday16.Caption = DateAdd("d", 15, NewDate)
    lblday17.Caption = DateAdd("d", 16, NewDate)
    lblday18.Caption = DateAdd("d", 17, NewDate)
    lblday19.Caption = DateAdd("d", 18, NewDate)
    lblday20.Caption = DateAdd("d", 19, NewDate)
    lblday21.Caption = DateAdd("d", 20, NewDate)
    lblday22.Caption = DateAdd("d", 21, NewDate)
    lblday23.Caption = DateAdd("d", 22, NewDate)
    lblday24.Caption = DateAdd("d", 23, NewDate)
    lblday25.Caption = DateAdd("d", 24, NewDate)
    lblday26.Caption = DateAdd("d", 25, NewDate)
    lblday27.Caption = DateAdd("d", 26, NewDate)
    lblday28.Caption = DateAdd("d", 27, NewDate)
    
    'lblday1.Caption = lvDate.ListItems(1).Text
End Sub

Private Sub Sample()
Dim i As Long
Dim x As Long
Dim NewCtr As Long
Dim NewDate As Date
Dim NewSubItem As Integer
Dim NewUsedListView As String
Dim NewShiftHour  As String
Dim NewTime As String
Dim NewPosition As String
Dim NewInsert As Long
Dim NewStatus As String
Dim msg

    lvListTrainee.ListItems.Clear
    
    For i = 1 To lvDate.ListItems.Count
        If cmbDate.Text = lvDate.ListItems(i).Text Then
            NewDate = lvDate.ListItems(i).Text
            NewSubItem = lvDate.ListItems(i).SubItems(1)
            NewUsedListView = lvDate.ListItems(i).SubItems(2)
         i = lvDate.ListItems.Count
        End If
    Next i


'msg = MsgBox(NewDate & " " & NewSubItem & " " & NewUsedListView)

If NewUsedListView = "First Cycle" Then
    NewInsert = 0
    
    If NewSubItem = 2 Then
        msg = MsgBox("Cannot Proceed to Task! Previous Cycle doesn't Exist anymore.", vbOKOnly + vbExclamation, "ETS Guide!")
        Exit Sub
    End If
    
    For i = 1 To lvFirstCycle.ListItems.Count
        NewShiftHour = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem), "-", "")
        If Len(Trim(NewShiftHour)) = "" Then
        Else
            NewTime = Mid(NewShiftHour, 1, 4)
            If IsNumeric(NewShiftHour) = True Then
                If cmbShift.Text = "Morning Shift" Then
                    If Int(NewTime) > "0559" And Int(NewTime) < 1200 Then
                        If cmbPosition.Text = "All" Then
                            'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                            'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                            NewCtr = NewSubItem
                            For x = 1 To 14
                                If NewCtr > 1 Then
                                    NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                    If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                        NewInsert = 1
                                        x = 14
                                    ElseIf IsNumeric(NewStatus) = True Then
                                        x = 14
                                    ElseIf Len(Trim(NewStatus)) = "" Then
                                        x = 14
                                    End If
                                    NewCtr = NewSubItem - x
                                End If
                            Next x
                            
                            'NewInsert = 1
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            End If
                            
                            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                                'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                                
                                NewCtr = NewSubItem
                                For x = 1 To 14
                                    If NewCtr > 1 Then
                                        NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                        If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                            NewInsert = 1
                                            x = 14
                                        ElseIf IsNumeric(NewStatus) = True Then
                                            x = 14
                                        ElseIf Len(Trim(NewStatus)) = "" Then
                                            x = 14
                                        End If
                                        NewCtr = NewSubItem - x
                                    End If
                                Next x
                            End If
                        End If
                    End If
                ElseIf cmbShift.Text = "Day Shift" Then
                    NewTime = Mid(NewShiftHour, 1, 4)
                    If NewTime < 1800 Then
                        If cmbPosition.Text = "All" Then
                            'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                            'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                            NewCtr = NewSubItem
                            For x = 1 To 14
                                If NewCtr > 1 Then
                                    NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                    If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                        NewInsert = 1
                                        x = 14
                                    ElseIf IsNumeric(NewStatus) = True Then
                                        x = 14
                                    ElseIf Len(Trim(NewStatus)) = "" Then
                                        x = 14
                                    End If
                                    NewCtr = NewSubItem - x
                                End If
                            Next x
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            End If
                            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                                'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                                NewCtr = NewSubItem
                                For x = 1 To 14
                                    If NewCtr > 1 Then
                                        NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                        If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                            NewInsert = 1
                                            x = 14
                                        ElseIf IsNumeric(NewStatus) = True Then
                                            x = 14
                                        ElseIf Len(Trim(NewStatus)) = "" Then
                                            x = 14
                                        End If
                                        NewCtr = NewSubItem - x
                                    End If
                                Next x
                            
                            End If
                        End If
                    End If
                ElseIf cmbShift.Text = "Late Day Shift" Then
                    NewTime = Mid(NewShiftHour, 1, 4)
                    If NewTime < 1930 Then
                        If cmbPosition.Text = "All" Then
                            'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                            'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                            NewCtr = NewSubItem
                            For x = 1 To 14
                                If NewCtr > 1 Then
                                    NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                    If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                        NewInsert = 1
                                        x = 14
                                    ElseIf IsNumeric(NewStatus) = True Then
                                        x = 14
                                    ElseIf Len(Trim(NewStatus)) = "" Then
                                        x = 14
                                    End If
                                    NewCtr = NewSubItem - x
                                End If
                            Next x
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            End If
                            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                                'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                                NewCtr = NewSubItem
                                For x = 1 To 14
                                    If NewCtr > 1 Then
                                        NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                        If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                            NewInsert = 1
                                            x = 14
                                        ElseIf IsNumeric(NewStatus) = True Then
                                            x = 14
                                        ElseIf Len(Trim(NewStatus)) = "" Then
                                            x = 14
                                        End If
                                        NewCtr = NewSubItem - x
                                    End If
                                Next x
                            
                            End If
                        End If
                    End If
                ElseIf cmbShift.Text = "Night Shift" Then
                    If NewTime < "2359" Then
                        If cmbPosition.Text = "All" Then
                            'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                            'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                            NewCtr = NewSubItem
                            For x = 1 To 14
                                If NewCtr > 1 Then
                                    NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                    If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                        NewInsert = 1
                                        x = 14
                                    ElseIf IsNumeric(NewStatus) = True Then
                                        x = 14
                                    ElseIf Len(Trim(NewStatus)) = "" Then
                                        x = 14
                                    End If
                                    NewCtr = NewSubItem - x
                                End If
                            Next x
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            End If
                            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                NewCtr = NewSubItem
                                For x = 1 To 14
                                    If NewCtr > 1 Then
                                        NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                        If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                            NewInsert = 1
                                            x = 14
                                        ElseIf IsNumeric(NewStatus) = True Then
                                            x = 14
                                        ElseIf Len(Trim(NewStatus)) = "" Then
                                            x = 14
                                        End If
                                        NewCtr = NewSubItem - x
                                    End If
                                Next x
                            End If
                        End If
                    End If
                ElseIf cmbShift.Text = "Late Night Shift" Then
                    If NewTime > "0111" And NewTime < "0500" Then
                        If cmbPosition.Text = "All" Then
                            'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                            'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
                            NewCtr = NewSubItem
                            For x = 1 To 14
                                If NewCtr > 1 Then
                                    NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                    If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                    ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                        NewInsert = 1
                                        x = 14
                                    ElseIf IsNumeric(NewStatus) = True Then
                                        x = 14
                                    ElseIf Len(Trim(NewStatus)) = "" Then
                                        x = 14
                                    End If
                                    NewCtr = NewSubItem - x
                                End If
                            Next x
                        Else
                            If cmbPosition.Text = "Dealer" Then
                                NewPosition = "DLR"
                            ElseIf cmbPosition.Text = "Pit Supervisor" Then
                                NewPosition = "SUP"
                            End If
                            If lvFirstCycle.ListItems(i).SubItems(1) = NewPosition Then
                                NewCtr = NewSubItem
                                For x = 1 To 14
                                    If NewCtr > 1 Then
                                        NewStatus = Replace(lvFirstCycle.ListItems(i).SubItems(NewSubItem - x), "-", "")
                                        If UCase(NewStatus) = "BER" Or UCase(NewStatus) = "ALWP" Or UCase(NewStatus) = "NWH" Or UCase(NewStatus) = "BIR" Or UCase(NewStatus) = "EL" Or UCase(NewStatus) = "SOLO" Or UCase(NewStatus) = "VL" Then
                                        ElseIf UCase(NewStatus) = "AWOL" Or UCase(NewStatus) = "SL" Or UCase(NewStatus) = "SHSL" Or UCase(NewStatus) = "LI" Then
                                            NewInsert = 1
                                            x = 14
                                        ElseIf IsNumeric(NewStatus) = True Then
                                            x = 14
                                        ElseIf Len(Trim(NewStatus)) = "" Then
                                            x = 14
                                        End If
                                        NewCtr = NewSubItem - x
                                    End If
                                Next x
                            End If
                        End If
                    End If
                End If
                
                'Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
                'li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
            Else
            End If
        End If
        
        If NewInsert = 1 Then
            Set li = lvListTrainee.ListItems.Add(, , lvFirstCycle.ListItems(i).Text)
            li.SubItems(1) = lvFirstCycle.ListItems(i).SubItems(1)
            NewInsert = 0
        End If
        
    Next i
End If

End Sub

Private Sub frmNewGroup_Click()

End Sub

Private Sub cmSave_Click()
    On Error GoTo ExportError

    Dim outputFolder As String
    Dim outputFilename As String
    Dim outputPath As String
    Dim reportDate As String
    Dim firstShift As String
    Dim currentShift As String
    Dim hasMultipleShifts As Boolean
    Dim i As Long

    Application.Visible = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 1 To lvListofAbsent.ListItems.Count
        currentShift = Trim$(CStr(lvListofAbsent.ListItems(i).SubItems(2)))
        If Len(currentShift) > 0 Then
            If Len(firstShift) = 0 Then
                firstShift = currentShift
            ElseIf StrComp(firstShift, currentShift, vbTextCompare) <> 0 Then
                hasMultipleShifts = True
                Exit For
            End If
        End If
    Next i

    reportDate = Format$(ThisWorkbook.Worksheets("Sheet1").Range("X1").Value, "d-mmm-yy")
    outputFolder = "\\mcp.com\dept$\FP&A\RFA\"

    If StrComp(cmbShift.Text, "All Shift", vbTextCompare) = 0 _
            Or hasMultipleShifts _
            Or Len(firstShift) = 0 Then
        outputFilename = "RFA Report " & reportDate & ".pdf"
    Else
        outputFilename = "RFA Report " & reportDate & " " & firstShift & ".pdf"
    End If

    outputPath = outputFolder & outputFilename
    If Len(Dir$(outputPath)) > 0 Then Kill outputPath

    ThisWorkbook.Worksheets("Sheet1").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath

    MsgBox "Done copying of file.", vbOKOnly
    ThisWorkbook.Close savechanges:=True
    Application.Quit
    Exit Sub

ExportError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = True
    MsgBox "Unable to export RFA PDF: " & Err.Description, _
        vbExclamation, "RFA Export"
End Sub

Private Sub CommandButton1_Click()
    'Call CallDatabaseLoc
    'Call SaveToWorksheet
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    'ActiveWindow.Visible = True
    Application.Windows(1).Visible = True
    Application.Visible = True
End Sub


Private Sub CommandButton2_Click()
'Dim msg
'    ThisWorkbook.Windows(1).Visible = True
    'Application.Visible = True
'msg = MsgBox("The name of the active window is " & ActiveWindow.Caption)

'ThisWorkbook.Windows(1).Visible = False
End Sub

Private Sub CommandButton3_Click()
    ThisWorkbook.Application.Visible = True
End Sub



Private Sub lvDate_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvListofAvailableShift_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvListTrainee_Click()
    If Not lvListTrainee.SelectedItem Is Nothing Then
        VarListItem = lvListTrainee.SelectedItem.Index
        Call LoadData
    ElseIf lvListTrainee.ListItems.Count = 0 Then
        'MsgBox "Nothing to Delete...", vbExclamation, "EISys Information!"
    Else
        'MsgBox "Please Highlight the Item to Delete...", vbInformation, "EISys Information!"
    End If
End Sub

Private Sub LoadData()
Dim msg
Dim i As Long
Dim bFound1 As Boolean
Dim bFound2 As Boolean
Dim VarNum As Long
    'msg = MsgBox(lvListTrainee.SelectedItem.Text)
    Call ClearShifts
    Call DefaultBackground
    
    For i = 1 To lvFirstCycle.ListItems.Count
            'If InStr(1, lvFirstCycle.ListItems(i).Text, lvListTrainee.SelectedItem.Text, vbTextCompare) Then
            If InStr(1, lvFirstCycle.ListItems(i).Text, lvListTrainee.ListItems(VarListItem).Text, vbTextCompare) Then
                bFound1 = True
                'z = lvFirstCycle.ListItems.Count
                lvFirstCycle.ListItems(i).Selected = True
                'lvFirstCycle.ListItems(i).EnsureVisible
                VarNum = lvFirstCycle.SelectedItem.Index
                'VarNum = VarListItem
                i = lvFirstCycle.ListItems.Count
            Else
                bFound1 = False
            End If
        
    Next i
    
    If bFound1 = True Then
        lbl1.Caption = lvFirstCycle.ListItems(VarNum).SubItems(2)
        lbl2.Caption = lvFirstCycle.ListItems(VarNum).SubItems(3)
        lbl3.Caption = lvFirstCycle.ListItems(VarNum).SubItems(4)
        lbl4.Caption = lvFirstCycle.ListItems(VarNum).SubItems(5)
        lbl5.Caption = lvFirstCycle.ListItems(VarNum).SubItems(6)
        lbl6.Caption = lvFirstCycle.ListItems(VarNum).SubItems(7)
        lbl7.Caption = lvFirstCycle.ListItems(VarNum).SubItems(8)
        lbl8.Caption = lvFirstCycle.ListItems(VarNum).SubItems(9)
        lbl9.Caption = lvFirstCycle.ListItems(VarNum).SubItems(10)
        lbl10.Caption = lvFirstCycle.ListItems(VarNum).SubItems(11)
        lbl11.Caption = lvFirstCycle.ListItems(VarNum).SubItems(12)
        lbl12.Caption = lvFirstCycle.ListItems(VarNum).SubItems(13)
        lbl13.Caption = lvFirstCycle.ListItems(VarNum).SubItems(14)
        lbl14.Caption = lvFirstCycle.ListItems(VarNum).SubItems(15)
    End If
    
    
    For i = 1 To lvSecondCycle.ListItems.Count
            If InStr(1, lvSecondCycle.ListItems(i).Text, lvListTrainee.ListItems(VarListItem).Text, vbTextCompare) Then
                bFound2 = True
                'z = lvFirstCycle.ListItems.Count
                lvSecondCycle.ListItems(i).Selected = True
                VarNum = lvSecondCycle.SelectedItem.Index
                'VarNum = VarListItem
                i = lvSecondCycle.ListItems.Count
            Else
                bFound2 = False
            End If
        
    Next i
    
    If bFound2 = True Then
        lbl15.Caption = lvSecondCycle.ListItems(VarNum).SubItems(2)
        lbl16.Caption = lvSecondCycle.ListItems(VarNum).SubItems(3)
        lbl17.Caption = lvSecondCycle.ListItems(VarNum).SubItems(4)
        lbl18.Caption = lvSecondCycle.ListItems(VarNum).SubItems(5)
        lbl19.Caption = lvSecondCycle.ListItems(VarNum).SubItems(6)
        lbl20.Caption = lvSecondCycle.ListItems(VarNum).SubItems(7)
        lbl21.Caption = lvSecondCycle.ListItems(VarNum).SubItems(8)
        lbl22.Caption = lvSecondCycle.ListItems(VarNum).SubItems(9)
        lbl23.Caption = lvSecondCycle.ListItems(VarNum).SubItems(10)
        lbl24.Caption = lvSecondCycle.ListItems(VarNum).SubItems(11)
        lbl25.Caption = lvSecondCycle.ListItems(VarNum).SubItems(12)
        lbl26.Caption = lvSecondCycle.ListItems(VarNum).SubItems(13)
        lbl27.Caption = lvSecondCycle.ListItems(VarNum).SubItems(14)
        lbl28.Caption = lvSecondCycle.ListItems(VarNum).SubItems(15)
    End If
    
    
    lblid.Caption = Mid(lvListTrainee.ListItems(VarListItem).Text, 2, 7)
    lblname.Caption = Mid(lvListTrainee.ListItems(VarListItem).Text, 11, 20)
    lblposition.Caption = lvListTrainee.ListItems(VarListItem).SubItems(1)
    
    Call LoadColor
End Sub

Private Sub LoadColor()
    If lvListTrainee.SelectedItem.SubItems(3) = "First Cycle" Then
        If lvListTrainee.SelectedItem.SubItems(2) = "2" Then
            frmdate1.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "3" Then
            frmdate2.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "4" Then
            frmDate3.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "5" Then
            frmDate4.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "6" Then
            frmDate5.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "7" Then
            frmDate6.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "8" Then
            frmDate7.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "9" Then
            frmDate8.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "10" Then
            frmDate9.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "11" Then
            frmDate10.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "12" Then
            frmDate11.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "13" Then
            frmDate12.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "14" Then
            frmDate13.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "15" Then
            frmDate14.BackColor = &H8000000D
        End If
    ElseIf lvListTrainee.SelectedItem.SubItems(3) = "Second Cycle" Then
        If lvListTrainee.SelectedItem.SubItems(2) = "2" Then
            frmDate15.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "3" Then
            frmDate16.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "4" Then
            frmDate17.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "5" Then
            frmDate18.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "6" Then
            frmDate19.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "7" Then
            frmDate20.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "8" Then
            frmDate21.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "9" Then
            frmDate22.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "10" Then
            frmDate23.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "11" Then
            frmDate24.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "12" Then
            frmDate25.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "13" Then
            frmDate26.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "14" Then
            frmDate27.BackColor = &H8000000D
        ElseIf lvListTrainee.SelectedItem.SubItems(2) = "15" Then
            frmDate28.BackColor = &H8000000D
        End If
    End If
    
    
'"&H80000004&"
End Sub
'Private Sub txtIDNumber_Enter()
'    cmdInsert.Default = True
'End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub lvListTrainee_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    'If lvListTrainee.ListItems.Count = 0 Then
    'Else
    '    If KeyCode = vbKeyDown Then
    '        If VarListItem = lvListTrainee.ListItems.Count Then
    '            VarListItem = lvListTrainee.ListItems.Count
    '        Else
    '            VarListItem = lvListTrainee.SelectedItem.Index + 1
    '        End If
    '    ElseIf KeyCode = vbKeyUp Then
    '        If VarListItem = 1 Then
    '            VarListItem = 1
    '        Else
    '            VarListItem = lvListTrainee.SelectedItem.Index - 1
    '        End If
    '    ElseIf KeyCode = vbKeyEnd Then
    '        VarListItem = lvListTrainee.ListItems.Count
    '    ElseIf KeyCode = vbKeyHome Then
    '        VarListItem = 1
    '    End If
        
    '    Call LoadData
    'End If
    
End Sub

Private Sub optAll_Click()
    lvListTrainee.ListItems.Clear
    txtSearch.Enabled = False
    txtSearch.Text = ""
    cmbDate.Enabled = False
    cmbPosition.Enabled = False
    cmbShift.Enabled = False
    Call ClearAll
End Sub

Private Sub optSL_Click()
    lvListTrainee.ListItems.Clear
    txtSearch.Enabled = False
    txtSearch.Text = ""
    cmbDate.Enabled = True
    cmbPosition.Enabled = True
    cmbShift.Enabled = True
    Call ClearAll
End Sub

Private Sub txtSearch_Change()
Dim msg
Dim i As Long
Dim bFound As Boolean

    Call StopTimer
    bFound = False
    For i = 1 To lvListTrainee.ListItems.Count
        If InStr(1, lvListTrainee.ListItems(i).Text, txtSearch.Text, vbTextCompare) Then
            bFound = True
            lvListTrainee.ListItems(i).Selected = True
            lvListTrainee.ListItems(i).EnsureVisible
            i = lvListTrainee.ListItems.Count
            Call lvListTrainee_Click
        Else
        End If
    Next i
    
    If bFound = False Then
        msg = MsgBox("No Record Found!", vbOKOnly + vbExclamation, "ETS Guide!")
        txtSearch.Text = ""
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
        txtSearch.SetFocus
    End If
    Call SetTimer

End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'If lvListTrainee.ListItems.Count = 0 Then
    'Else
    '    If KeyCode = KeyCodeConstants.vbKeyDown Then
    '        If VarListItem = lvListTrainee.ListItems.Count Then
    '            VarListItem = lvListTrainee.ListItems.Count
    '        Else
    '            VarListItem = VarListItem + 1
    '        End If
    '        KeyCode = 0
    '    ElseIf KeyCode = KeyCodeConstants.vbKeyUp Then
    '        If VarListItem = 1 Then
    '            VarListItem = 1
    '        Else
    '            VarListItem = VarListItem - 1
    '        End If
    '        KeyCode = 0
    '    ElseIf KeyCode = KeyCodeConstants.vbKeyEnd Then
    '        VarListItem = lvListTrainee.ListItems.Count
    '        KeyCode = 0
    '    ElseIf KeyCode = KeyCodeConstants.vbKeyHome Then
    '        VarListItem = 1
    '        KeyCode = 0
    '    End If
    '    Call LoadData
    'End If
End Sub

'Private Sub txtIDNumber_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'    cmdInsert.Default = False
'End Sub

Private Sub UserForm_Activate()
    AddToForm MIN_BOX
    'Application.Visible = False
    'ThisWorkbook.Windows(1).Visible = False
    
    'ThisWorkbook.Application.Visible = False
    'ThisWorkbook.Application.VBE.MainWindow.Visible = False
    AppTasklist Me
End Sub

Private Sub UserForm_Initialize()

Dim NewDate As String
Dim bFound As Boolean
Dim startupErrorDescription As String

    'optSL.Value = True
    On Error GoTo ErrHandler
    lvAll.Visible = False
    lvListofAbsent.Visible = False
    lvFirstCycle.Visible = False
    
    lvSecondCycle.Visible = False
    lvDate.Visible = False
    
    cmdPrint.Enabled = False
    Dim NewTime
    
    With Me.lvListofAvailableShift
        .ColumnHeaders.Add , , "", 50
        .HideColumnHeaders = True
        .FullRowSelect = True
        .View = lvwReport
    End With
    
    
    With Me.lvListTrainee
        .ColumnHeaders.Add , , "", 250
        .ColumnHeaders.Add , , "", 60
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "", 0
        .Gridlines = True
        .HideColumnHeaders = True
        .FullRowSelect = True
        .View = lvwReport
    End With
    
    With Me.lvFirstCycle
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
    End With
    
    With Me.lvSecondCycle
        .Gridlines = True
        .HideColumnHeaders = True
        .View = lvwReport
    End With
    
    With Me.lvDate
        .ColumnHeaders.Add , , "Date", 100
        .ColumnHeaders.Add , , "Column", 100
        .ColumnHeaders.Add , , "Cycle", 150
        .Gridlines = True
        .HideColumnHeaders = True
        .View = lvwReport
    End With
    
    With Me.lvAll
        .ColumnHeaders.Add , , "Date", 100
        .ColumnHeaders.Add , , "Name", 100
        .ColumnHeaders.Add , , "Name", 100
        .ColumnHeaders.Add , , "Shift", 100
        .ColumnHeaders.Add , , "Cycle", 150
        .ColumnHeaders.Add , , "Cycle", 150
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
    End With
    
    With Me.lvListofAbsent
        .ColumnHeaders.Add , , "ID", 100
        .ColumnHeaders.Add , , "Name", 100
        .ColumnHeaders.Add , , "Position", 100
        .ColumnHeaders.Add , , "Shift", 300
        .ColumnHeaders.Add , , "Remarks", 100
        .Gridlines = True
        .HideColumnHeaders = True
        .View = lvwReport
    End With
    
    
    'txtDate.Text = Format(Date, "MM/DD/YYYY")
    lvListTrainee.ListItems.Clear
    'frmTrack.Visible = False
    If Not LoadStartupDataWithRetry() Then
        SetTrackingEnabled False
    End If
    Exit Sub

ErrHandler:
    startupErrorDescription = Err.Description
    SetTrackingEnabled False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Visible = True
    On Error Resume Next
    ThisWorkbook.Windows(1).Visible = True
    On Error GoTo 0
    MsgBox "The RFA form could not be initialized: " & _
        startupErrorDescription, _
        vbCritical, "RFA Startup"
End Sub
Private Sub CallDatabaseLoc()
    Dim firstRosterPath As String
    Dim secondRosterPath As String
    Dim errorMessage As String

    If Not TryLoadConfiguration(firstRosterPath, secondRosterPath, errorMessage) Then
        Err.Raise vbObjectError + 701, "CallDatabaseLoc", errorMessage
    End If

    NewDatabaseLoc1 = firstRosterPath
    NewDatabaseLoc2 = secondRosterPath
End Sub

Sub ReadDataFromCloseFile()
    Call LoadStartupDataWithRetry
End Sub

Private Function LoadStartupDataWithRetry() As Boolean
    Dim errorMessage As String
    Dim userChoice As VbMsgBoxResult

    Call StopTimer
    SetTrackingEnabled False

    Do
        errorMessage = ""
        If TryLoadStartupData(errorMessage) Then
            mHasValidRosterData = True
            SetTrackingEnabled True
            Call CallShiftStarts
            Call SetTimer
            LoadStartupDataWithRetry = True
            Exit Function
        End If

        mHasValidRosterData = False
        userChoice = MsgBox( _
            errorMessage & vbCrLf & vbCrLf & _
            "Correct the file or network problem, then select Retry." & vbCrLf & _
            "Select Cancel to leave tracking disabled.", _
            vbCritical + vbRetryCancel, _
            "RFA Data Loading")

        If userChoice <> vbRetry Then
            SetTrackingEnabled False
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            Application.AskToUpdateLinks = True
            Application.Visible = True
            On Error Resume Next
            ThisWorkbook.Windows(1).Visible = True
            On Error GoTo 0
            LoadStartupDataWithRetry = False
            Exit Function
        End If
    Loop
End Function

Private Function TryLoadStartupData(ByRef errorMessage As String) As Boolean
    Dim firstRosterPath As String
    Dim secondRosterPath As String
    Dim firstRosterData As Variant
    Dim secondRosterData As Variant
    Dim firstStartDate As Date
    Dim secondStartDate As Date
    Dim loadSucceeded As Boolean
    Dim unexpectedNumber As Long
    Dim unexpectedDescription As String

    On Error GoTo UnexpectedFailure

    mIsLoadingData = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False

    If Not TryLoadConfiguration( _
            firstRosterPath, secondRosterPath, errorMessage) Then
        GoTo CleanUp
    End If

    If Not TryLoadRoster( _
            firstRosterPath, "First-cycle roster", firstRosterData, _
            firstStartDate, errorMessage) Then
        GoTo CleanUp
    End If

    If Not TryLoadRoster( _
            secondRosterPath, "Second-cycle roster", secondRosterData, _
            secondStartDate, errorMessage) Then
        GoTo CleanUp
    End If

    NewDatabaseLoc1 = firstRosterPath
    NewDatabaseLoc2 = secondRosterPath

    PopulateRosterControls firstRosterData, secondRosterData, firstStartDate
    InitializeFilterChoices
    SelectDefaultDate
    loadSucceeded = True
    GoTo CleanUp

UnexpectedFailure:
    unexpectedNumber = Err.Number
    unexpectedDescription = Err.Description
    errorMessage = BuildLoadError( _
        "Unexpected startup error", "", unexpectedNumber, _
        unexpectedDescription, _
        "Close other Excel dialogs and verify the network connection.")

CleanUp:
    mIsLoadingData = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    TryLoadStartupData = loadSucceeded
End Function

Private Function TryLoadConfiguration( _
        ByRef firstRosterPath As String, _
        ByRef secondRosterPath As String, _
        ByRef errorMessage As String) As Boolean

    Const CONFIG_PATH As String = "\\mcp.com\dept$\FP&A\RFA\Database.xlsx"

    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim loadSucceeded As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo OpenFailure

    If Not FileIsAccessible(CONFIG_PATH) Then
        errorMessage = BuildLoadError( _
            "Configuration file is missing or inaccessible", CONFIG_PATH, _
            0, "", _
            "Connect to the company network or VPN and verify file permissions.")
        GoTo CleanUp
    End If

    Set sourceWorkbook = Workbooks.Open( _
        Filename:=CONFIG_PATH, UpdateLinks:=0, ReadOnly:=True, _
        AddToMru:=False)

    Set sourceSheet = GetWorksheet(sourceWorkbook, "Sheet1")
    If sourceSheet Is Nothing Then
        errorMessage = BuildLoadError( _
            "Configuration workbook is invalid", CONFIG_PATH, 0, "", _
            "Add a worksheet named Sheet1 containing roster paths in A1 and B1.")
        GoTo CleanUp
    End If

    If IsError(sourceSheet.Range("A1").Value) _
            Or IsError(sourceSheet.Range("B1").Value) Then
        errorMessage = BuildLoadError( _
            "Configuration workbook contains invalid roster paths", _
            CONFIG_PATH, 0, "", _
            "Replace errors in Sheet1 cells A1 and B1 with valid file paths.")
        GoTo CleanUp
    End If

    firstRosterPath = Trim$(CStr(sourceSheet.Range("A1").Value2))
    secondRosterPath = Trim$(CStr(sourceSheet.Range("B1").Value2))

    If Len(firstRosterPath) = 0 Or Len(secondRosterPath) = 0 Then
        errorMessage = BuildLoadError( _
            "Configuration workbook is incomplete", CONFIG_PATH, 0, "", _
            "Enter the first-cycle path in A1 and second-cycle path in B1.")
        GoTo CleanUp
    End If

    If Not FileIsAccessible(firstRosterPath) Then
        errorMessage = BuildLoadError( _
            "First-cycle roster is missing or inaccessible", _
            firstRosterPath, 0, "", _
            "Verify the path in Database.xlsx Sheet1!A1 and network access.")
        GoTo CleanUp
    End If

    If Not FileIsAccessible(secondRosterPath) Then
        errorMessage = BuildLoadError( _
            "Second-cycle roster is missing or inaccessible", _
            secondRosterPath, 0, "", _
            "Verify the path in Database.xlsx Sheet1!B1 and network access.")
        GoTo CleanUp
    End If

    loadSucceeded = True
    GoTo CleanUp

OpenFailure:
    failureNumber = Err.Number
    failureDescription = Err.Description
    errorMessage = BuildLoadError( _
        "Configuration file could not be opened", CONFIG_PATH, _
        failureNumber, failureDescription, _
        "Verify network access, file permissions, and that the workbook is valid.")

CleanUp:
    On Error Resume Next
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    Set sourceSheet = Nothing
    Set sourceWorkbook = Nothing
    On Error GoTo 0
    TryLoadConfiguration = loadSucceeded
End Function

Private Function TryLoadRoster( _
        ByVal rosterPath As String, _
        ByVal rosterRole As String, _
        ByRef rosterData As Variant, _
        ByRef startDate As Date, _
        ByRef errorMessage As String) As Boolean

    Dim sourceWorkbook As Workbook
    Dim rosterSheet As Worksheet
    Dim rosterRange As Range
    Dim loadSucceeded As Boolean
    Dim hasEmployee As Boolean
    Dim rowIndex As Long
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo OpenFailure

    If Not FileIsAccessible(rosterPath) Then
        errorMessage = BuildLoadError( _
            rosterRole & " is missing or inaccessible", rosterPath, 0, "", _
            "Verify the configured path, network connection, and permissions.")
        GoTo CleanUp
    End If

    Set sourceWorkbook = Workbooks.Open( _
        Filename:=rosterPath, UpdateLinks:=0, ReadOnly:=True, _
        Password:="<REDACTED>", WriteResPassword:="<REDACTED>", _
        AddToMru:=False)

    Set rosterSheet = GetWorksheet(sourceWorkbook, "Current Roster")
    If rosterSheet Is Nothing Then
        errorMessage = BuildLoadError( _
            rosterRole & " is invalid", rosterPath, 0, "", _
            "Add a worksheet named Current Roster.")
        GoTo CleanUp
    End If

    Set rosterRange = rosterSheet.Range("A1").CurrentRegion
    If rosterRange.Rows.Count < 3 Or rosterRange.Columns.Count < 16 Then
        errorMessage = BuildLoadError( _
            rosterRole & " has an invalid layout", rosterPath, 0, "", _
            "Current Roster must contain at least 3 rows and 16 columns.")
        GoTo CleanUp
    End If

    rosterData = rosterRange.Value2
    If Not IsDate(rosterData(2, 3)) Then
        errorMessage = BuildLoadError( _
            rosterRole & " has an invalid start date", rosterPath, 0, "", _
            "Current Roster cell C2 must contain a valid cycle start date.")
        GoTo CleanUp
    End If

    For rowIndex = 3 To UBound(rosterData, 1)
        If Len(SafeCellText(rosterData(rowIndex, 1))) > 0 _
                And SafeCellText(rosterData(rowIndex, 1)) <> "-" Then
            hasEmployee = True
            Exit For
        End If
    Next rowIndex

    If Not hasEmployee Then
        errorMessage = BuildLoadError( _
            rosterRole & " contains no employee rows", rosterPath, 0, "", _
            "Add at least one employee beginning on row 3 of Current Roster.")
        GoTo CleanUp
    End If

    startDate = CDate(rosterData(2, 3))
    loadSucceeded = True
    GoTo CleanUp

OpenFailure:
    failureNumber = Err.Number
    failureDescription = Err.Description
    errorMessage = BuildLoadError( _
        rosterRole & " could not be opened or read", rosterPath, _
        failureNumber, failureDescription, _
        "Verify the file format, access permissions, and workbook credentials.")

CleanUp:
    On Error Resume Next
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    Set rosterRange = Nothing
    Set rosterSheet = Nothing
    Set sourceWorkbook = Nothing
    On Error GoTo 0
    TryLoadRoster = loadSucceeded
End Function

Private Sub PopulateRosterControls( _
        ByVal firstRosterData As Variant, _
        ByVal secondRosterData As Variant, _
        ByVal firstStartDate As Date)

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim dayIndex As Long
    Dim cycleColumn As Long
    Dim item As ListItem
    Dim employeeText As String
    Dim mappedDate As Date

    lvFirstCycle.ListItems.Clear
    lvFirstCycle.ColumnHeaders.Clear
    lvSecondCycle.ListItems.Clear
    lvSecondCycle.ColumnHeaders.Clear
    lvDate.ListItems.Clear
    lvListofAvailableShift.ListItems.Clear
    lvListTrainee.ListItems.Clear
    lvAll.ListItems.Clear
    lvListofAbsent.ListItems.Clear
    cmbDate.Clear

    For columnIndex = 1 To UBound(firstRosterData, 2)
        lvFirstCycle.ColumnHeaders.Add , , _
            SafeCellText(firstRosterData(1, columnIndex)), 90
    Next columnIndex

    For rowIndex = 3 To UBound(firstRosterData, 1)
        employeeText = SafeCellText(firstRosterData(rowIndex, 1))
        If Len(employeeText) > 0 And employeeText <> "-" Then
            Set item = lvFirstCycle.ListItems.Add(Text:=employeeText)
            For columnIndex = 2 To UBound(firstRosterData, 2)
                item.ListSubItems.Add Text:= _
                    SafeCellText(firstRosterData(rowIndex, columnIndex))
            Next columnIndex
        End If
    Next rowIndex

    For columnIndex = 1 To UBound(secondRosterData, 2)
        lvSecondCycle.ColumnHeaders.Add , , _
            SafeCellText(secondRosterData(1, columnIndex)), 90
    Next columnIndex

    For rowIndex = 3 To UBound(secondRosterData, 1)
        employeeText = SafeCellText(secondRosterData(rowIndex, 1))
        If Len(employeeText) > 0 And employeeText <> "-" Then
            Set item = lvSecondCycle.ListItems.Add(Text:=employeeText)
            For columnIndex = 2 To UBound(secondRosterData, 2)
                item.ListSubItems.Add Text:= _
                    SafeCellText(secondRosterData(rowIndex, columnIndex))
            Next columnIndex
        End If
    Next rowIndex

    cycleColumn = 1
    For dayIndex = 1 To 28
        mappedDate = DateAdd("d", dayIndex - 1, firstStartDate)
        Set item = lvDate.ListItems.Add(, , Format$(mappedDate, "mm/dd/yyyy"))
        cmbDate.AddItem Format$(mappedDate, "mm/dd/yyyy")
        item.SubItems(1) = cycleColumn + 1
        If dayIndex <= 14 Then
            item.SubItems(2) = "First Cycle"
        Else
            item.SubItems(2) = "Second Cycle"
        End If
        If cycleColumn = 14 Then
            cycleColumn = 1
        Else
            cycleColumn = cycleColumn + 1
        End If
    Next dayIndex

    ClearAll
End Sub

Private Sub InitializeFilterChoices()
    Dim currentTime As String

    cmbShift.Clear
    cmbShift.AddItem "All Shift"
    cmbShift.AddItem "Morning Shift"
    cmbShift.AddItem "Day Shift"
    cmbShift.AddItem "Late Day Shift"
    cmbShift.AddItem "Night Shift"
    cmbShift.AddItem "Late Night Shift"
    cmbShift.Text = "All Shift"

    cmbPosition.Clear
    cmbPosition.AddItem "All"
    cmbPosition.AddItem "Dealer"
    cmbPosition.AddItem "Pit Supervisor"
    cmbPosition.AddItem "Pit Manager"
    cmbPosition.Text = "All"

    currentTime = Format$(Time(), "HHMM")
    If currentTime > "0500" And currentTime < "1200" Then
        cmbShift.Text = "Morning Shift"
    ElseIf currentTime > "1159" And currentTime < "1800" Then
        cmbShift.Text = "Day Shift"
    ElseIf currentTime > "1759" And currentTime < "1930" Then
        cmbShift.Text = "Late Day Shift"
    ElseIf currentTime > "1929" And currentTime < "2359" Then
        cmbShift.Text = "Night Shift"
    ElseIf currentTime > "0111" And currentTime < "0500" Then
        cmbShift.Text = "Late Night Shift"
    End If

    txtSearch.Text = ""
    txtSearch.Enabled = False
End Sub

Private Sub SelectDefaultDate()
    Dim todayText As String
    Dim itemIndex As Long
    Dim foundToday As Boolean

    todayText = Format$(Date, "MM/DD/YYYY")
    For itemIndex = 0 To cmbDate.ListCount - 1
        If StrComp(todayText, cmbDate.List(itemIndex), vbTextCompare) = 0 Then
            cmbDate.Text = todayText
            foundToday = True
            Exit For
        End If
    Next itemIndex

    If Not foundToday And cmbDate.ListCount > 0 Then
        MsgBox "Today's Date (" & todayText & _
            ") is not available in the validated roster period.", _
            vbExclamation, "ETS Confirmation!"
        cmbDate.Text = cmbDate.List(cmbDate.ListCount - 1)
    End If
End Sub

Private Sub SetTrackingEnabled(ByVal isEnabled As Boolean)
    On Error Resume Next
    cmbDate.Enabled = isEnabled
    cmbPosition.Enabled = isEnabled
    cmbShift.Enabled = isEnabled
    cmdTrackNow.Enabled = isEnabled
    cmdPrint.Enabled = False
    txtSearch.Enabled = False
    cmdRefresh.Enabled = True
    On Error GoTo 0
End Sub

Private Function FileIsAccessible(ByVal filePath As String) As Boolean
    On Error GoTo NotAccessible
    If Len(Trim$(filePath)) = 0 Then Exit Function
    FileIsAccessible = Len(Dir$(filePath, _
        vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0
    Exit Function

NotAccessible:
    FileIsAccessible = False
End Function

Private Function GetWorksheet( _
        ByVal sourceWorkbook As Workbook, _
        ByVal worksheetName As String) As Worksheet

    On Error Resume Next
    Set GetWorksheet = sourceWorkbook.Worksheets(worksheetName)
    On Error GoTo 0
End Function

Private Function SafeCellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Or IsNull(cellValue) Or IsEmpty(cellValue) Then
        SafeCellText = ""
    Else
        SafeCellText = Trim$(CStr(cellValue))
    End If
End Function

Private Function BuildLoadError( _
        ByVal heading As String, _
        ByVal filePath As String, _
        ByVal errorNumber As Long, _
        ByVal errorDescription As String, _
        ByVal correctiveAction As String) As String

    Dim message As String

    message = heading
    If Len(filePath) > 0 Then message = message & vbCrLf & "File: " & filePath
    If errorNumber <> 0 Then _
        message = message & vbCrLf & "Excel error " & _
            CStr(errorNumber) & ": " & errorDescription
    If Len(correctiveAction) > 0 Then _
        message = message & vbCrLf & "Required action: " & correctiveAction
    BuildLoadError = message
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Application.Visible = True
    'ThisWorkbook.Close savechanges:=False
    'ThisWorkbook.Close savechanges:=False
    '.DisplayAlerts = False
    ThisWorkbook.Application.Quit
    ThisWorkbook.Close savechanges = False
    
End Sub

Private Sub UserForm_Terminate()
    'ActiveWorkbook.Close
    ThisWorkbook.Application.Quit
    ThisWorkbook.Close savechanges = False
    'Application.Visible = True
End Sub
