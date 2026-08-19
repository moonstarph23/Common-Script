Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    Call RegisterOCX
End Sub

Sub RegisterOCX()
    Dim ocxPath As String
    Dim fd As FileDialog
    Dim ref As Object
    Dim refFound As Boolean

    ' Check if the OCX is already registered
    refFound = False
    On Error Resume Next
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Name = "MSComctlLib" Then
            refFound = True
            Exit For
        End If
    Next ref
    On Error GoTo 0

    ' If already registered, exit the subroutine
    If refFound Then
        'MsgBox "MSCOMCTL.OCX is already registered.", vbInformation
        Exit Sub
    End If

    ' Set default path to the OCX file
    ocxPath = "C:\Windows\SysWOW64\MSCOMCTL.OCX"
 

    ' Check if the OCX file exists
    If Dir(ocxPath) = "" Then
        MsgBox "Default OCX file not found at: " & ocxPath, vbExclamation
        
        ' Prompt the user to select the OCX file
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Select MSCOMCTL.OCX file"
            .Filters.Clear
            .Filters.Add "OCX Files", "*.ocx", 1
            .AllowMultiSelect = False
            If .Show = True Then
                ocxPath = .SelectedItems(1)
                If Dir(ocxPath) = "" Then
                    MsgBox "Selected file does not exist, please verify the path.", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "Operation canceled. The OCX file is required to continue.", vbCritical
                Exit Sub
            End If
        End With
    End If

    ' Attempt to register the OCX file
    On Error Resume Next
    Shell "regsvr32 /s """ & ocxPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Failed to register OCX. Error: " & Err.Description, vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Check if registration was successful
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromFile ocxPath
    If Err.Number = 0 Then
        MsgBox "MSCOMCTL.OCX registered successfully.", vbInformation
    Else
        MsgBox "Error adding reference to VBA project: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
End Sub
