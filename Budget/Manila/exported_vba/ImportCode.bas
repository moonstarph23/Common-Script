Attribute VB_Name = "ImportCode"
Sub ImportAndReplaceModule()
    Dim fd As FileDialog
    Dim selectedFile As String
    Dim vbProj As Object
    Dim moduleName As String
    Dim comp As Object
    Dim found As Boolean

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a .bas file to import"
    fd.Filters.Clear
    fd.Filters.Add "VBA Module", "*.bas"
    fd.AllowMultiSelect = False

    If fd.Show = -1 Then
        selectedFile = fd.SelectedItems(1)
        moduleName = Mid(selectedFile, InStrRev(selectedFile, "\") + 1)
        moduleName = Left(moduleName, Len(moduleName) - 4) ' Remove .bas extension

        Set vbProj = ThisWorkbook.VBProject
        found = False

        ' Check for existing module and remove it
        For Each comp In vbProj.VBComponents
            If comp.Name = moduleName Then
                vbProj.VBComponents.Remove comp
                found = True
                Exit For
            End If
        Next comp

        vbProj.VBComponents.Import selectedFile

        If found Then
            MsgBox "Module '" & moduleName & "' was replaced.", vbInformation
        Else
            MsgBox "Module '" & moduleName & "' was imported.", vbInformation
        End If
    Else
        MsgBox "No file selected.", vbExclamation
    End If
End Sub

