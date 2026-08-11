Attribute VB_Name = "RUN"
Sub RunMacro()
    Dim ws As Worksheet

    Call ClearSEM

    For Each ws In ThisWorkbook.Worksheets
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    Next ws

    Call CopysfrSheet
    Call CopyFilterSheet
    Call CopyMISheet
    Call CopyRAWSheet
    Call CopyECCSheet
    Call CopyCDSheet
    Call CopyFBSheet
    Call CopyHASheet
    Call CopySPSheet
    Call CopyOTHERSheet
End Sub
