Sub RunMacro()
    Dim ws As Worksheet

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
End Sub
