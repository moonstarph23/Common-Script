Sub ClearSEM()

'

'

    ' Define the worksheets
    Set wsInput = ThisWorkbook.Sheets("SFR (MTD)")

    ' Unfilter the SFR sheet if it is filtered
    If wsInput.AutoFilterMode Then
        If wsInput.FilterMode Then
           wsInput.ShowAllData
        End If
    End If
    Sheets("SFR (MTD)").Select
    Columns("A:B").Select
    Selection.ClearContents

' Define the worksheets
    Set wsTemp = ThisWorkbook.Sheets("MENU ITEM 2")

    ' Unfilter the TEMPLATE sheet if it is filtered
    If wsTemp.AutoFilterMode Then
        If wsTemp.FilterMode Then
            wsTemp.ShowAllData
        End If
    End If

'Find the last used row in column A
    lastRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRow >= 2 Then
        wsTemp.Range("A2:S" & lastRow).ClearContents
    End If

    ' Define the worksheets
    Set wsSum = ThisWorkbook.Sheets("RAW")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsSum.AutoFilterMode Then
        If wsSum.FilterMode Then
            wsSum.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowC = wsSum.Cells(wsSum.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowC >= 2 Then
        wsSum.Range("A2:AO" & lastRowC).ClearContents
    End If

    ' Define the worksheets
    Set wsCD = ThisWorkbook.Sheets("EMP CLOSED CHECK")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsCD.AutoFilterMode Then
        If wsCD.FilterMode Then
            wsCD.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowS = wsCD.Cells(wsCD.Rows.Count, "B").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowS >= 4 Then
        wsCD.Range("A4:M" & lastRowS).ClearContents
        wsCD.Range("M3").ClearContents
    End If
End Sub
