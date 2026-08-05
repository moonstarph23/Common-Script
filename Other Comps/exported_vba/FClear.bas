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
    wsInput.Columns("A:B").ClearContents

' Define the worksheets
    Set wsTemp = ThisWorkbook.Sheets("MENU ITEM 2")

    ' Unfilter the TEMPLATE sheet if it is filtered
    If wsTemp.AutoFilterMode Then
        If wsTemp.FilterMode Then
            wsTemp.ShowAllData
        End If
    End If

    ' Define the worksheets
    Set wsSum = ThisWorkbook.Sheets("RAW")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsSum.AutoFilterMode Then
        If wsSum.FilterMode Then
            wsSum.ShowAllData
        End If
    End If

    ' Define the worksheets
    Set wsCD = ThisWorkbook.Sheets("EMP CLOSED CHECK")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsCD.AutoFilterMode Then
        If wsCD.FilterMode Then
            wsCD.ShowAllData
        End If
    End If

End Sub
