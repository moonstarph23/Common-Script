Sub Clear()

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
    Set wsCD = ThisWorkbook.Sheets("Casino Drink")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsCD.AutoFilterMode Then
        If wsCD.FilterMode Then
            wsCD.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowS = wsCD.Cells(wsCD.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowS >= 8 Then
        wsCD.Range("A8:R" & lastRowS).ClearContents
    End If

        ' Define the worksheets
    Set wsFB = ThisWorkbook.Sheets("Csino F&B COMP")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsFB.AutoFilterMode Then
        If wsFB.FilterMode Then
            wsFB.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowF = wsFB.Cells(wsFB.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowF >= 10 Then
        wsFB.Range("A10:T" & lastRowF).ClearContents
    End If

            ' Define the worksheets
    Set wsHA = ThisWorkbook.Sheets("Htl Amenity")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsHA.AutoFilterMode Then
        If wsHA.FilterMode Then
            wsHA.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowH = wsHA.Cells(wsHA.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowH >= 10 Then
        wsHA.Range("A10:R" & lastRowH).ClearContents
    End If

                ' Define the worksheets
    Set wsSP = ThisWorkbook.Sheets("COMP SPA PACKAGE")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsSP.AutoFilterMode Then
        If wsSP.FilterMode Then
            wsSP.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowP = wsSP.Cells(wsSP.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowP >= 8 Then
        wsSP.Range("A8:N" & lastRowP).ClearContents
    End If

                ' Define the worksheets
    Set wsOT = ThisWorkbook.Sheets("OTHERS")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsOT.AutoFilterMode Then
        If wsOT.FilterMode Then
            wsOT.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowO = wsOT.Cells(wsOT.Rows.Count, "A").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowO >= 8 Then
        wsOT.Range("A8:N" & lastRowO).ClearContents
    End If

    ' Define the worksheets
    Set wsEF = ThisWorkbook.Sheets("EMP CLOSED CHECK")

    ' Unfilter the SUMMARY sheet if it is filtered
    If wsEF.AutoFilterMode Then
        If wsEF.FilterMode Then
            wsEF.ShowAllData
        End If
    End If

    'Find the last used row in column A
    lastRowX = wsEF.Cells(wsEF.Rows.Count, "B").End(xlUp).Row

    ' Check if there are rows to clear
    If lastRowX >= 4 Then
        wsEF.Range("A4:L" & lastRowX).ClearContents
    End If
End Sub
