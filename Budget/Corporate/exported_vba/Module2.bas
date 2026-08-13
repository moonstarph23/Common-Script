Attribute VB_Name = "Module2"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    activeSheet.Shapes.Range(Array("Refresh Graph")).Select
    activeSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select
    ActiveChart.FullSeriesCollection(0).Points(0).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).Points(4).Select
    ActiveChart.FullSeriesCollection(0).Points(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).Points(4).Select
    ActiveChart.FullSeriesCollection(0).Points(3).ApplyDataLabels
    Range("Z3").Select
    activeSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(4).Select
    ActiveChart.FullSeriesCollection(0).Points(3).ApplyDataLabels
End Sub


Public Sub FilterSheets()
    Dim ws As Worksheet
    Dim sheetToRename As Worksheet
    Dim currentSheetName As String
    Dim newSheetName As String
    Dim lastRow As Long
    Dim i As Long
    

               
            activeSheet.Range("$A$7:$XFB$1356").AutoFilter Field:=63, Criteria1:="SHOW"


   
End Sub
Sub UnfilterActiveSheet()
    With ThisWorkbook.activeSheet
        On Error Resume Next
        If .AutoFilterMode Then .ShowAllData
        On Error GoTo 0
    End With
End Sub


Sub Refresh_By_Dept_Graph()
'
' Refresh_Graph Macro
'

'Unfilter all data
    activeSheet.Range("$I$8:$K$33").AutoFilter Field:=1
    activeSheet.Range("$I$8:$K$33").AutoFilter Field:=2
    activeSheet.Range("$I$8:$K$33").AutoFilter Field:=3
 
 'Set Total for the bars
    activeSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    ActiveChart.FullSeriesCollection(1).Points(2).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(3).Select
    ActiveChart.FullSeriesCollection(1).Points(3).IsTotal = False
    
    ActiveChart.FullSeriesCollection(1).Points(9).Select
    ActiveChart.FullSeriesCollection(1).Points(9).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(10).Select
    ActiveChart.FullSeriesCollection(1).Points(10).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(11).Select
    ActiveChart.FullSeriesCollection(1).Points(11).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(12).Select
    ActiveChart.FullSeriesCollection(1).Points(12).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(13).Select
    ActiveChart.FullSeriesCollection(1).Points(13).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(14).Select
    ActiveChart.FullSeriesCollection(1).Points(14).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(15).Select
    ActiveChart.FullSeriesCollection(1).Points(15).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(16).Select
    ActiveChart.FullSeriesCollection(1).Points(16).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(17).Select
    ActiveChart.FullSeriesCollection(1).Points(17).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(18).Select
    ActiveChart.FullSeriesCollection(1).Points(18).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(19).Select
    ActiveChart.FullSeriesCollection(1).Points(19).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(20).Select
    ActiveChart.FullSeriesCollection(1).Points(20).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(21).Select
    ActiveChart.FullSeriesCollection(1).Points(21).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(22).Select
    ActiveChart.FullSeriesCollection(1).Points(22).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(23).Select
    ActiveChart.FullSeriesCollection(1).Points(23).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(24).Select
    ActiveChart.FullSeriesCollection(1).Points(24).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(25).Select
    ActiveChart.FullSeriesCollection(1).Points(25).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(26).Select
    ActiveChart.FullSeriesCollection(1).Points(26).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(27).Select
    ActiveChart.FullSeriesCollection(1).Points(27).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(28).Select
    ActiveChart.FullSeriesCollection(1).Points(28).IsTotal = True

    
'Show non zero data
    activeSheet.Range("$K$8:$K$33").AutoFilter Field:=3, Criteria1:="Show"
    



    
End Sub

Sub Refresh_By_Dept_Graph_in_thousands()
'
' Refresh_Graph Macro
'

' Unfilter all data
    activeSheet.Range("$I$8:$K$33").AutoFilter Field:=1
    activeSheet.Range("$I$8:$K$33").AutoFilter Field:=2
    activeSheet.Range("$I$8:$K$33").AutoFilter Field:=3
    
'Set Total for the bars
    activeSheet.ChartObjects("Chart 2").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    ActiveChart.FullSeriesCollection(1).Points(2).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(3).Select
    ActiveChart.FullSeriesCollection(1).Points(3).IsTotal = False
    
    ActiveChart.FullSeriesCollection(1).Points(9).Select
    ActiveChart.FullSeriesCollection(1).Points(9).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(10).Select
    ActiveChart.FullSeriesCollection(1).Points(10).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(11).Select
    ActiveChart.FullSeriesCollection(1).Points(11).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(12).Select
    ActiveChart.FullSeriesCollection(1).Points(12).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(13).Select
    ActiveChart.FullSeriesCollection(1).Points(13).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(14).Select
    ActiveChart.FullSeriesCollection(1).Points(14).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(15).Select
    ActiveChart.FullSeriesCollection(1).Points(15).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(16).Select
    ActiveChart.FullSeriesCollection(1).Points(16).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(17).Select
    ActiveChart.FullSeriesCollection(1).Points(17).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(18).Select
    ActiveChart.FullSeriesCollection(1).Points(18).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(19).Select
    ActiveChart.FullSeriesCollection(1).Points(19).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(20).Select
    ActiveChart.FullSeriesCollection(1).Points(20).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(21).Select
    ActiveChart.FullSeriesCollection(1).Points(21).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(22).Select
    ActiveChart.FullSeriesCollection(1).Points(22).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(23).Select
    ActiveChart.FullSeriesCollection(1).Points(23).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(24).Select
    ActiveChart.FullSeriesCollection(1).Points(24).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(25).Select
    ActiveChart.FullSeriesCollection(1).Points(25).IsTotal = False
    ActiveChart.FullSeriesCollection(1).Points(26).Select
    ActiveChart.FullSeriesCollection(1).Points(26).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(27).Select
    ActiveChart.FullSeriesCollection(1).Points(27).IsTotal = True
    ActiveChart.FullSeriesCollection(1).Points(28).Select
    ActiveChart.FullSeriesCollection(1).Points(28).IsTotal = True
    
'Show non zero data
    activeSheet.Range("$I$8:$I$33").AutoFilter Field:=1, Criteria1:="Show"
    



    
End Sub
Sub Refresh_Table()
'
' Refresh_Table Macro
'

'

    activeSheet.Range("$AO$11:$AT$53").AutoFilter Field:=1, Criteria1:="Show"
End Sub




