Attribute VB_Name = "Module1"
Sub Blue()

'
' Blue Macro
'

'
    Dim intCount As Integer ' Declares a variable
    For intCount = 1 To 96 ' This will iterate the command 96 times
    ActiveChart.FullSeriesCollection(intCount).Select ' This defines which series you want to color
    With Selection.Format.Fill ' Fills in the line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240) ' color code for light blue
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line ' Fillins in the border of the line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240)
        .Transparency = 0
    End With
    Next intCount ' Will go to the next series to color
    ActiveChart.ChartArea.Select
End Sub
Sub experi1()
Attribute experi1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' experi1 Macro
'

'
    Range("B39:M68").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$B$39:$M$68")
    ActiveChart.SetSourceData
    ActiveWindow.SmallScroll Down:=-3
End Sub
Sub experi3()
Attribute experi3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' experi3 Macro
'

'
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$B$39:$M$68")
End Sub
Sub experi2()
Attribute experi2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' experi2 Macro
'

'
    ActiveChart.ChartType = xlColumnClustered
End Sub
Sub experi4()
Attribute experi4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' experi4 Macro
'

'
    ActiveSheet.ChartObjects("Chart 5").Activate
    ActiveChart.SetSourceData
End Sub
Sub AddCallout()
Attribute AddCallout.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AddCallout Macro
'

'
    ActiveChart.FullSeriesCollection(95).Select
    ActiveChart.FullSeriesCollection(95).Points(14).Select
    ActiveChart.SetElement (msoElementDataLabelCallout)
    ActiveChart.FullSeriesCollection(95).DataLabels.Select
    Selection.Format.AutoShapeType = msoShapeDownArrowCallout
    ActiveChart.FullSeriesCollection(95).Points(14).DataLabel.Select
    Selection.Left = 408.936
    Selection.Top = 89.123
    ActiveChart.ChartArea.Select
End Sub
