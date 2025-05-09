Attribute VB_Name = "Module10"
'Courtesy of Excel Tips https://excel.tips.net/T002429_Specifying_Chart_Sizes.html
Sub StandardSizeForChart() ' Makes every chart in a sheeet a standard size
    Dim cht As ChartObject
    For Each cht In ActiveSheet.ChartObjects
        cht.Chart.ChartArea.AutoScaleFont = False
        cht.Height = Application.CentimetersToPoints(14) ' 14 cm
        cht.Width = Application.CentimetersToPoints(33) ' 33 cm
    Next cht
End Sub
