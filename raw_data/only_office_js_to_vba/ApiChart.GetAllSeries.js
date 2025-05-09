### Description / Описание

This code creates a financial overview chart by populating data into an Excel worksheet, adding a combination bar and line chart, customizing series colors, and recording each series' chart type.

Этот код создает диаграмму финансового обзора, заполняя данные в рабочий лист Excel, добавляя комбинированную диаграмму столбцов и линий, настраивая цвета серий и записывая тип каждой серии диаграммы.

```vba
' This VBA code creates a financial overview chart by populating data, adding a combination chart,
' customizing series colors, and recording each series' chart type.
' Этот VBA код создает диаграмму финансового обзора, заполняя данные, добавляя комбинированную диаграмму,
' настраивая цвета серий и записывая тип каждой серии диаграммы.

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series
    Dim nSeries As Integer
    Dim sSeriesType As String
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set header values
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a combo bar and line chart
    Set oChart = oWorksheet.Shapes.AddChart2(240, xlColumnClustered, 200, 100, 360, 270).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    
    ' Set chart type to combination (Bar and Line)
    oChart.ChartType = xlColumnClustered
    oChart.SeriesCollection.NewSeries
    oChart.SeriesCollection(2).ChartType = xlLine
    
    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Set series fill colors
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark gray
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    
    ' Iterate through all series and write their type to column F
    For nSeries = 1 To oChart.SeriesCollection.Count
        Set oSeries = oChart.SeriesCollection(nSeries)
        sSeriesType = GetChartTypeName(oSeries.ChartType)
        oWorksheet.Range("F" & nSeries).Value = nSeries & " Series Type = " & sSeriesType
    Next nSeries
End Sub

' Function to convert ChartType enum to string
Function GetChartTypeName(chartType As XlChartType) As String
    Select Case chartType
        Case xlColumnClustered
            GetChartTypeName = "Clustered Column"
        Case xlLine
            GetChartTypeName = "Line"
        ' Add other cases as needed
        Case Else
            GetChartTypeName = "Other"
    End Select
End Function
```

```javascript
// This example gets all series of ApiChart class and inserts their types into the table.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
var aAllSeries = oChart.GetAllSeries();
var oSeries, sSeriesType;
for(var nSeries = 0; nSeries < aAllSeries.length; ++nSeries) {
    oSeries = aAllSeries[nSeries];
    sSeriesType = oSeries.GetChartType();
    oWorksheet.GetRange("F" + (nSeries + 1)).SetValue((nSeries + 1) + " Series Type = " + sSeriesType);
} 
```