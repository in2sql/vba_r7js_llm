### Financial Overview Chart Creation / Создание диаграммы Финансовый обзор

This script populates an Excel worksheet with financial data and creates a 3D bar chart titled "Financial Overview." It sets specific fill colors for the data series and outlines the chart legend.
Этот скрипт заполняет рабочий лист Excel финансовыми данными и создает 3D столбчатую диаграмму с названием "Финансовый обзор". Устанавливает определенные цвета заливки для серий данных и обводку легенды диаграммы.

```javascript
// JavaScript (OnlyOffice) Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values for the header
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart(
    "'Sheet1'!$A$1:$D$3", 
    true, 
    "bar3D", 
    2, 
    100 * 36000, 
    70 * 36000, 
    0, 
    2 * 36000, 
    7, 
    3 * 36000
);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set the fill color for the first data series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second data series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create and set the stroke for the legend outline
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetLegendOutLine(oStroke);
```

```vba
' VBA Code Equivalent

Sub CreateFinancialOverviewChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values for the header
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    ' Set row labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a 3D bar chart to the worksheet
    Dim oChart As Chart
    Set oChart = oWorksheet.Shapes.AddChart2(251, xlBarClustered, 100, 70, 300, 200).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    
    ' Set the chart type to 3D bar
    oChart.ChartType = xlBarClustered
    oChart.ChartStyle = 251 ' Custom style number
    
    ' Set the chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Set fill color for the first series
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Set fill color for the second series
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Set legend outline
    With oChart.Legend.Format.Line
        .Visible = msoTrue
        .Weight = 0.5
        .ForeColor.RGB = RGB(51, 51, 51)
    End With
End Sub
```