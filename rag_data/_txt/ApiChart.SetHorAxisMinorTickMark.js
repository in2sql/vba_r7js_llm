**Description / Описание**

English:
This script populates an active sheet with financial data, creates a scatter chart titled "Financial Overview", configures minor tick marks on the horizontal axis, and customizes the marker fills and outlines for different data series.

Русский:
Этот скрипт заполняет активный лист финансовыми данными, создает точечную диаграмму с заголовком "Финансовый обзор", настраивает дополнительные деления на горизонтальной оси и настраивает заливку и контуры маркеров для различных серий данных.

```javascript
// JavaScript Code for OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells B1, C1, D1
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels in A2 and A3
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data for Projected Revenue
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("D2").SetValue(280);

// Set data for Estimated Costs
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D3").SetValue(280);

// Add a scatter chart with specified range and positioning
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set minor tick marks on the horizontal axis
oChart.SetHorAxisMinorTickMark("out");

// Create and set marker fill for the first data series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create and set marker outline for the first data series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill for the second data series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Create and set marker outline for the second data series
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```

```vba
' VBA Code Equivalent for Excel

Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in cells B1, C1, D1
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    ' Set labels in A2 and A3
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Set data for Projected Revenue
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("D2").Value = 280
    
    ' Set data for Estimated Costs
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D3").Value = 280
    
    ' Add a scatter chart with specified range
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Top:=70, Width:=400, Height:=300)
    oChart.Chart.ChartType = xlXYScatter
    
    ' Set the data source for the chart
    oChart.Chart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    
    ' Set the chart title
    oChart.Chart.HasTitle = True
    oChart.Chart.ChartTitle.Text = "Financial Overview"
    oChart.Chart.ChartTitle.Font.Size = 13
    
    ' Set minor tick marks on the horizontal axis
    With oChart.Chart.Axes(xlCategory)
        .HasMinorGridlines = True
        .MinorTickMark = xlTickMarkOutside
    End With
    
    ' Customize marker for first data series
    With oChart.Chart.SeriesCollection(1)
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 7
        .Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        .Format.Line.Weight = 0.5
        .Format.Line.ForeColor.RGB = RGB(51, 51, 51)
    End With
    
    ' Customize marker for second data series
    With oChart.Chart.SeriesCollection(2)
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 7
        .Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
        .Format.Line.Weight = 0.5
        .Format.Line.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```