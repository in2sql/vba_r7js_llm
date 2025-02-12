## Description / Описание

**English:** This code populates a worksheet with financial data, creates a scatter chart with customized markers and outlines, and sets the chart title and axis tick marks.

**Russian:** Этот код заполняет лист с финансовыми данными, создает точечную диаграмму с настраиваемыми маркерами и контурами, а также устанавливает заголовок диаграммы и метки основных делений оси.

```javascript
// This example specifies the major tick mark for the vertical axis.
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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);
oChart.SetVertAxisMajorTickMark("cross");
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```

```vba
' This VBA code populates a worksheet with financial data, creates a scatter chart with customized markers and outlines,
' and sets the chart title and axis tick marks.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim series1 As Series
    Dim series2 As Series
    Dim fill1 As FillFormat
    Dim fill2 As FillFormat
    Dim line1 As LineFormat
    Dim line2 As LineFormat
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate the worksheet with data
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a scatter chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=500)
    Set chart = chartObj.Chart
    chart.ChartType = xlXYScatter
    
    ' Set the data range for the chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 13
    
    ' Set major tick mark for vertical axis
    chart.Axes(xlValue).MajorTickMark = xlTickMarkCross
    
    ' Customize first series markers
    Set series1 = chart.SeriesCollection(1)
    Set fill1 = series1.MarkerBackgroundColor = RGB(51, 51, 51)
    series1.MarkerStyle = xlMarkerStyleCircle
    series1.MarkerSize = 7
    series1.Format.Line.Weight = 0.5
    series1.Format.Line.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Customize second series markers
    Set series2 = chart.SeriesCollection(2)
    Set fill2 = series2.MarkerBackgroundColor = RGB(255, 111, 61)
    series2.MarkerStyle = xlMarkerStyleCircle
    series2.MarkerSize = 7
    series2.Format.Line.Weight = 0.5
    series2.Format.Line.ForeColor.RGB = RGB(255, 111, 61)
End Sub
```