## Description / Описание

**English:**  
This script populates an Excel worksheet with financial data for the years 2014 to 2016, creates a scatter chart titled "Financial Overview" with customized markers and outlines, and sets the major tick mark style for the horizontal axis to "cross".

**Russian:**  
Этот скрипт заполняет рабочий лист Excel финансовыми данными за годы 2014–2016, создаёт точечную диаграмму с названием "Финансовый обзор" с настраиваемыми маркерами и контурами, а также устанавливает стиль основных делений по горизонтальной оси на "крест".

---

## OnlyOffice JS Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row titles
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set financial data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a scatter chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set the major tick mark style for the horizontal axis to "cross"
oChart.SetHorAxisMajorTickMark("cross");

// Create and set marker fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create and set marker outline for the first series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Create and set marker outline for the second series
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```

---

## Excel VBA Code

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set header years
oWorksheet.Range("B1").Value = 2014
oWorksheet.Range("C1").Value = 2015
oWorksheet.Range("D1").Value = 2016

' Set row titles
oWorksheet.Range("A2").Value = "Projected Revenue"
oWorksheet.Range("A3").Value = "Estimated Costs"

' Set financial data
oWorksheet.Range("B2").Value = 200
oWorksheet.Range("B3").Value = 250
oWorksheet.Range("C2").Value = 240
oWorksheet.Range("C3").Value = 260
oWorksheet.Range("D2").Value = 280
oWorksheet.Range("D3").Value = 280

' Add a scatter chart to the worksheet
Dim oChart As Chart
Set oChart = oWorksheet.Shapes.AddChart2(251, xlXYScatterLines).Chart

' Set the data range for the chart
oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")

' Set the chart title
oChart.HasTitle = True
oChart.ChartTitle.Text = "Financial Overview"
oChart.ChartTitle.Font.Size = 13

' Set the major tick mark style for the horizontal axis to "cross"
oChart.Axes(xlCategory).MajorTickMark = xlTickMarkCross

' Customize marker fill for the first series
With oChart.SeriesCollection(1).Format.Fill
    .Visible = msoTrue
    .Solid
    .ForeColor.RGB = RGB(51, 51, 51)
End With

' Customize marker outline for the first series
With oChart.SeriesCollection(1).Format.Line
    .Visible = msoTrue
    .Weight = 0.5
    .ForeColor.RGB = RGB(51, 51, 51)
End With

' Customize marker fill for the second series
With oChart.SeriesCollection(2).Format.Fill
    .Visible = msoTrue
    .Solid
    .ForeColor.RGB = RGB(255, 111, 61)
End With

' Customize marker outline for the second series
With oChart.SeriesCollection(2).Format.Line
    .Visible = msoTrue
    .Weight = 0.5
    .ForeColor.RGB = RGB(255, 111, 61)
End With
```