**Description / Описание:**

*English*: This code sets specific values in cells, adds a 3D bar chart, sets the chart title, formats the axis numbers, and assigns fill colors to the chart series.

*Russian*: Этот код устанавливает конкретные значения в ячейки, добавляет 3D столбчатую диаграмму, устанавливает заголовок диаграммы, форматирует числа оси и назначает цвета заливки для рядов диаграммы.

```javascript
// JavaScript (OnlyOffice API) code

// This script sets values in specific cells, creates a 3D bar chart, sets the title and axis number format, and applies fill colors to the chart series.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("B1").SetValue(2014); // Set cell B1 to 2014
oWorksheet.GetRange("C1").SetValue(2015); // Set cell C1 to 2015
oWorksheet.GetRange("D1").SetValue(2016); // Set cell D1 to 2016
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set cell A2
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set cell A3
oWorksheet.GetRange("B2").SetValue(200); // Set cell B2
oWorksheet.GetRange("B3").SetValue(250); // Set cell B3
oWorksheet.GetRange("C2").SetValue(240); // Set cell C2
oWorksheet.GetRange("C3").SetValue(260); // Set cell C3
oWorksheet.GetRange("D2").SetValue(280); // Set cell D2
oWorksheet.GetRange("D3").SetValue(280); // Set cell D3

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13); // Set chart title with font size
oChart.SetAxieNumFormat("0.00", "left"); // Set axis number format

// Set fill color for first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Set fill color for second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' VBA code equivalent to the OnlyOffice JS code

Sub CreateFinancialChart()
    ' This macro sets values in specific cells, creates a 3D bar chart, sets the title and axis number format, and applies fill colors to the chart series.
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Set cell values
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
    
    ' Add a 3D bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=260)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xlBar3DClustered
        
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Format axis numbers
        .Axes(xlCategory).TickLabels.NumberFormat = "0.00"
        
        ' Set fill color for first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set fill color for second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```