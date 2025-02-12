**Description:**
This code sets specific values in an Excel worksheet and creates a 3D bar chart to visualize financial data.
Этот код устанавливает определенные значения в рабочем листе Excel и создает 3D столбчатую диаграмму для визуализации финансовых данных.

```vba
' VBA code to set cell values and create a 3D bar chart

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range

    ' Set reference to the active sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set values in cells
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

    ' Define the range for the chart
    Set chartRange = ws.Range("A1:D3")

    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=3600, Top:=700, Height:=3600)
    With chartObj.Chart
        .ChartType = xlBar3DClustered ' Set chart type to 3D Bar
        .SetSourceData Source:=chartRange ' Set the data range
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview" ' Set chart title
        .Axes(xlValue).NumberFormat = "0.00" ' Set axis number format

        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' First series color
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Second series color
    End With
End Sub
```

```javascript
// JavaScript code for OnlyOffice to set cell values and create a 3D bar chart

var oWorksheet = Api.GetActiveSheet();

// Set values in cells
oWorksheet.GetRange("B1").SetValue(2014); // Set year 2014
oWorksheet.GetRange("C1").SetValue(2015); // Set year 2015
oWorksheet.GetRange("D1").SetValue(2016); // Set year 2016
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set label for revenue
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set label for costs
oWorksheet.GetRange("B2").SetValue(200); // Revenue for 2014
oWorksheet.GetRange("B3").SetValue(250); // Costs for 2014
oWorksheet.GetRange("C2").SetValue(240); // Revenue for 2015
oWorksheet.GetRange("C3").SetValue(260); // Costs for 2015
oWorksheet.GetRange("D2").SetValue(280); // Revenue for 2016
oWorksheet.GetRange("D3").SetValue(280); // Costs for 2016

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13); // Set chart title
oChart.SetAxieNumFormat("0.00", "left"); // Set axis number format

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false); // First series color
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false); // Second series color
```