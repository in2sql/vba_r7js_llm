```javascript
// This script sets up data and creates a 3D bar chart titled 'Financial Overview' with customized series colors and outlines.
// Этот скрипт настраивает данные и создает 3D столбчатую диаграмму с заголовком 'Financial Overview' с настраиваемыми цветами серий и контуром.

// JavaScript (OnlyOffice) Code
var oWorksheet = Api.GetActiveSheet();

// Set years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row headers
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add 3D Bar Chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Customize series fills
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Customize data point outline
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetDataPointOutLine(oStroke, 1, 0, false);
```

```vba
' This script sets up data and creates a 3D bar chart titled 'Financial Overview' with customized series colors and outlines.
' Этот скрипт настраивает данные и создает 3D столбчатую диаграмму с заголовком 'Financial Overview' с настраиваемыми цветами серий и контуром.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Set years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016

    ' Set row headers
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"

    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280

    ' Add 3D Bar Chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
    With chartObj.Chart
        .ChartType = xl3DColumn ' Set chart type to 3D Column
        .SetSourceData Source:=ws.Range("A1:D3") ' Set the data range for the chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview" ' Set chart title
        
        ' Customize series colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Set first series color
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Set second series color
        
        ' Customize data point outline
        .SeriesCollection(2).Format.Line.Weight = 0.5 ' Set outline weight
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(51, 51, 51) ' Set outline color
    End With
End Sub
```