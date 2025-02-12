**Description / Описание**

English: This code creates a 3D bar chart with specified data, sets the title and legend position, and applies fill colors to the series.

Russian: Этот код создает 3D-гистограмму с заданными данными, устанавливает заголовок и положение легенды, а также применяет цвета заливки к сериям.

```vba
' VBA code
' This VBA code creates a 3D bar chart with specified data, sets the title and legend position, and applies fill colors to the series.

Sub CreateChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
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
    
    ' Add chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=360, Top:=70, Height:=200)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xl3DColumnClustered ' Sets a 3D clustered column chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        .Legend.Position = xlLegendPositionRight
        
        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark gray
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    End With
End Sub
```

```javascript
// JS code
// This JavaScript code creates a 3D bar chart with specified data, sets the title and legend position, and applies fill colors to the series.

var oWorksheet = Api.GetActiveSheet();

// Set cell values
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

// Add chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 100, 70, 360, 200, 0, 2);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set legend position
oChart.SetLegendPos("right");

// Set series fill colors
var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark gray
oChart.SetSeriesFill(oFill1, 0, false);
var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
oChart.SetSeriesFill(oFill2, 1, false);
```