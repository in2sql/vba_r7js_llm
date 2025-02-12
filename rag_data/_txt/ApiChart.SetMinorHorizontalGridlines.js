# Description / Описание

This script populates specific cells with data, creates a 3D bar chart, sets the chart title, customizes the series fill colors, and configures the minor horizontal gridlines.  
Этот скрипт заполняет определенные ячейки данными, создает 3D гистограмму, устанавливает заголовок графика, настраивает цвета заливки серий и конфигурирует мелкие горизонтальные сетки.

```javascript
// Populate cells with data
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

// Create a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Customize the first series fill color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Customize the second series fill color
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Set the minor horizontal gridlines
var oStroke = Api.CreateStroke(1 * 5000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMinorHorizontalGridlines(oStroke);
```

```vba
' Populate cells with data
Sub CreateFinancialChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
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
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 200, 100, 420, 300).Chart
    cht.SetSourceData Source:=ws.Range("A1:D3")
    cht.ChartType = xl3DBarClustered
    
    ' Set the chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13
    
    ' Customize the first series fill color
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Customize the second series fill color
    cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Set the minor horizontal gridlines
    With cht.Axes(xlValue)
        .HasMinorGridlines = True
        .MinorGridlines.Format.Line.Weight = 1
        .MinorGridlines.Format.Line.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```