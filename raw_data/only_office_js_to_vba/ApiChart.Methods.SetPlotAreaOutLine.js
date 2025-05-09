### Code Description / Описание кода

**English:**  
This code sets up financial data in specific cells, creates a 3D bar chart titled "Financial Overview", applies specific fill colors to the chart series, and sets the outline for the plot area.

**Russian:**  
Данный код устанавливает финансовые данные в определенные ячейки, создает 3D гистограмму с заголовком «Финансовый обзор», применяет определенные цвета заливки к рядам графика и задает контур области построения.

---

#### OnlyOffice JS Code

```javascript
// This example sets the outline to the chart plot area.
var oWorksheet = Api.GetActiveSheet();

// Set years in the header
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels for data series
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values for Projected Revenue
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("D2").SetValue(280);

// Set data values for Estimated Costs
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart with specified parameters
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the title of the chart
oChart.SetTitle("Financial Overview", 13);

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create and set the outline stroke for the plot area
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetPlotAreaOutLine(oStroke);
```

---

#### Excel VBA Code

```vba
' This VBA code sets up financial data, creates a 3D bar chart, applies fill colors, and sets the plot area outline.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set years in the header
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set labels for data series
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values for Projected Revenue
    ws.Range("B2").Value = 200
    ws.Range("C2").Value = 240
    ws.Range("D2").Value = 280
    
    ' Set data values for Estimated Costs
    ws.Range("B3").Value = 250
    ws.Range("C3").Value = 260
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xl3DBarClustered
    
    ' Set the title of the chart
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    chart.ChartTitle.Font.Size = 13
    
    ' Apply fill color to the first series
    With chart.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 51, 51)
        .Solid
    End With
    
    ' Apply fill color to the second series
    With chart.SeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    
    ' Set the outline for the plot area
    With chart.PlotArea.Format.Line
        .Visible = msoTrue
        .Weight = 0.5
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```