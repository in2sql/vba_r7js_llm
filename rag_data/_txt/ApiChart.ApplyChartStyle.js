**Description / Описание:**

This code sets specific values in cells, adds a 3D bar chart titled "Financial Overview", applies a chart style, and sets the fill and outline for each series with specific colors.

Этот код устанавливает определенные значения в ячейки, добавляет 3D столбчатую диаграмму с заголовком "Финансовый обзор", применяет стиль диаграммы и задает заливку и обводку для каждой серии с определенными цветами.

```vba
' VBA Code
Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim fillColor1 As Long
    Dim fillColor2 As Long
    Dim strokeWeight As Single
    
    ' Set the active worksheet
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
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
    Set chart = chartObj.Chart
    chart.ChartType = xlBar3DClustered
    
    ' Set the data range for the chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    
    ' Set chart title
    With chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        ' Apply chart style
        .ApplyChartStyle 2
    End With
    
    ' Set fill and outline for series
    fillColor1 = RGB(51, 51, 51) ' Dark gray
    fillColor2 = RGB(255, 111, 61) ' Orange
    strokeWeight = 0.5
    
    With chart.SeriesCollection(1).Format
        .Fill.ForeColor.RGB = fillColor1
        .Line.ForeColor.RGB = fillColor1
        .Line.Weight = strokeWeight
    End With
    
    With chart.SeriesCollection(2).Format
        .Fill.ForeColor.RGB = fillColor2
        .Line.ForeColor.RGB = fillColor2
        .Line.Weight = strokeWeight
    End With
End Sub
```

```javascript
// JavaScript Code for OnlyOffice API
// This code sets specific cell values, adds a 3D bar chart titled "Financial Overview",
// applies a chart style, and sets the fill and outline for each series with specific colors.

function createFinancialChart() {
    // Get the active worksheet
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
    
    // Add a 3D bar chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D",
        2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Apply chart style
    oChart.ApplyChartStyle(2);
    
    // Set fill and outline for series
    var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill1, 0, false);
    
    var oStroke1 = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
    oChart.SetSeriesOutLine(oStroke1, 0, false);
    
    var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill2, 1, false);
    
    var oStroke2 = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
    oChart.SetSeriesOutLine(oStroke2, 1, false);
}
```