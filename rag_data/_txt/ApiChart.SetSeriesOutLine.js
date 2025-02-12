### Description / Описание
**English:**  
This script populates a worksheet with financial data for the years 2014 to 2016, creates a 3D bar chart titled "Financial Overview", and customizes the series fills and outlines.

**Russian:**  
Этот скрипт заполняет лист финансовыми данными за 2014–2016 годы, создает 3D столбчатую диаграмму с заголовком "Обзор финансов" и настраивает заливку и контуры серий.

```vba
' Excel VBA Equivalent Code

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim fillColor1 As Long
    Dim fillColor2 As Long
    Dim strokeColor As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate header values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Populate row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Populate data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=36000, Top:=100, Height:=70 * 360)
    Set chart = chartObj.Chart
    chart.ChartType = xlBarClustered ' Using xlBarClustered as VBA may not have a direct 3D equivalent
    
    ' Set the data range for the chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    
    ' Set series fill colors
    fillColor1 = RGB(51, 51, 51)
    fillColor2 = RGB(255, 111, 61)
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor1
    chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor2
    
    ' Set series outlines
    strokeColor = RGB(51, 51, 51)
    With chart.SeriesCollection(2).Format.Line
        .Visible = msoTrue
        .Weight = 0.5
        .ForeColor.RGB = strokeColor
    End With
End Sub
```

```javascript
// OnlyOffice JS Equivalent Code

// This script populates a worksheet with financial data, creates a 3D bar chart, and customizes its appearance.

function createFinancialChart() {
    var oWorksheet = Api.GetActiveSheet();
    
    // Populate header values
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);
    
    // Populate row labels
    oWorksheet.GetRange("A2").SetValue("Projected Revenue");
    oWorksheet.GetRange("A3").SetValue("Estimated Costs");
    
    // Populate data values
    oWorksheet.GetRange("B2").SetValue(200);
    oWorksheet.GetRange("B3").SetValue(250);
    oWorksheet.GetRange("C2").SetValue(240);
    oWorksheet.GetRange("C3").SetValue(260);
    oWorksheet.GetRange("D2").SetValue(280);
    oWorksheet.GetRange("D3").SetValue(280);
    
    // Add a 3D bar chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 36000 * 100, 36000 * 70, 0, 36000 * 2, 7, 36000 * 3);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Set series fill colors
    var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill1, 0, false);
    
    var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill2, 1, false);
    
    // Set series outlines
    var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
    oChart.SetSeriesOutLine(oStroke, 1, false);
}
```