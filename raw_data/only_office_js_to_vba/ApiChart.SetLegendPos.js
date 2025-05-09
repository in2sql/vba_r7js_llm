**Description:**
This code populates a worksheet with financial data for 2014-2016, creates a 3D bar chart, sets the title and legend position, and formats the series fills.
Этот код заполняет лист финансовыми данными за 2014-2016 годы, создает 3D столбчатую диаграмму, устанавливает заголовок и положение легенды, а также форматирует заливку серий.

**OnlyOffice JS Code:**
```javascript
// Populate worksheet with data
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

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set legend position to the right
oChart.SetLegendPos("right");

// Set fill color for series 0
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Set fill color for series 1
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

**VBA Code:**
```vba
' Populate worksheet with data
Sub CreateFinancialChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=300, Height:=200)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xl3DBarClustered
        
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set legend position to the right
        .Legend.Position = xlLegendPositionRight
        
        ' Set fill color for series 1
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set fill color for series 2
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```