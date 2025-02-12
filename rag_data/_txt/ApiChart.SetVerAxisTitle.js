### Description / Описание

**English:**  
This code populates an Excel sheet with revenue and cost data from 2014 to 2016 and creates a 3D bar chart titled "Financial Overview" with customized series colors using both OnlyOffice API and Excel VBA.

**Russian:**  
Этот код заполняет лист Excel данными о доходах и расходах с 2014 по 2016 годы и создает 3D столбчатую диаграмму с заголовком "Financial Overview" и настроенными цветами серий как с использованием OnlyOffice API, так и Excel VBA.

```javascript
// OnlyOffice JS Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values for years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row titles
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

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set the vertical axis title
oChart.SetVerAxisTitle("USD In Hundred Thousands", 10);

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' Excel VBA Equivalent Code

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim oFill As Object
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Set header values for years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set row titles
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
    
    ' Add a 3D bar chart to the worksheet
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=400, Top:=100, Height:=300)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3")
        .ChartType = xlBarClustered ' Excel VBA does not have a direct 3D bar type equivalent
        ' Set the chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set the vertical axis title
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "USD In Hundred Thousands"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        
        ' Set the fill color for the first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set the fill color for the second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```