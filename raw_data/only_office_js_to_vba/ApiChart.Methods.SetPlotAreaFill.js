**Description**

English: This example populates a worksheet with financial data, creates a 3D bar chart, sets the chart title, and applies specific fill colors to the chart series and plot area.

Russian: Этот пример заполняет рабочий лист финансовыми данными, создает 3D-гистограмму, устанавливает заголовок диаграммы и применяет определенные цвета заливки к сериям диаграммы и области построения.

**OnlyOffice JavaScript Code**

```javascript
// Populate the worksheet with financial data
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("B1").SetValue(2014); // Set year 2014
oWorksheet.GetRange("C1").SetValue(2015); // Set year 2015
oWorksheet.GetRange("D1").SetValue(2016); // Set year 2016
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set label for revenue
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set label for costs
oWorksheet.GetRange("B2").SetValue(200); // Set projected revenue for 2014
oWorksheet.GetRange("B3").SetValue(250); // Set estimated costs for 2014
oWorksheet.GetRange("C2").SetValue(240); // Set projected revenue for 2015
oWorksheet.GetRange("C3").SetValue(260); // Set estimated costs for 2015
oWorksheet.GetRange("D2").SetValue(280); // Set projected revenue for 2016
oWorksheet.GetRange("D3").SetValue(280); // Set estimated costs for 2016

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13); // Set chart title with font size 13

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create and set fill color for the plot area
oFill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128));
oChart.SetPlotAreaFill(oFill);
```

**Excel VBA Code**

```vba
' Populate the worksheet with financial data and create a 3D bar chart
Sub CreateFinancialChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set years in header
    ws.Range("B1").Value = 2014 ' Set year 2014
    ws.Range("C1").Value = 2015 ' Set year 2015
    ws.Range("D1").Value = 2016 ' Set year 2016
    
    ' Set labels for revenue and costs
    ws.Range("A2").Value = "Projected Revenue" ' Set label for revenue
    ws.Range("A3").Value = "Estimated Costs" ' Set label for costs
    
    ' Enter financial data
    ws.Range("B2").Value = 200 ' Set projected revenue for 2014
    ws.Range("B3").Value = 250 ' Set estimated costs for 2014
    ws.Range("C2").Value = 240 ' Set projected revenue for 2015
    ws.Range("C3").Value = 260 ' Set estimated costs for 2015
    ws.Range("D2").Value = 280 ' Set projected revenue for 2016
    ws.Range("D3").Value = 280 ' Set estimated costs for 2016
    
    ' Add a 3D bar chart to the worksheet
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=400, Height:=300)
    With chartObj.Chart
        .ChartType = xl3DBarClustered ' Set chart type to 3D clustered bar
        .SetSourceData Source:=ws.Range("A1:D3") ' Set the data range for the chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview" ' Set chart title
        .ChartTitle.Font.Size = 13 ' Set chart title font size
        
        ' Set fill color for the first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set fill color for the second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
        
        ' Set fill color for the plot area
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(128, 128, 128)
    End With
End Sub
```