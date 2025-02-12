**English:** This code populates an Excel sheet with financial data for the years 2014-2016, creates a 3D bar chart titled "Financial Overview" with the horizontal axis labeled "Year", and sets specific fill colors for each data series.

**Russian:** Этот код заполняет лист Excel финансовыми данными за 2014-2016 годы, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор", названной горизонтальной осью "Год", а также устанавливает определенные цвета заливки для каждой серии данных.

```vba
' VBA Code Equivalent

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Populate the cells with data
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Set oChart = oWorksheet.Shapes.AddChart2(251, xlBarClustered, 200, 100, 360, 270).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    
    ' Set the chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Set the horizontal axis title
    oChart.Axes(xlCategory, xlPrimary).HasTitle = True
    oChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Year"
    oChart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 11
    
    ' Set fill color for the first series
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Set fill color for the second series
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This script populates the active sheet with financial data, creates a 3D bar chart titled "Financial Overview",
// sets the horizontal axis title to "Year", and defines specific colors for each data series.

// Populating the worksheet with data
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

// Adding a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Setting the chart title
oChart.SetTitle("Financial Overview", 13);

// Setting the horizontal axis title
oChart.SetHorAxisTitle("Year", 11);

// Setting fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Setting fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```