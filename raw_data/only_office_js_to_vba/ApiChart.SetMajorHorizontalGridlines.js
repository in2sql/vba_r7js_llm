**Description:**

*English:* This code sets up a financial overview worksheet by populating data for projected revenue and estimated costs from 2014 to 2016 and creates a 3D bar chart visualizing the data.

*Russian:* Этот код настраивает лист финансового обзора, заполняя данные о прогнозируемом доходе и оцененных расходах за 2014–2016 годы и создает 3D столбчатую диаграмму для визуализации данных.

---

**VBA Code:**

```vba
' Create a financial overview with projected revenue and estimated costs and add a 3D bar chart

Sub CreateFinancialOverview()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series
    Dim oFill As Object
    Dim oStroke As Object

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set header values
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016

    ' Set labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"

    ' Set data values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280

    ' Add a 3D bar chart
    Set oChart = oWorksheet.Shapes.AddChart2(240, xlBarClustered, 100, 100, 360, 270).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    oChart.ChartType = xl3DBarClustered

    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"

    ' Set series fill colors
    Set oSeries = oChart.SeriesCollection(1)
    oSeries.Format.Fill.ForeColor.RGB = RGB(51, 51, 51)

    Set oSeries = oChart.SeriesCollection(2)
    oSeries.Format.Fill.ForeColor.RGB = RGB(255, 111, 61)

    ' Set major horizontal gridlines
    With oChart.Axes(xlCategory).MajorGridlines.Format.Line
        .Weight = 1
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```

---

**OnlyOffice JS Code:**

```javascript
// Create a financial overview with projected revenue and estimated costs and add a 3D bar chart

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
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

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Set major horizontal gridlines
var oStroke = Api.CreateStroke(1 * 15000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMajorHorizontalGridlines(oStroke);
```