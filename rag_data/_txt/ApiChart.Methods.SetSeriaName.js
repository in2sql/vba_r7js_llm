```rag
# Description / Описание
This code sets specific values in an Excel worksheet and creates a 3D bar chart titled "Financial Overview" with customized series names and colors.
Этот код устанавливает определенные значения в рабочем листе Excel и создает 3D столбчатую диаграмму с заголовком "Финансовый обзор" с настраиваемыми названиями серий и цветами.
```

```vba
' VBA Code Equivalent
Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oFill As Object

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set values in cells
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
    Set oChart = oWorksheet.Shapes.AddChart2(251, xlBarClustered, 100, 70, 400, 300).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    oChart.ChartType = xl3DBarClustered

    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"

    ' Set series names
    oChart.SeriesCollection(1).Name = "Projected Sales"

    ' Set series fill colors
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent
// This example sets specific values in the worksheet and creates a 3D bar chart with customized series properties.
var oWorksheet = Api.GetActiveSheet();

// Set values in cells
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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series names
oChart.SetSeriaName("Projected Sales", 0);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```