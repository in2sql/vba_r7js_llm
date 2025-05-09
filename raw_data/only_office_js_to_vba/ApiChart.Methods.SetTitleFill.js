# Description / Описание

**English:** This code populates specific cells with data, adds a 3D bar chart to the worksheet, sets the chart title, and applies specific fill colors to chart series and the title.

**Russian:** Этот код заполняет определенные ячейки данными, добавляет 3D столбчатую диаграмму на лист, устанавливает заголовок диаграммы и применяет определенные цвета заливки к сериям диаграммы и заголовку.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells B1, C1, D1
oWorksheet.GetRange("B1").SetValue(2014); // Set cell B1 to 2014
oWorksheet.GetRange("C1").SetValue(2015); // Set cell C1 to 2015
oWorksheet.GetRange("D1").SetValue(2016); // Set cell D1 to 2016

// Set headers in A2 and A3
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set cell A2 to "Projected Revenue"
oWorksheet.GetRange("A3").SetValue("Estimated Costs");    // Set cell A3 to "Estimated Costs"

// Set data values
oWorksheet.GetRange("B2").SetValue(200); // Set cell B2 to 200
oWorksheet.GetRange("B3").SetValue(250); // Set cell B3 to 250
oWorksheet.GetRange("C2").SetValue(240); // Set cell C2 to 240
oWorksheet.GetRange("C3").SetValue(260); // Set cell C3 to 260
oWorksheet.GetRange("D2").SetValue(280); // Set cell D2 to 280
oWorksheet.GetRange("D3").SetValue(280); // Set cell D3 to 280

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the chart title with font size 13
oChart.SetTitle("Financial Overview", 13);

// Create and set fill for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Create a dark gray fill
oChart.SetSeriesFill(oFill, 0, false); // Apply fill to series 0

// Create and set fill for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create an orange fill
oChart.SetSeriesFill(oFill, 1, false); // Apply fill to series 1

// Create and set fill for the chart title
oFill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128)); // Create a gray fill
oChart.SetTitleFill(oFill); // Apply fill to the chart title
```

```vba
' VBA code equivalent to the OnlyOffice JavaScript example

Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set values in cells B1, C1, D1
    oWorksheet.Range("B1").Value = 2014 ' Set cell B1 to 2014
    oWorksheet.Range("C1").Value = 2015 ' Set cell C1 to 2015
    oWorksheet.Range("D1").Value = 2016 ' Set cell D1 to 2016

    ' Set headers in A2 and A3
    oWorksheet.Range("A2").Value = "Projected Revenue" ' Set cell A2 to "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"    ' Set cell A3 to "Estimated Costs"

    ' Set data values
    oWorksheet.Range("B2").Value = 200 ' Set cell B2 to 200
    oWorksheet.Range("B3").Value = 250 ' Set cell B3 to 250
    oWorksheet.Range("C2").Value = 240 ' Set cell C2 to 240
    oWorksheet.Range("C3").Value = 260 ' Set cell C3 to 260
    oWorksheet.Range("D2").Value = 280 ' Set cell D2 to 280
    oWorksheet.Range("D3").Value = 280 ' Set cell D3 to 280

    ' Add a 3D bar chart to the worksheet
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Top:=70, Width:=300, Height:=200)
    With oChart.Chart
        .ChartType = xlBarClustered ' Set chart type to 3D bar
        .SetSourceData Source:=oWorksheet.Range("A1:D3") ' Set data range
        .HasTitle = True ' Enable chart title
        .ChartTitle.Text = "Financial Overview" ' Set chart title text
        .ChartTitle.Font.Size = 13 ' Set chart title font size

        ' Set fill for the first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark gray fill

        ' Set fill for the second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange fill

        ' Set fill for the chart title
        .ChartTitle.Format.Fill.ForeColor.RGB = RGB(128, 128, 128) ' Gray fill
    End With
End Sub
```