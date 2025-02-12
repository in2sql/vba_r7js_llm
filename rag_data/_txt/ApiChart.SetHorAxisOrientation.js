### Description / Описание
**English:** This code populates an Excel sheet with financial data for the years 2014 to 2016, creates a 3D bar chart titled "Financial Overview", sets the horizontal axis orientation, and applies specific fill colors to the chart series.

**Russian:** Этот код заполняет лист Excel финансовыми данными за годы 2014–2016, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор", настраивает ориентацию горизонтальной оси и применяет определенные цвета заливки к сериям диаграммы.

```vba
' VBA Code Equivalent to OnlyOffice JS Example

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim rng As Range

    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

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

    ' Define the range for the chart
    Set rng = oWorksheet.Range("A1:D3")

    ' Add a 3D Bar Chart
    Set oChart = oWorksheet.Shapes.AddChart2(251, xlBarClustered, 100, 70, 360, 180).Chart
    oChart.SetSourceData Source:=rng

    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"

    ' Set horizontal axis orientation
    oChart.Axes(xlCategory).ReversePlotOrder = False

    ' Set series fill colors
    With oChart.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
        .Solid
    End With

    With oChart.SeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61) ' Orange
        .Solid
    End With
End Sub
```

```javascript
// JavaScript Code Using OnlyOffice API

// This example specifies the direction of the data displayed on the horizontal axis.
// Этот пример задает направление отображаемых данных по горизонтальной оси.
var oWorksheet = Api.GetActiveSheet();

// Populate the cells with data
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

// Add a 3D Bar Chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);

// Set horizontal axis orientation
oChart.SetHorAxisOrientation(false);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```