## Description / Описание

**English:** This script populates an Excel worksheet with financial data, creates a 3D bar chart based on the data, customizes the chart's appearance by setting the title and series colors, and retrieves the chart's class type, inserting it into a specified cell.

**Russian:** Этот скрипт заполняет лист Excel финансовыми данными, создает 3D столбчатую диаграмму на основе этих данных, настраивает внешний вид диаграммы, устанавливая заголовок и цвета серий, а также получает тип класса диаграммы и вставляет его в указанную ячейку.

```vba
' VBA Code to populate worksheet, create a 3D bar chart, customize it, and retrieve the chart's class type

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim oChart As Chart
    Dim sClassType As String
    Dim fillColor1 As Long
    Dim fillColor2 As Long

    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Insert values into cells
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
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=3600, Top:=100, Height:=2500)
    Set oChart = chartObj.Chart
    oChart.ChartType = xlColumn3D ' 3D bar chart

    ' Set the data range for the chart
    oChart.SetSourceData Source:=ws.Range("A1:D3")

    ' Set the chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13

    ' Set series fill colors
    fillColor1 = RGB(51, 51, 51) ' Dark gray
    fillColor2 = RGB(255, 111, 61) ' Orange
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor1
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor2
    oChart.SeriesCollection(3).Format.Fill.ForeColor.RGB = fillColor1

    ' Get the class type of the chart (VBA does not have a direct method, so using TypeName)
    sClassType = TypeName(oChart)
    ws.Range("F1").Value = "Class Type: " & sClassType
End Sub
```

```javascript
// JavaScript Code to populate worksheet, create a 3D bar chart, customize it, and retrieve the chart's class type

// This example gets a class type and inserts it into the table.
var oWorksheet = Api.GetActiveSheet();

// Insert values into cells
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

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get the class type of the chart
var sClassType = oChart.GetClassType();

// Insert the class type into cell F1
oWorksheet.GetRange("F1").SetValue("Class Type: " + sClassType);
```