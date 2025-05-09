# This script populates an Excel sheet with data and creates a 3D bar chart with customized title, legend font size, and series colors.
# Этот скрипт заполняет лист Excel данными и создает 3D столбчатую диаграмму с настроенным заголовком, размером шрифта легенды и цветами рядов.

```javascript
// This example specifies the legend font size.
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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);
oChart.SetLegendFontSize(13);
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' This script populates an Excel sheet with data and creates a 3D bar chart with customized title, legend font size, and series colors.
' Этот скрипт заполняет лист Excel данными и создает 3D столбчатую диаграмму с настроенным заголовком, размером шрифта легенды и цветами рядов.

Sub CreateChart()
    Dim ws As Worksheet
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xl3DColumn
        
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set legend font size
        .Legend.Font.Size = 13
        
        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```