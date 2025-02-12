**Description: This code sets specific cell values in a worksheet and adds a 3D bar chart with a customized title and series colors.  
Описание: Этот код устанавливает значения определенных ячеек на листе и добавляет 3D-столбчатую диаграмму с настраиваемым заголовком и цветами серий.**

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set category labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values for 2014
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("B4").SetValue(2020);

// Set data values for 2015
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("C4").SetValue(2021);

// Set data values for 2016
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
oWorksheet.GetRange("D4").SetValue(2022);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set the category (X-axis) formula
oChart.SetCatFormula("'Sheet1'!$B$4:$D$4");

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' VBA Code Equivalent

Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set category labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values for 2014
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("B4").Value = 2020
    
    ' Set data values for 2015
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("C4").Value = 2021
    
    ' Set data values for 2016
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    ws.Range("D4").Value = 2022
    
    ' Add a 3D bar chart to the worksheet
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 200, 100, 400, 300).Chart ' 251 for 3D bar
    cht.SetSourceData Source:=ws.Range("A1:D3")
    
    ' Set the chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13
    
    ' Set category axis labels
    cht.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    cht.Axes(xlCategory).CategoryNames = ws.Range("B4:D4")
    
    ' Customize series fills
    With cht.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 51, 51)
        .Solid
    End With
    
    With cht.SeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
End Sub
```