**Description / Описание**

English:  
This code sets specific cell values in a worksheet and creates a 3D bar chart titled "Financial Overview" with customized categories and series fills.

Russian:  
Этот код устанавливает определенные значения ячеек в рабочем листе и создает 3D столбчатую диаграмму с заголовком "Финансовый обзор", настроенными категориями и заливкой серий.

---

```vba
' VBA Code to set cell values and create a 3D bar chart with customized categories and fills

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim fillColor As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set cell values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("B4").Value = 2020
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("C4").Value = 2021
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    ws.Range("D4").Value = 2022
    
    ' Add a 3D Bar Chart
    Set cht = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    With cht.Chart
        .SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3")
        .ChartType = xl3DBarClustered
        
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set category labels
        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .Axes(xlCategory).CategoryNames = ws.Range("'Sheet1'!$B$4:$D$4")
        
        ' Set series fill colors
        ' Series 1
        fillColor = RGB(51, 51, 51)
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor
        .SeriesCollection(1).Format.Fill.Visible = msoTrue
        
        ' Series 2
        fillColor = RGB(255, 111, 61)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor
        .SeriesCollection(2).Format.Fill.Visible = msoTrue
    End With
End Sub
```

```javascript
// JavaScript Code to set cell values and create a 3D bar chart with customized categories and fills using OnlyOffice API

function createFinancialOverviewChart(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set cell values
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);
    oWorksheet.GetRange("A2").SetValue("Projected Revenue");
    oWorksheet.GetRange("A3").SetValue("Estimated Costs");
    oWorksheet.GetRange("B2").SetValue(200);
    oWorksheet.GetRange("B3").SetValue(250);
    oWorksheet.GetRange("B4").SetValue(2020);
    oWorksheet.GetRange("C2").SetValue(240);
    oWorksheet.GetRange("C3").SetValue(260);
    oWorksheet.GetRange("C4").SetValue(2021);
    oWorksheet.GetRange("D2").SetValue(280);
    oWorksheet.GetRange("D3").SetValue(280);
    oWorksheet.GetRange("D4").SetValue(2022);
    
    // Add a 3D Bar Chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Set category labels
    oChart.SetCatFormula("'Sheet1'!$B$4:$D$4");
    
    // Set series 1 fill color
    var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill1, 0, false);
    
    // Set series 2 fill color
    var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill2, 1, false);
}
```