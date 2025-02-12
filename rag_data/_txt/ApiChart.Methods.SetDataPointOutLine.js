**Description / Описание:**

This code sets financial data in a worksheet, creates a 3D bar chart, and customizes its appearance by setting colors and outlines.

Этот код устанавливает финансовые данные в рабочем листе, создает 3D столбчатую диаграмму и настраивает ее внешний вид, устанавливая цвета и контуры.

---

**VBA Code:**

```vba
' VBA code to set financial data, create a 3D bar chart, and customize its appearance
Sub CreateFinancialChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add 3D bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xl3DBar
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    End With
    
    ' Set outline for the second series
    With chartObj.Chart.SeriesCollection(2).Points(1).Format.Line
        .Visible = msoTrue
        .Weight = 0.5
        .ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
    End With
End Sub
```

---

**OnlyOffice JavaScript Code:**

```javascript
// This example shows how to set the outline to the data point.
function createFinancialChart() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set header values
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);
    
    // Set row labels
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
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Create and set the fill for the first series
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill, 0, false);
    
    // Create and set the fill for the second series
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill, 1, false);
    
    // Create and set the outline for a data point
    var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
    oChart.SetDataPointOutLine(oStroke, 1, 0, false);
}

// Call the function to create the chart
createFinancialChart();
```