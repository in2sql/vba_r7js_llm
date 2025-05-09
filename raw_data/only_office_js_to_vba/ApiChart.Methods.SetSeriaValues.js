**Description**
English: This script sets specific cell values, creates a 3D bar chart, and customizes the chart's title and series fills.
Russian: Этот скрипт устанавливает значения в определенные ячейки, создает 3D столбчатую диаграмму и настраивает заголовок диаграммы и заливку серий.

```vba
' This VBA script sets specific cell values and creates a 3D bar chart
Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set year values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("B4").Value = 260
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("C4").Value = 270
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    ws.Range("D4").Value = 300
    
    ' Create chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=250)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3")
        .ChartType = xlBar3DClustered
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Set series values
        .SeriesCollection(1).Values = ws.Range("'Sheet1'!$B$4:$D$4")
        
        ' Customize data labels
        Dim ser As Series
        Set ser = .SeriesCollection(1)
        ser.HasDataLabels = True
        ser.DataLabels.ShowValue = True
        
        ' Set series fill colors
        ser.Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
        ser.Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    End With
End Sub
```

```javascript
// This JavaScript code sets specific cell values, creates a 3D bar chart, and customizes the chart's title and series fills using the OnlyOffice API
var oWorksheet = Api.GetActiveSheet();

// Set year values
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("B4").SetValue(260);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("C4").SetValue(270);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
oWorksheet.GetRange("D4").SetValue(300);

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series values
oChart.SetSeriaValues("'Sheet1'!$B$4:$D$4", 1);

// Customize data labels
oChart.SetShowPointDataLabel(1, 0, false, false, true, false);
oChart.SetShowPointDataLabel(1, 1, false, false, true, false);
oChart.SetShowPointDataLabel(1, 2, false, false, true, false);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```