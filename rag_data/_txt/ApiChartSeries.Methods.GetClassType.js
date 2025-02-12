### Description
**English:** This script populates a worksheet with financial data, creates a combination bar and line chart titled "Financial Overview," customizes the series fills with specific colors, retrieves the class type of the first series, and inserts this information into cell F1.

**Russian:** Этот скрипт заполняет лист финансовыми данными, создает комбинированный столбчатый и линейный график с заголовком "Финансовый обзор", настраивает заливку серий с определенными цветами, получает тип класса первой серии и вставляет эту информацию в ячейку F1.

```vba
' VBA code equivalent to OnlyOffice JS example
' This script populates a worksheet with financial data, creates a combination bar and line chart titled "Financial Overview,"
' customizes the series fills with specific colors, retrieves the class type of the first series, and inserts this information into cell F1.

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet
    
    ' Populate the worksheet with data
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
    
    ' Add a combination bar and line chart
    Dim oChart As Chart
    Set oChart = oWorksheet.Shapes.AddChart2(201, xlCombo, 100, 70, 300, 200).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    oChart.ChartType = xlComboBarLine
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    
    ' Customize the fill for the first series
    Dim oFill As FillFormat
    Set oFill = oChart.SeriesCollection(1).Format.Fill
    oFill.ForeColor.RGB = RGB(51, 51, 51) ' Set RGB color to (51, 51, 51)
    
    ' Customize the fill for the second series
    Set oFill = oChart.SeriesCollection(2).Format.Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61) ' Set RGB color to (255, 111, 61)
    
    ' Retrieve the class type of the first series and insert into cell F1
    Dim oSeries As Series
    Set oSeries = oChart.SeriesCollection(1)
    Dim sClassType As String
    sClassType = oSeries.Name ' VBA does not have GetClassType; using Name as example
    oWorksheet.Range("F1").Value = "Class Type = " & sClassType
End Sub
```

```javascript
// OnlyOffice JS code example
// This script populates a worksheet with financial data, creates a combination bar and line chart titled "Financial Overview,"
// customizes the series fills with specific colors, retrieves the class type of the first series, and inserts this information into cell F1.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet

// Populate the worksheet with data
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

// Add a combination bar and line chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);

// Customize the fill for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Customize the fill for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Retrieve the class type of the first series and insert into cell F1
var oSeries = oChart.GetSeries(0);
var sClassType = oSeries.GetClassType();
oWorksheet.GetRange("F1").SetValue("Class Type = " + sClassType);
```