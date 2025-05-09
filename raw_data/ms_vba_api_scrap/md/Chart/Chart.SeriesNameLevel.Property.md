# Chart SeriesNameLevel Property

## Business Description
Returns a XlSeriesNameLevel constant referring to the level of where the series names are being sourced from. Integer Read/Write.

## Behavior
Returns aXlSeriesNameLevel Enumeration (Excel)constant referring to the level of where the series names are being sourced from.IntegerRead/Write.

## Example Usage
```vba
Sheets(1).Range("C1:E1").Value2 = "Sample_Row1"
    Sheets(1).Range("C2:E2").Value2 = "Sample_Row2"
    Sheets(1).Range("A3:A5").Value2 = "Sample_ColA"
    Sheets(1).Range("B3:B5").Value2 = "Sample_ColB"
    Sheets(1).Range("C3:E5").Formula = "=row()"
    Dim crt As Chart
    Set crt = Sheets(1).ChartObjects.Add(0, 0, 500, 200).Chart
    crt.SetSourceData Sheets(1).Range("A1:E5")
    ' Set the series names to only use column B
    crt.SeriesNameLevel = 1
    ' Use columns A and B for the series names
    crt.SeriesNameLevel = xlSeriesNameLevelAll
    ' Use row 1 for the category labels
    crt.CategoryLabelLevel = 0
    ' Use rows 1 and 2 for the category labels
    crt.CategoryLabelLevel = xlCategoryLabelLevelAll
```