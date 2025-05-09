# CalculatedMembers AddCalculatedMember Method

## Business Description
Adds a calculated field or calculated item to a PivotTable.

## Behavior
Adds a calculated field or calculated item to a PivotTable.

## Example Usage
```vba
Sub AddCalculatedMeasure()

Dim pvt As PivotTable
Dim strName As String
Dim strFormula As String
Dim strDisplayFolder As String
Dim strMeasureGroup As String

Set pvt = Sheet1.PivotTables("PivotTable1")
strName = "[Measures].[Internet Sales Amount 25 %]"
strFormula = "[Measures].[Internet Sales Amount]*1.25"
strDisplayFolder = "My Folder\Percent Calculations" 
strMeasureGroup = "Internet Sales"

pvt.CalculatedMembers.AddCalculatedMemberName:=strName, Formula:=strFormula, Type:=xlCalculatedMeasure, DisplayFolder:=strDisplayFolder, MeasureGroup:=strMeasureGroup, NumberFormat:=xlNumberFormatTypePercent

End Sub
```