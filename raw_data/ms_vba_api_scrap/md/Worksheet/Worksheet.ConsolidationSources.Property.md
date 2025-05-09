# Worksheet ConsolidationSources Property

## Business Description
Returns an array of string values that name the source sheets for the worksheet's current consolidation. Returns Empty if there's no consolidation on the sheet. Read-only Variant.

## Behavior
Returns an array of string values that name the source sheets for the worksheet's current consolidation. ReturnsEmptyif there's no consolidation on the sheet. Read-onlyVariant.

## Example Usage
```vba
Set newSheet = Worksheets.Add 
newSheet.Range("A1").Value = "Consolidation Sources" 
aSources = Worksheets("Sheet1").ConsolidationSourcesIf IsEmpty(aSources) Then 
 newSheet.Range("A2").Value = "none" 
Else 
 For i = 1 To UBound(aSources) 
 newSheet.Cells(i + 1, 1).Value = aSources(i) 
 Next i 
End If 
newSheet.Columns("A:B").AutoFit
```