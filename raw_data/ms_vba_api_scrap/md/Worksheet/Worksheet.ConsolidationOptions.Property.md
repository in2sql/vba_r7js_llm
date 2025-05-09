# Worksheet ConsolidationOptions Property

## Business Description
Returns a three-element array of consolidation options, as shown in the following table. If the element is True, that option is set. Read-only Variant.

## Behavior
Returns a three-element array of consolidation options, as shown in the following table. If the element isTrue, that option is set. Read-onlyVariant.

## Example Usage
```vba
Set newSheet = Worksheets.Add 
aOptions = Worksheets("Sheet1").ConsolidationOptionsnewSheet.Range("A1").Value = "Use labels in top row" 
newSheet.Range("A2").Value = "Use labels in left column" 
newSheet.Range("A3").Value = "Create links to source data" 
For i = 1 To 3 
 If aOptions(i) = True Then 
 newSheet.Cells(i, 2).Value = "True" 
 Else 
 newSheet.Cells(i, 2).Value = "False" 
 End If 
Next i 
newSheet.Columns("A:B").AutoFit
```