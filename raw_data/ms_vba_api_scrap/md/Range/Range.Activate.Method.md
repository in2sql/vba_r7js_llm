# Range Activate Method

## Business Description
Activates a single cell, which must be inside the current selection. To select a range of cells, use the Select method.

## Behavior
Activates a single cell, which must be inside the current selection. To select a range of cells, use theSelectmethod.

## Example Usage
```vba
Worksheets("Sheet1").ActivateRange("A1:C3").Select 
Range("B2").Activate
```