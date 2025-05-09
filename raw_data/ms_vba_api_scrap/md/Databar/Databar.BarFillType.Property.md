# Databar BarFillType Property

## Business Description
Returns or sets how a data bar is filled with color. Read/write

## Behavior
Returns or sets how a data bar is filled with color. Read/write

## Example Usage
```vba
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
myDataBar.BarFillType= xlDataBarFillSolid
```