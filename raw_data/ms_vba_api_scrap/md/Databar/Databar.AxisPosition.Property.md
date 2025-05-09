# Databar AxisPosition Property

## Business Description
Returns or sets the position of the axis of the data bars specified by a conditional formatting rule. Read/write

## Behavior
Returns or sets the position of the axis of the data bars specified by a conditional formatting rule. Read/write

## Example Usage
```vba
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
myDataBar.AxisPosition= xlDataBarAxisMidpoint
```