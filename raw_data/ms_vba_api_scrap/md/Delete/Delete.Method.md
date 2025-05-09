# Delete Method

## Business Description
Delete method as it applies to all objects in the Applies To list except the Range object.

## Behavior
Delete method as it applies to all objects in the Applies To list except theRangeobject.

## Example Usage
```vba
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:D10").DeleteShift:=xlShiftToLeft
```