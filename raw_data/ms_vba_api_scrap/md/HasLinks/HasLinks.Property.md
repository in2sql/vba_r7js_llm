# HasLinks Property

## Business Description
True if the specified chart has links to an external data source. Read-only Boolean.

## Behavior
Trueif the specified chart has links to an external data source. Read-onlyBoolean.

## Example Usage
```vba
With myChart.Application 
 If .HasLinks= False Then 
 .DataSheet.Range("A1:D4").Clear 
 End If 
End With
```