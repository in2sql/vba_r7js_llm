# Copy Method

## Business Description
Copy method as it applies to the ChartArea object.

## Behavior
Copy method as it applies to theChartAreaobject.

## Example Usage
```vba
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:D4").Copy_ 
 Destination:= mySheet.Range("E5")
```