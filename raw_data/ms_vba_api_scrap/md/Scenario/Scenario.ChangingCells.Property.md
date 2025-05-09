# Scenario ChangingCells Property

## Business Description
Returns a Range object that represents the changing cells for a scenario. Read-only.

## Behavior
Returns aRangeobject that represents the changing cells for a scenario. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveSheet.Scenarios(1).ChangingCells.Select
```