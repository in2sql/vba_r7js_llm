# PivotTable ClearTable Method

## Business Description
The ClearTable method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables.

## Behavior
TheClearTablemethod is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.

## Example Usage
```vba
ActiveSheet.PivotTables(1).ClearTable()
```