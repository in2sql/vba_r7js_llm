# PivotCell ServerActions Property

## Business Description
Represents a collection of actions consisting of OLAP-defined actions which can be executed. The actions are specific to PivotTables existing at a worksheet-level. Read-only

## Behavior
Represents a collection ofactionsconsisting of OLAP-defined actions which can be executed. The actions are specific to PivotTables existing at a worksheet-level. Read-only

## Example Usage
```vba
ActiveSheet.ChartObjects("Chart 1").Chart.PivotLayout.PivotTable.PivotColumnAxis.PivotLines(index of line).PivotLineCells(index of cells).ServerAction("OLAP Action name").Execute
```