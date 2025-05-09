# Using Events with Embedded Charts

## Business Description
Events are enabled for chart sheets by default. Before you can use events with a Chart object that represents an embedded chart, you must create a new class module and declare an object of type Chart with events.

## Behavior
Events are enabled for chart sheets by default. Before you can use events with aChartobject that represents an embedded chart, you must create a new class module and declare an object of typeChartwith events. For example, assume that a new class module is created and named EventClassModule. The new class module contains the following code.

## Example Usage
```vba
Dim myClassModule As New EventClassModule 
 
Sub InitializeChart() 
 Set myClassModule.myChartClass = _ 
 Charts(1).ChartObjects(1).Chart 
End Sub
```