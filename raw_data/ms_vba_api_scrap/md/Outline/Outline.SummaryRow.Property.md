# Outline SummaryRow Property

## Business Description
Returns or sets the location of the summary rows in the outline. Read/write XlSummaryRow.

## Behavior
Returns or sets the location of the summary rows in the outline.   Read/writeXlSummaryRow.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Selection.AutoOutline 
With ActiveSheet.Outline 
 .SummaryRow= xlAbove 
 .SummaryColumn = xlRight 
 .AutomaticStyles = True 
End With
```