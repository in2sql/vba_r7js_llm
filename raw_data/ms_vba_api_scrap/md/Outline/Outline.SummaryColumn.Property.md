# Outline SummaryColumn Property

## Business Description
Returns or sets the location of the summary columns in the outline. Read/write XlSummaryColumn.

## Behavior
Returns or sets the location of the summary columns in the outline.   Read/writeXlSummaryColumn.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Selection.AutoOutline 
With ActiveSheet.Outline 
 .SummaryRow = xlAbove 
 .SummaryColumn= xlRight 
 .AutomaticStyles = True 
End With
```