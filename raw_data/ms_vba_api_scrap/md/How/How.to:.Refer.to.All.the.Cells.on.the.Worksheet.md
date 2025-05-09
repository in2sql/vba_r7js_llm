# How to: Refer to All the Cells on the Worksheet

## Business Description
When you apply the Cells property to a worksheet without specifying an index number, the method returns a Range object that represents all the cells on the worksheet.

## Behavior
When you apply theCellsproperty to a worksheet without specifying an index number, the method returns aRangeobject that represents all the cells on the worksheet. The followingSubprocedure clears the contents from all the cells on Sheet1 in the active workbook.

## Example Usage
```vba
Sub ClearSheet() 
 Worksheets("Sheet1").Cells.ClearContents 
End Sub
```