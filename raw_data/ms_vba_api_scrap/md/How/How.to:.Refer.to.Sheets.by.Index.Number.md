# How to: Refer to Sheets by Index Number

## Business Description
An index number is a sequential number assigned to a sheet, based on the position of its sheet tab (counting from the left) among sheets of the same type.

## Behavior
An index number is a sequential number assigned to a sheet, based on the position of its sheet tab (counting from the left) among sheets of the same type. The following procedure uses theWorksheetsproperty to activate the first worksheet in the active workbook.

## Example Usage
```vba
Sub FirstOne() 
 Worksheets(1).Activate 
End Sub
```