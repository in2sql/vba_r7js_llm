# How to: Refer to More Than One Sheet

## Business Description
Use the Array function to identify a group of sheets. The following example selects three sheets in the active workbook.

## Behavior
Use theArrayfunction to identify a group of sheets. The following example selects three sheets in the active workbook.

## Example Usage
```vba
Sub Several() 
 Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Select 
End Sub
```