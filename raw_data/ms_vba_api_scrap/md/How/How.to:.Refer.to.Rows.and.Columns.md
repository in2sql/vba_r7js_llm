# How to: Refer to Rows and Columns

## Business Description
Use the Rows property or the Columns property to work with entire rows or columns. These properties return a Range object that represents a range of cells. In the following example, Rows(1) returns row one on Sheet1.

## Behavior
Use theRowsproperty or theColumnsproperty to work with entire rows or columns. These properties return aRangeobject that represents a range of cells. In the following example,Rows(1)returns row one on Sheet1. TheBoldproperty of theFontobject for the range is then set toTrue.

## Example Usage
```vba
Sub RowBold() 
    Worksheets("Sheet1").Rows(1).Font.Bold = True 
End Sub
```