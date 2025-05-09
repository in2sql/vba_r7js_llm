# How to: Refer to Cells by Using Index Numbers

## Business Description
You can use the Cells property to refer to a single cell by using row and column index numbers. This property returns a Range object that represents a single cell. In the following example, Cells(6,1) returns cell A6 on Sheet1.

## Behavior
You can use theCellsproperty to refer to a single cell by using row and column index numbers. This property returns aRangeobject that represents a single cell. In the following example,Cells(6,1)returns cell A6 on Sheet1. TheValueproperty is then set to 10.

## Example Usage
```vba
Sub EnterValue() 
 Worksheets("Sheet1").Cells(6, 1).Value = 10 
End Sub
```