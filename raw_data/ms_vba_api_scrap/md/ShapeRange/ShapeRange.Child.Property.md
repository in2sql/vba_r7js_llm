# ShapeRange Child Property

## Business Description
Returns msoTrue if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriStatehttp://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515(Office.15).aspx.

## Behavior
ReturnsmsoTrueif the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-onlyMsoTriState.

## Example Usage
```vba
Sub FillChildShape() 
 
    'Select the first shape in the drawing canvas. 
    ActiveSheet.Shapes(1).CanvasItems(1).Select 
 
    'Fill selected shape if it is a child shape. 
    If Selection.ShapeRange.Child= msoTrue Then 
        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(100, 0, 200) 
    Else 
        MsgBox "This shape is not a child shape." 
    End If 
 
End Sub
```