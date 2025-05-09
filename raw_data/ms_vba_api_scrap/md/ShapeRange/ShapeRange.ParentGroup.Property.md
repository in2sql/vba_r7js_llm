# ShapeRange ParentGroup Property

## Business Description
Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.

## Behavior
Returns aShapeobject that represents the common parent shape of a child shape or a range of child shapes.

## Example Usage
```vba
Sub ParentGroup() 
 
 Dim pgShape As Shape 
 
 With ActiveSheet.Shapes 
 .AddShape Type:=1, Left:=10, Top:=10, _ 
 Width:=100, Height:=100 
 .AddShape Type:=2, Left:=110, Top:=120, _ 
 Width:=100, Height:=100 
 .Range(Array(1, 2)).Group 
 End With 
 
 ' Using the child shape in the group get the Parent shape. 
 Set pgShape = ActiveSheet.Shapes(1).GroupItems(1).ParentGroupMsgBox "The two shapes will now be deleted." 
 
 ' Delete the parent shape. 
 pgShape.Delete 
 
End Sub
```