# Characters Count Property

## Business Description
Returns a Long value that represents the number of objects in the collection.

## Behavior
Returns aLongvalue that represents the number of objects in the collection.

## Example Usage
```vba
Sub MakeSuperscript() 
 Dim n As Integer 
 
 n = Worksheets("Sheet1").Range("A1").Characters.CountWorksheets("Sheet1").Range("A1").Characters(n, 1) _ 
 .Font.Superscript = True 
End Sub
```