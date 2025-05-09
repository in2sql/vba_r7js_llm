# Window RangeFromPoint Method

## Business Description
Returns the Shape or Range object that is positioned at the specified pair of screen coordinates. If there isn't a shape located at the specified coordinates, this method returns Nothing.

## Behavior
Returns theShapeorRangeobject that is positioned at the specified pair of screen coordinates. If there isn't a shape located at the specified coordinates, this method returnsNothing.

## Example Usage
```vba
Private Function AltText(ByVal intMouseX As Integer, _ 
 ByVal intMouseY as Integer) As String 
 Set objShape = ActiveWindow.RangeFromPoint_ 
 (x:=intMouseX, y:=intMouseY) 
 If Not objShape Is Nothing Then 
 With objShape 
 Select Case .Type 
 Case msoChart, msoLine, msoPicture: 
 AltText = .AlternativeText 
 Case Else: 
 AltText = "" 
 End Select 
 End With 
 Else 
 AltText = "" 
 End If 
End Function
```