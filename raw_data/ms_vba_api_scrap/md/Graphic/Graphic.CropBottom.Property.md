# Graphic CropBottom Property

## Business Description
Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

## Behavior
Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/writeSingle.

## Example Usage
```vba
percentToCrop = InputBox( _ 
 "What percentage do you want to crop off" & _ 
 " the bottom of this picture?") 
Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1) 
With shapeToCrop.Duplicate 
 .ScaleHeight 1, True 
 origHeight = .Height 
 .Delete 
End With 
cropPoints = origHeight * percentToCrop / 100 
shapeToCrop.PictureFormat.CropBottom= cropPoints
```