# Graphic CropLeft Property

## Business Description
Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

## Behavior
Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/writeSingle.

## Example Usage
```vba
percentToCrop = InputBox( _ 
 "What percentage do you want to crop" & _ 
 " off the left of this picture?") 
Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1) 
With shapeToCrop.Duplicate 
 .ScaleWidth 1, True 
 origWidth = .Width 
 .Delete 
End With 
cropPoints = origWidth * percentToCrop / 100 
shapeToCrop.PictureFormat.CropLeft= cropPoints
```