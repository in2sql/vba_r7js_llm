# PageSetup LeftHeaderPicture Property

## Business Description
Returns a Graphic object that represents the picture for the left section of the header. Used to set attributes about the picture.

## Behavior
Returns aGraphicobject that represents the picture for the left section of the header. Used to set attributes about the picture.

## Example Usage
```vba
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.LeftHeaderPicture.FileName = "C:\Sample.jpg" 
 .Height = 275.25 
 .Width = 463.5 
 .Brightness = 0.36 
 .ColorType = msoPictureGrayscale 
 .Contrast = 0.39 
 .CropBottom = -14.4 
 .CropLeft = -28.8 
 .CropRight = -14.4 
 .CropTop = 21.6 
 End With 
 
 ' Enable the image to show up in the left header. 
 ActiveSheet.PageSetup.LeftHeader = "&G" 
 
End Sub
```