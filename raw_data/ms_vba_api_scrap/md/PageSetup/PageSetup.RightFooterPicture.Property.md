# PageSetup RightFooterPicture Property

## Business Description
Returns a Graphic object that represents the picture for the right section of the footer. Used to set attributes of the picture.

## Behavior
Returns aGraphicobject that represents the picture for the right section of the footer. Used to set attributes of the picture.

## Example Usage
```vba
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.RightFooterPicture.FileName = "C:\Sample.jpg" 
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
 
 ' Enable the image to show up in the right footer. 
 ActiveSheet.PageSetup.RightFooter = "&G" 
 
End Sub
```