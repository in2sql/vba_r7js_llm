# Graphic Object

## Business Description
Contains properties that apply to header and footer picture objects.

## Behavior
Contains properties that apply to header and footer picture objects.

## Example Usage
```vba
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.LeftFooterPicture 
 .FileName = "C:\Sample.jpg" 
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
 
 ' Enable the image to show up in the left footer. 
 ActiveSheet.PageSetup.LeftFooter = "&G" 
 
End Sub
```