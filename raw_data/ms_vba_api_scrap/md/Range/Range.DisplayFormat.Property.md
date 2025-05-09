# Range DisplayFormat Property

## Business Description
Returns a DisplayFormat object that represents the display settings for the specified range. Read-only

## Behavior
Returns aDisplayFormatobject that represents the display settings for the specified range. Read-only

## Example Usage
```vba
Function getColorIndex()
   getColorIndex = ActiveCell.DisplayFormat.Interior.ColorIndex
End Function
```