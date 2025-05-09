# ErrorCheckingOptions IndicatorColorIndex Property

## Business Description
Returns or sets the color of the indicator for error checking options. Read/write XlColorIndex.

## Behavior
Returns or sets the color of the indicator for  error checking options. Read/writeXlColorIndex.

## Example Usage
```vba
Sub CheckIndexColor() 
 
 If Application.ErrorCheckingOptions.IndicatorColorIndex= xlColorIndexAutomatic Then 
 MsgBox "Your indicator color for error checking is set to the default system color." 
 Else 
 MsgBox "Your indicator color for error checking is not set to the default system color." 
 End If 
 
End Sub
```