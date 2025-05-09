# Font Background Property

## Business Description
Returns or sets the type of background for text used in charts. Read/write Variant which is set to one of the constants of XlBackground.

## Behavior
Returns or sets the type of  background for text used in charts. Read/writeVariantwhich is set to one of the constants ofXlBackground.

## Example Usage
```vba
Sub UseBackground() 
 
 With Worksheets(1).ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "Rainfall Totals by Month" 
 With .ChartTitle.Font 
 .Size = 10 
 .Background= xlBackgroundTransparent 
 End With 
 End With 
 
End Sub
```