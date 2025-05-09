# Worksheet Name Property

## Business Description
Returns or sets a String value that represents the object name.

## Behavior
Returns or sets aStringvalue that represents the object name.

## Example Usage
```vba
' This macro sets today's date as the name for the current sheet 
Sub NameWorksheetByDate() 
    Range("D5").Select 
    Selection.Formula = "=text(now(),""mmm dd yyyy"")" 
    Selection.Copy 
    Selection.PasteSpecial Paste:=xlValues 
    Application.CutCopyMode = False 
    Selection.Columns.AutoFit 
    ActiveSheet.Name = Range("D5").Value 
    Range("D5").Value = "" 
End Sub
```