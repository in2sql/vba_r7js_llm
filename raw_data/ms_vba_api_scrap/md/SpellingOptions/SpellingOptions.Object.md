# SpellingOptions Object

## Business Description
Represents the various spell checking options for a worksheet.

## Behavior
Represents the various spell checking options for a worksheet.

## Example Usage
```vba
Sub IgnoreAllCAPS() 
 
 ' Place mispelled versions of the same word in all caps and mixed case. 
 Range("A1").Formula = "Testt" 
 Range("A2").Formula = "TESTT" 
 
 With Application.SpellingOptions 
 .SuggestMainOnly = True 
 .IgnoreCaps = True 
 End With 
 
 ' Run a spell check. 
 Cells.CheckSpelling 
 
End Sub
```