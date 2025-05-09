# SpellingOptions ArabicModes Property

## Business Description
Returns or sets the mode for the Arabic spelling checker. Read/write XlArabicModes.

## Behavior
Returns or sets the mode for the Arabic spelling checker. Read/writeXlArabicModes.

## Example Usage
```vba
Sub SpellCheck() 
 
 If Application.SpellingOptions.ArabicModes<> xlArabicBothStrict Then 
 Application.SpellingOptions.ArabicModes= xlArabicBothStrict 
 MsgBox "Spell checking for Arabic mode has been changed to a strict setting." 
 Else 
 MsgBox "Spell checking for Arabic mode is already in a strict setting." 
 End If 
 
End Sub
```