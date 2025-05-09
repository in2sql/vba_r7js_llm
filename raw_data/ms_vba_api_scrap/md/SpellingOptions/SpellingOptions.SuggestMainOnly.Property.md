# SpellingOptions SuggestMainOnly Property

## Business Description
When set to True, instructs Microsoft Excel to suggest words from only the main dictionary, for using the spelling checker. False removes the limits of suggesting words from only the main dictionary, for using the spelling checker. Read/write Boolean.

## Behavior
When set toTrue, instructs Microsoft Excel to suggest words from only the main dictionary, for using the spelling checker.Falseremoves the limits of suggesting words from only the main dictionary, for using the spelling checker. Read/writeBoolean.

## Example Usage
```vba
Sub UsingMainDictionary() 
 
 ' Check the setting of suggesting words only from the main dictionary. 
 If Application.SpellingOptions.SuggestMainOnly= True Then 
 MsgBox "Spell checking option suggestions will only come from the main dictionary." 
 Else 
 MsgBox "Spell checking option suggestions are not limited to the main dictionary." 
 End If 
 
End Sub
```