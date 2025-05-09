# SpellingOptions IgnoreCaps Property

## Business Description
False instructs Microsoft Excel to check for uppercase words, True instructs Excel to ignore words in uppercase when using the spelling checker. Read/write Boolean.

## Behavior
Falseinstructs Microsoft Excel to check for uppercase words,Trueinstructs Excel to ignore words in uppercase when using the spelling checker. Read/writeBoolean.

## Example Usage
```vba
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreCaps= True Then 
 MsgBox "Spelling options for checking uppercase words is disabled." 
 Else 
 MsgBox "Spelling options for checking uppercase words is enabled." 
 End If 
 
End Sub
```