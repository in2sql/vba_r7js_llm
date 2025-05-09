# SpellingOptions IgnoreFileNames Property

## Business Description
False instructs Microsoft Excel to check for Internet and file addresses, True instructs Excel to ignore Internet and file addresses when using the spell checker. Read/write Boolean.

## Behavior
Falseinstructs Microsoft Excel to check for Internet and file addresses,Trueinstructs Excel to ignore Internet and file addresses when using the spell checker. Read/writeBoolean.

## Example Usage
```vba
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreFileNames= True Then 
 MsgBox "Spelling options for checking Internet and file addresses is disabled." 
 Else 
 MsgBox "Spelling options for checking Internet and file addresses is enabled." 
 End If 
 
End Sub
```