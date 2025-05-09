# SpellingOptions KoreanUseAutoChangeList Property

## Business Description
When set to True, this enables Microsoft Excel to use the auto-change list for Korean words when using the spelling checker. Read/write Boolean.

## Behavior
When set toTrue, this enables Microsoft Excel to use the auto-change list for Korean words when using the spelling checker. Read/writeBoolean.

## Example Usage
```vba
Sub KoreanSpellCheck() 
 
 If Application.SpellingOptions.KoreanUseAutoChangeList= True Then 
 MsgBox "The spell checking feature to auto-change Korean words is on." 
 Else 
 MsgBox "The spell checking feature to auto-change Korean words is off." 
 End If 
 
End Sub
```