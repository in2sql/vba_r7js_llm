# SpellingOptions KoreanProcessCompound Property

## Business Description
When set to True, this enables Microsoft Excel to process Korean compound nouns when using the spelling checker. Read/write Boolean.

## Behavior
When set toTrue, this enables Microsoft Excel to process Korean compound nouns when using the spelling checker. Read/writeBoolean.

## Example Usage
```vba
Sub KoreanSpellCheck() 
 
 If Application.SpellingOptions.KoreanProcessCompound= True Then 
 MsgBox "The spell checking feature to process Korean compound nouns is on." 
 Else 
 MsgBox "The spell checking feature to process Korean compound nouns is off." 
 End If 
 
End Sub
```