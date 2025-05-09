# SpellingOptions KoreanCombineAux Property

## Business Description
When set to True, Microsoft Excel combines Korean auxiliary verbs and adjectives when spelling is checked. Read/write Boolean.

## Behavior
When set toTrue, Microsoft Excel combines Korean auxiliary verbs and adjectives when spelling is checked. Read/writeBoolean.

## Example Usage
```vba
Sub KoreanSpellCheck() 
 
 If Application.SpellingOptions.KoreanCombineAux= True Then 
 MsgBox "The option to combine Korean auxiliary verbs and adjectives while checking spelling is on." 
 Else 
 MsgBox "The option to combine Korean auxiliary verbs and adjectives while checking spelling is off." 
 End If 
 
End Sub
```