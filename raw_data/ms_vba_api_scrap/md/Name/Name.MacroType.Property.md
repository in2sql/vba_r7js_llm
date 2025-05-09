# Name MacroType Property

## Business Description
Returns or sets what the name refers to. Read/write XlXLMMacroType.

## Behavior
Returns or sets what the name refers to. Read/writeXlXLMMacroType.

## Example Usage
```vba
With ActiveWorkbook.Names(1) 
 If .MacroType<> xlNotXLM Then 
 MsgBox "The category for this name is " & .Category 
 Else 
 MsgBox "This name does not refer to" & _ 
 " a custom function or command." 
 End If 
End With
```