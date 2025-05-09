# Name CategoryLocal Property

## Business Description
Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command. Read/write String.

## Behavior
Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command. Read/writeString.

## Example Usage
```vba
With ActiveWorkbook.Names(1) 
 If .MacroType <> xlNone Then 
 MsgBox "The category for this name is " & .CategoryLocalElse 
 MsgBox "This name does not refer to" & _ 
 " a custom function or command." 
 End If 
End With
```