# Workbook EnableAutoRecover Property

## Business Description
Saves changed files, of all formats, on a timed interval. Read/write Boolean.

## Behavior
Saves changed files, of all formats, on a timed interval. Read/writeBoolean.

## Example Usage
```vba
Sub UseAutoRecover() 
 
 ' Check to see if the feature is enabled, if not, enable it. 
 If ActiveWorkbook.EnableAutoRecover= False Then 
 ActiveWorkbook.EnableAutoRecover= True 
 MsgBox "The AutoRecover feature has been enabled." 
 Else 
 MsgBox "The AutoRecover feature is already enabled." 
 End If 
 
End Sub
```