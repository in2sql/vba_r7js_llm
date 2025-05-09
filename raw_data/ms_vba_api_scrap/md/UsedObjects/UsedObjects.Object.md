# UsedObjects Object

## Business Description
Represents objects that have been allocated in a workbook.

## Behavior
Represents objects that have been allocated in a workbook.

## Example Usage
```vba
Sub CountUsedObjects() 
 
 MsgBox "The number of used objects in this application is: " & _ 
 Application.UsedObjects.Count 
 
End Sub
```