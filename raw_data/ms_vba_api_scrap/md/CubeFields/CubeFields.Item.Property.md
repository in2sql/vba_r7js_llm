# CubeFields Item Property

## Business Description
Returns a single object from a collection.

## Behavior
Returns a single object from a collection.

## Example Usage
```vba
blnFoundName = False 
For Each objPT in ActiveSheet.PivotTables 
 Set objCubeField = _ 
 objPT.CubeFields.Item(1) 
 If instr(1,objCubeField.Name, "Paris") <> 0 Then 
 blnFoundName = True 
 Exit For 
 End If 
Next objPT
```