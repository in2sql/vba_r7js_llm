# ListDataFormat Object

## Business Description
The ListDataFormat object holds all the data type properties of the ListColumn object. These properties are read-only.

## Behavior
TheListDataFormatobject holds all the data type properties of theListColumnobject. These properties are read-only.

## Example Usage
```vba
Dim objListObject As ListObject 
Dim objDataRange As Range 
Dim strListGUID as String 
Dim strServerName as String 
 
strServerName = "http://<servername>/_vti_bin" 
strListGUID = "{<listguid>}" 
 
Set objListObject = Sheet1.ListObjects.Add(xlSrcExternal, _ 
 Array(strServerName, strListGUID), True, xlYes, Range("A1")) 
 
With objListObject.ListColumns(2) 
 Set objDataRange = .Range.Offset(1, 0).Resize(.Range.Rows.Count - 2, 1) 
 If .ListDataFormat.Type = xlListDataTypeText And .ListDataFormat.Required Then 
 objDataRange.Value = "Hello World" 
 End If 
End With
```