# ListObjects Add Method

## Business Description
Creates a new list object.

## Behavior
Creates a new list object.

## Example Usage
```vba
Set objListObject = ActiveWorkbook.Worksheets(1).ListObjects.Add(SourceType:= xlSrcExternal, _ 
Source:= Array(strServerName, strListName, strListGUID), LinkSource:=True, _ 
TableStyleName:=xlGuess, Destination:=Range("A10"))
```