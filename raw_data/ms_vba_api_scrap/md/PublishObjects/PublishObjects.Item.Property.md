# PublishObjects Item Property

## Business Description
Returns a single object from a collection.

## Behavior
Returns a single object from a collection.

## Example Usage
```vba
strTargetDivID = ActiveWorkbook.PublishObjects.Item(1).DivID 
Open "\\server1\reports\q198.htm" For Input As #1 
Open "\\server1\reports\newq1.htm" For Output As #2 
While Not EOF(1) 
 Line Input #1, strFileLine 
 If InStr(strFileLine, strTargetDivID) > 0 And _ 
 InStr(strFileLine, "<div") > 0 Then 
 Print #2, "<!--Saved item-->" 
 End If 
 Print #2, strFileLine 
Wend 
Close #2 
Close #1
```