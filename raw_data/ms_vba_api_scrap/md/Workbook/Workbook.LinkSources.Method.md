# Workbook LinkSources Method

## Business Description
Returns an array of links in the workbook. The names in the array are the names of the linked documents, editions, or DDE or OLE servers. Returns Empty if there are no links.

## Behavior
Returns an array of links in the workbook. The names in the array are the names of the linked documents, editions, or DDE or OLE servers. ReturnsEmptyif there are no links.

## Example Usage
```vba
aLinks = ActiveWorkbook.LinkSources(xlOLELinks) 
If Not IsEmpty(aLinks) Then 
 For i = 1 To UBound(aLinks) 
 MsgBox "Link " & i & ":" & Chr(13) & aLinks(i) 
 Next i 
End If
```