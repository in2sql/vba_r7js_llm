# QueryTable TextFileTrailingMinusNumbers Property

## Business Description
True for Microsoft Excel to treat numbers imported as text that begin with a "-" symbol as a negative symbol. False for Excel to treat numbers imported as text that begin with a "-" symbol as text. Read/write Boolean.

## Behavior
Truefor Microsoft Excel to treat numbers imported as text that begin with a "-" symbol as a negative symbol.Falsefor Excel to treat numbers imported as text that begin with a "-" symbol as text. Read/writeBoolean.

## Example Usage
```vba
Sub CheckQueryTableSetting() 
 
 ' Determine setting for TextFileTrailingMinusNumbers 
 If Range("A1").QueryTable.TextFileTrailingMinusNumbers= True Then 
 MsgBox "Numbers imported as text that begin with a '-' symbol " & _ 
 "will be treated as a negative symbol." 
 Else 
 MsgBox "Numbers imported as text that begin with a '-' symbol " & _ 
 "will not be treated as a negative symbol." 
 End If 
 
End Sub
```