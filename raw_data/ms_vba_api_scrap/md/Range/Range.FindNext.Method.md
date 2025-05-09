# Range FindNext Method

## Business Description
Continues a search that was begun with the Find method. Finds the next cell that matches those same conditions and returns a Range object that represents that cell. This does not affect the selection or the active cell.

## Behavior
Continues a search that was begun with theFindmethod. Finds the next cell that matches those same conditions and returns aRangeobject that represents that cell. This does not affect the selection or the active cell.

## Example Usage
```vba
With Worksheets(1).Range("a1:a500") 
    Set c = .Find(2, lookin:=xlValues) 
    If Not c Is Nothing Then 
        firstAddress = c.Address 
        Do 
            c.Value = 5 
            Set c = .FindNext(c) 
        Loop While Not c Is Nothing And c.Address <> firstAddress 
    End If 
End With
```