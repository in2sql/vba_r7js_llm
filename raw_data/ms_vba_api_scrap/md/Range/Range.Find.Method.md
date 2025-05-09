# Range Find Method

## Business Description
Finds specific information in a range.

## Behavior
Finds specific information in a range.

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