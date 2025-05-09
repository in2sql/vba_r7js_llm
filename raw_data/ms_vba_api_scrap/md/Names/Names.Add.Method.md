# Names Add Method

## Business Description
Defines a new name for a range of cells.

## Behavior
Defines a new name for a range of cells.

## Example Usage
```vba
Sub MakeRange() 
 
    ActiveWorkbook.Names.Add_ 
        Name:="tempRange", _ 
        RefersTo:="=Sheet1!$A$1:$D$3" 
 
End Sub
```