# Hyperlink CreateNewDocument Method

## Business Description
Creates a new document linked to the specified hyperlink.

## Behavior
Creates a new document linked to the specified hyperlink.

## Example Usage
```vba
With Worksheets(1) 
 Set objHyper = _ 
 .Hyperlinks.Add(Anchor:=.Range("A10"), _ 
 Address:="\\Server1\Annual\Report.xls") 
 objHyper.CreateNewDocument_ 
 FileName:="\\Server1\Annual\Report.xls", _ 
 EditNow:=True, Overwrite:=True 
End With
```