# OLEObject AutoUpdate Property

## Business Description
True if the OLE object is updated automatically when the source changes. Valid only if the object is linked (its OLEType property must be xlOLELink). Read-only Boolean.

## Behavior
Trueif the OLE object is updated automatically when the source changes. Valid only if the object is linked (itsOLETypeproperty must bexlOLELink). Read-onlyBoolean.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A1").Value = "Name" 
Range("B1").Value = "Link Status" 
Range("C1").Value = "AutoUpdate Status" 
i = 2 
For Each obj In ActiveSheet.OLEObjects 
 Cells(i, 1) = obj.Name 
 If obj.OLEType = xlOLELink Then 
 Cells(i, 2) = "Linked" 
 Cells(i, 3) = obj.AutoUpdateElse 
 Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```