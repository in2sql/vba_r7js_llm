# Chart OLEObjects Method

## Business Description
Returns an object that represents either a single OLE object (an OLEObject ) or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.

## Behavior
Returns an object that represents either a single OLE object (anOLEObject) or a collection of all OLE objects (anOLEObjectscollection) on the chart or sheet. Read-only.

## Example Usage
```vba
Set newSheet = Worksheets.Add 
i = 2 
newSheet.Range("A1").Value = "Name" 
newSheet.Range("B1").Value = "Link Type" 
For Each obj In Worksheets("Sheet1").OLEObjectsnewSheet.Cells(i, 1).Value = obj.Name 
 If obj.OLEType = xlOLELink Then 
 newSheet.Cells(i, 2) = "Linked" 
 Else 
 newSheet.Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```