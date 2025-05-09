# OLEObject OLEType Property

## Business Description
Returns the OLE object type. Can be one of the following XlOLEType constants: xlOLELink or xlOLEEmbed.

## Behavior
Returns the OLE object type. Can be one of the followingXlOLETypeconstants:xlOLELinkorxlOLEEmbed. ReturnsxlOLELinkif the object is linked (it exists outside of the file), or returnsxlOLEEmbedif the object is embedded (it's entirely contained within the file). Read-onlyLong.

## Example Usage
```vba
Set newSheet = Worksheets.Add 
i = 2 
newSheet.Range("A1").Value = "Name" 
newSheet.Range("B1").Value = "Link Type" 
For Each obj In Worksheets("Sheet1").OLEObjects 
 newSheet.Cells(i, 1).Value = obj.Name 
 If obj.OLEType= xlOLELink Then 
 newSheet.Cells(i, 2) = "Linked" 
 Else 
 newSheet.Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```