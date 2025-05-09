# Returning an Object from a Collection

## Business Description
The Item property of a collection returns a single object from that collection. The following example sets the firstBook variable to a Workbook object that represents the first workbook in the Workbooks collection.

## Behavior
TheItemproperty of a collection returns a single object from that collection. The following example sets thefirstBookvariable to aWorkbookobject that represents the first workbook in theWorkbookscollection.

## Example Usage
```vba
ActiveWorkbook.Worksheets.Add.Name = "A New Sheet" 
With Worksheets("A New Sheet") 
 .Range("A5:A10").Formula = "=RAND()" 
End With
```