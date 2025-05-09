# Areas object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
There's no singular Area object; individual members of the Areas collection are Range objects. The Areas collection contains one Range object for each discrete, contiguous range of cells within the selection. If the selection contains only one area, the Areas collection contains a single Range object that corresponds to that selection.

## Example
```vba
Set rangeToUse = Selection 
If rangeToUse.Areas.Count = 1 Then 
 myOperation rangeToUse 
Else 
 For Each singleArea in rangeToUse.Areas 
 myOperation singleArea 
 Next 
End If
```

