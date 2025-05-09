# HeaderFooter Object

## Business Description
Represents a single header or footer. The HeaderFooter object is a member of the HeadersFooters collection.

## Behavior
Represents a single header or footer. TheHeaderFooterobject is a member of theHeadersFooterscollection.

## Example Usage
```vba
With ActiveSheet.PageSetup 
 .CenterHeader = "&D&T" 
 .OddAndEvenPagesHeaderFooter = False 
 .DifferentFirstPageHeaderFooter = False 
 .ScaleWithDocHeaderFooter = True 
 .AlignMarginsHeaderFooter = True 
End With
```