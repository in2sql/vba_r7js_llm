# VPageBreak Location Property

## Business Description
Returns or sets the cell (a Range object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. Read/write Range.

## Behavior
Returns or sets the cell (aRangeobject) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. Read/writeRange.

## Example Usage
```vba
Worksheets(1).VPageBreaks(1).Location= Worksheets(1).Range("e5")
```