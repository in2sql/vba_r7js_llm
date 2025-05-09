# Range End Property

## Business Description
Returns a Range object that represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW. Read-only Range object.

## Behavior
Returns aRangeobject that represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW. Read-onlyRangeobject.

## Example Usage
```vba
Range("B4").End(xlUp).Select
```