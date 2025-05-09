# Range Precedents Property

## Business Description
Returns a Range object that represents all the precedents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one precedent. Read-only.

## Behavior
Returns aRangeobject that represents all the precedents of a cell. This can be a multiple selection (a union ofRangeobjects) if there's more than one precedent. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A1").Precedents.Select
```