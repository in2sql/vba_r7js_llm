# Range Dependents Property

## Business Description
Returns a Range object that represents the range containing all the dependents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one dependent. Read-only Range object.

## Behavior
Returns aRangeobject that represents the range containing all the dependents of a cell. This can be a multiple selection (a union ofRangeobjects) if there's more than one dependent. Read-onlyRangeobject.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A1").Dependents.Select
```