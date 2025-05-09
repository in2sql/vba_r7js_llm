# Value Property

## Business Description
Returns the value of the specified cell. If the cell is empty, Value returns the value Empty (use the IsEmpty function to test for this case).

## Behavior
Returns the value of the specified cell. If the cell is empty, Value returns the value Empty (use the IsEmpty function to test for this case). If the Range object contains more than one cell, this property returns an array of values (use the IsArray function to test for this case). Read/write Variant.

## Example Usage
```vba
myChart.Application.DataSheet.Range("A1").Value= 3.14159
```