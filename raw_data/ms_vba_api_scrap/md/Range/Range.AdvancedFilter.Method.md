# Range AdvancedFilter Method

## Business Description
Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.

## Behavior
Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.

## Example Usage
```vba
Range("Database").AdvancedFilter_ 
 Action:=xlFilterInPlace, _ 
 CriteriaRange:=Range("Criteria")
```