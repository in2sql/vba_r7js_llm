# SortFields Object

## Business Description
The SortFields collection is a collection of SortField objects. It allows developers to store a sort state on workbooks, lists, and autofilters.

## Behavior
TheSortFieldscollection is a collection ofSortFieldobjects. It allows developers to store a sort state on workbooks, lists, and autofilters.

## Example Usage
```vba
ActiveWorksheet.SortFields.Add Key:=Range("A1"), Order:=xlDescending 
ActiveWorksheet.SortFields.Add Key:=Range("B1"), Order:=xlDescending 
ActiveWorksheet.SortFields.Sort Header:=xlGuess
```