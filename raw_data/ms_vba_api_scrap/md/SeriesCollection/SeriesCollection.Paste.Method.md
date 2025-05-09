# SeriesCollection Paste Method

## Business Description
Pastes data from the Clipboard into the specified series collection.

## Behavior
Pastes data from the Clipboard into the specified series collection.

## Example Usage
```vba
Worksheets("Sheet1").Range("C1:C5").Copy 
Charts("Chart1").SeriesCollection.Paste
```