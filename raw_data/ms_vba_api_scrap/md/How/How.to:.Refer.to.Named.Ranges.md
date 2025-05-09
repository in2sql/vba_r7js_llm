# How to: Refer to Named Ranges

## Business Description
Ranges are easier to identify by name than by A1 notation. To name a selected range, click the name box at the left end of the formula bar, type a name, and then press ENTER.

## Behavior
Ranges are easier to identify by name than by A1 notation. To name a selected range, click the name box at the left end of the formula bar, type a name, and then press ENTER.

## Example Usage
```vba
Sub FormatRange() 
    Range("MyBook.xls!MyRange").Font.Italic = True 
End Sub
```