# Worksheet Change Event

## Business Description
Occurs when cells on the worksheet are changed by the user or by an external link.

## Behavior
Occurs when cells on the worksheet are changed by the user or by an external link.

## Example Usage
```vba
Private Sub Worksheet_Change(ByVal Target as Range) 
    Target.Font.ColorIndex = 5 
End Sub
```