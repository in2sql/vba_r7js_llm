# Slicer Cut Method

## Business Description
Cuts the specified slicer and copies it to the clipboard.

## Behavior
Cuts the specified slicer and copies it to the clipboard.

## Example Usage
```vba
ActiveSheet.Shapes.Range(Array("Customer")).Select 
Selection.CutActiveSheet.Paste
```