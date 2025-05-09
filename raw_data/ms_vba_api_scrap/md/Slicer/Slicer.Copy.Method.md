# Slicer Copy Method

## Business Description
Copies the specified slicer to the clipboard.

## Behavior
Copies the specified slicer to the clipboard.

## Example Usage
```vba
ActiveSheet.Shapes.Range(Array("Customer")).Select 
Selection.CopyActiveSheet.Paste
```