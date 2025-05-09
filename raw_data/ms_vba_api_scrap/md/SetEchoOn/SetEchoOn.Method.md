# SetEchoOn Method

## Business Description
Returns a Chart object.

## Behavior
Returns a Chart object.

## Example Usage
```vba
Sub UseEchoOn() 
 
 Dim grpOne As Graph.Chart 
 
 Set grpOne = Application.ActiveSheet.OLEObjects(1).Object 
 
 grpOne.SetEchoOnEnd Sub
```