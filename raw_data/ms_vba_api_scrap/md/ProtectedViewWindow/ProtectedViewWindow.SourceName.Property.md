# ProtectedViewWindow SourceName Property

## Business Description
Returns the name of the source file that is open in the specified Protected View window. Read-only

## Behavior
Returns the name of the source file that is open in the specifiedProtected Viewwindow. Read-only

## Example Usage
```vba
MsgBox ActiveProtectedViewWindow.SourcePath & "\" _ 
 & ActiveProtectedViewWindow.SourceName
```