# ProtectedViewWindow SourcePath Property

## Business Description
Returns the path of the source file that is open in the specified Protected View window. Read-only

## Behavior
Returns the path of the source file that is open in the specifiedProtected Viewwindow. Read-only

## Example Usage
```vba
MsgBox ActiveProtectedViewWindow.SourcePath& Application.PathSeparator _ 
 & ActiveProtectedViewWindow.SourceName
```