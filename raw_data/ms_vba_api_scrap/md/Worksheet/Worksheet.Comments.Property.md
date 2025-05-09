# Worksheet Comments Property

## Business Description
Returns a Comments collection that represents all the comments for the specified worksheet. Read-only.

## Behavior
Returns aCommentscollection that represents all the comments for the specified worksheet. Read-only.

## Example Usage
```vba
For Each c in ActiveSheet.Comments 
 If c.Author= "Jean Selva" Then c.Delete 
Next
```