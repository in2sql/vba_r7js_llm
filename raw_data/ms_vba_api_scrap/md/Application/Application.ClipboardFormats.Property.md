# Application.ClipboardFormats property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
aFmts = Application.ClipboardFormats 
For Each fmt In aFmts 
 If fmt = xlClipboardFormatRTF Then 
 MsgBox "Clipboard contains rich text" 
 End If 
Next
```

## Parameters
- **Index**: Optional

## Remarks
This property returns an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element of the array with one of the XlClipboardFormat constants.

## Example
No VBA example available.
