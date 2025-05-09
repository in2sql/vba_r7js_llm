# Application.MacroOptions method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Function TestMacro() 
    MsgBox ActiveWorkbook.Name 
End Function 
 
Sub AddUDFToCustomCategory() 
    Application.MacroOptions Macro:="TestMacro", Category:="My Custom Category" 
End Sub
```

## Parameters
- **Macro**: Optional
- **Description**: Optional
- **HasMenu**: Optional
- **MenuText**: Optional
- **HasShortcutKey**: Optional
- **ShortcutKey**: Optional
- **Category**: Optional
- **StatusBar**: Optional
- **HelpContextID**: Optional
- **HelpFile**: Optional
- **ArgumentDescriptions**: Optional

## Remarks
The following table lists which integers are mapped to the built-in categories that can be used in the Category parameter.

## Example
```vba
Function TestMacro() 
    MsgBox ActiveWorkbook.Name 
End Function 
 
Sub AddUDFToCustomCategory() 
    Application.MacroOptions Macro:="TestMacro", Category:="My Custom Category" 
End Sub
```

