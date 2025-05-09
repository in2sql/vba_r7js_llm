**Description / Описание**  
This script adds a custom function to the worksheet, sets a cell with the function, removes the function, and updates another cell.  
Этот скрипт добавляет пользовательскую функцию на лист, устанавливает ячейку с функцией, удаляет функцию и обновляет другую ячейку.

```vba
' VBA Code Equivalent

' Module: CustomFunctions

' Function that adds two numbers
Function ADD(first As Variant, second As Variant) As Variant
    ADD = first + second
End Function

' Subroutine to add the function, set cell values, and remove the function
Sub ManageCustomFunction()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set cell A1 with the custom ADD function
    ws.Range("A1").Formula = "=ADD(1, 2)"
    
    ' Remove the ADD function by deleting the function from the module
    ' Note: Deleting VBA code programmatically requires access to the VBA project
    ' and appropriate permissions. This example assumes the function is in a standard module.
    Dim VBProj As Object
    Dim VBComp As Object
    Dim CodeMod As Object
    Dim LineNum As Long
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents("Module1") ' Adjust module name as necessary
    Set CodeMod = VBComp.CodeModule
    
    ' Find the line where the ADD function starts
    LineNum = CodeMod.ProcStartLine("ADD", vbext_pk_Proc)
    
    ' Delete the entire ADD function
    CodeMod.DeleteLines LineNum, CodeMod.ProcCountLines("ADD", vbext_pk_Proc)
    
    ' Set cell A3 with a message indicating the function was removed
    ws.Range("A3").Value = "The ADD custom function was removed."
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example clears the current custom function.
Api.AddCustomFunctionLibrary("LibraryName", function(){
    /**
     * Function that returns the sum of two arguments
     * @customfunction
     * @param {any} first First argument.
     * @param {any} second Second argument.
     * @returns {any} Sum of first and second arguments.
     */
    Api.AddCustomFunction(function ADD(first, second) {
        return first + second;
    });
});

var oWorksheet = Api.GetActiveSheet();

// Set cell A1 with the custom ADD function
oWorksheet.GetRange("A1").SetValue("=ADD(1, 2)");

// Remove the ADD custom function
Api.RemoveCustomFunction("ADD");

// Set cell A3 with a message indicating the function was removed
oWorksheet.GetRange("A3").SetValue("The ADD custom function was removed.");
```