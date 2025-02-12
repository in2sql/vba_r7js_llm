**Description:**
English: This code defines a custom function 'ADD' that adds two numbers, sets a formula in cell A1 using this function, clears all custom functions, and sets a message in cell A3 indicating the removal of custom functions.
Русский: Этот код определяет пользовательскую функцию 'ADD', которая складывает два числа, устанавливает формулу в ячейке A1 с использованием этой функции, очищает все пользовательские функции и устанавливает сообщение в ячейке A3, указывая на удаление пользовательских функций.

**OnlyOffice JavaScript:**
```javascript
// This example clears all added custom functions.
Api.AddCustomFunctionLibrary("LibraryName", function(){
    /**
     * Function that adds two arguments
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
// Set formula in A1 using the custom ADD function
oWorksheet.GetRange("A1").SetValue("=ADD(1, 2)");
// Clear all custom functions
Api.ClearCustomFunctions();
// Set message in A3 indicating custom functions were removed
oWorksheet.GetRange("A3").SetValue("All the custom functions were removed.");
```

**Excel VBA:**
```vba
' This code defines a custom function 'ADD', sets a formula in A1, clears the custom function, and sets a message in A3.

' Define the ADD function
Function ADD(first As Variant, second As Variant) As Variant
    ' Returns the sum of two arguments
    ADD = first + second
End Function

Sub ManageCustomFunctions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set formula in A1 using the ADD function
    ws.Range("A1").Formula = "=ADD(1, 2)"
    
    ' Clear the custom functions by removing the formula
    ws.Range("A1").Value = ws.Range("A1").Value
    
    ' Set message in A3 indicating custom functions were removed
    ws.Range("A3").Value = "All the custom functions were removed."
End Sub
```