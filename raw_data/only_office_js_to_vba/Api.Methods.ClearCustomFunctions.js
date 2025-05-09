**Description / Описание**

English: This code clears all added custom functions, adds a custom function `ADD` that returns the sum of two arguments, sets cell A1 to use the `ADD` function with values 1 and 2, clears the custom functions, and sets cell A3 to indicate that all custom functions were removed.

Русский: Этот код очищает все добавленные пользовательские функции, добавляет пользовательскую функцию `ADD`, которая возвращает сумму двух аргументов, устанавливает значение ячейки A1 для использования функции `ADD` с значениями 1 и 2, очищает пользовательские функции и устанавливает значение ячейки A3, указывая на то, что все пользовательские функции были удалены.

```vba
' VBA Code Equivalent

' Define the custom function ADD
Function ADD(first As Variant, second As Variant) As Variant
    ' Returns the sum of two arguments
    ADD = first + second
End Function

' Subroutine to manage custom functions
Sub ManageCustomFunctions()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Set cell A1 to use the ADD function
    ws.Range("A1").Value = "=ADD(1, 2)"
    
    ' Clear custom functions (VBA does not support removing functions at runtime)
    ' Here, we simply inform the user that custom functions are cleared
    ws.Range("A3").Value = "All the custom functions were removed."
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example clears all added custom functions.
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
// Set cell A1 to use the ADD function with values 1 and 2
oWorksheet.GetRange("A1").SetValue("=ADD(1, 2)");
// Clear all custom functions
Api.ClearCustomFunctions();
// Inform the user that all custom functions have been removed
oWorksheet.GetRange("A3").SetValue("All the custom functions were removed.");
```