```javascript
// This script defines a custom function ADD that sums two arguments and sets it in cell A1.
// Этот скрипт определяет пользовательскую функцию ADD, которая складывает два аргумента и помещает результат в ячейку A1.

// Add a custom function library named "LibraryName"
Api.AddCustomFunctionLibrary("LibraryName", function(){
    /**
     * Function that adds two arguments
     * @customfunction
     * @param {any} first First argument.
     * @param {any} second Second argument.
     * @returns {any} Sum of first and second arguments.
    */
    Api.AddCustomFunction(function ADD(first, second) {
        return first + second; // Return the sum of first and second arguments
    });
});

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to use the ADD function with arguments 1 and 2
oWorksheet.GetRange('A1').SetValue('=ADD(1,2)');
```

```vba
' This script defines a custom function ADD that sums two arguments and sets it in cell A1.
' Этот скрипт определяет пользовательскую функцию ADD, которая складывает два аргумента и помещает результат в ячейку A1.

Sub DefineCustomFunctionAndSetValue()
    ' Set the value of cell A1 to use the ADD function with arguments 1 and 2
    ThisWorkbook.ActiveSheet.Range("A1").Formula = "=ADD(1,2)"
End Sub

' Function that adds two arguments
Function ADD(first As Variant, second As Variant) As Variant
    ADD = first + second ' Return the sum of first and second arguments
End Function
```