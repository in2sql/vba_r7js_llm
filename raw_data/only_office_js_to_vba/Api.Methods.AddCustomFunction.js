---
**Description:**

This code defines a custom function `ADD` that adds two numbers and sets the formula `=ADD(1,2)` in cell A1.

Этот код определяет пользовательскую функцию `ADD`, которая складывает два числа, и устанавливает формулу `=ADD(1,2)` в ячейку A1.

---

```vba
' VBA Code Equivalent

' Define a custom function ADD that adds two numbers
Function ADD(first As Variant, second As Variant) As Variant
    ADD = first + second
End Function

' Subroutine to set the formula in cell A1
Sub SetCustomFunctionFormula()
    ' Set the formula in cell A1 to =ADD(1,2)
    ThisWorkbook.ActiveSheet.Range("A1").Formula = "=ADD(1,2)"
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example calculates custom function result.

// Add a custom function library named "LibraryName"
Api.AddCustomFunctionLibrary("LibraryName", function(){
    /**
     * Function that adds two numbers
     * @customfunction
     * @param {any} first First argument.
     * @param {any} second Second argument.
     * @returns {any} Sum of first and second arguments.
    */
    Api.AddCustomFunction(function ADD(first, second) {
        return first + second;
    });
});

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the formula '=ADD(1,2)' in cell A1
oWorksheet.GetRange('A1').SetValue('=ADD(1,2)');
```