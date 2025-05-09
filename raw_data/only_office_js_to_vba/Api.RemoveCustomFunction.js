# Description / Описание

**English:**  
This example clears the current custom function library, adds a new custom function `ADD` that returns the sum of two arguments, sets a formula in cell A1 to use the `ADD` function, removes the `ADD` custom function, and sets a message in cell A3 indicating that the `ADD` function was removed.

**Russian:**  
Этот пример очищает текущую библиотеку пользовательских функций, добавляет новую пользовательскую функцию `ADD`, которая возвращает сумму двух аргументов, устанавливает формулу в ячейку A1 для использования функции `ADD`, удаляет пользовательскую функцию `ADD` и устанавливает сообщение в ячейку A3, указывающее, что функция `ADD` была удалена.

```javascript
// Clear the current custom function library and add a new 'ADD' function
Api.AddCustomFunctionLibrary("LibraryName", function(){
    /**
     * Returns the sum of two arguments
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
// Set the formula in cell A1 to use the 'ADD' function
oWorksheet.GetRange("A1").SetValue("=ADD(1, 2)");

// Remove the 'ADD' custom function from the library
Api.RemoveCustomFunction("add");
// Set a message in cell A3 indicating the 'ADD' function was removed
oWorksheet.GetRange("A3").SetValue("The ADD custom function was removed.");
```

```vba
' Clear the current custom function library and add a new 'ADD' function
Api.AddCustomFunctionLibrary "LibraryName", Sub()
    ' Function that returns the sum of two arguments
    Api.AddCustomFunction "ADD", Function(first As Variant, second As Variant) As Variant
        ADD = first + second
    End Function
End Sub

Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet
' Set the formula in cell A1 to use the 'ADD' function
oWorksheet.GetRange("A1").SetValue "=ADD(1, 2)"
' Remove the 'ADD' custom function from the library
Api.RemoveCustomFunction "add"
' Set a message in cell A3 indicating the 'ADD' function was removed
oWorksheet.GetRange("A3").SetValue "The ADD custom function was removed."
```