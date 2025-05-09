# Window Caption Property

## Business Description
Returns or sets a Variant value that represents the name that appears in the title bar of the document window.

## Behavior
Returns or sets aVariantvalue that represents the name that appears in the title bar of the document window.

## Example Usage
```vba
ActiveWorkbook.Windows(1).Caption= "Consolidated Balance Sheet" 
ActiveWorkbook.Windows("Consolidated Balance Sheet") _ 
 .ActiveSheet.Calculate
```