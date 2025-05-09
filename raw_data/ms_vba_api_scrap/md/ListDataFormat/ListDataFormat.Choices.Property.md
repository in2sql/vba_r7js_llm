# ListDataFormat Choices Property

## Business Description
Returns an Array of String values that contains the choices offered to the user by the ListLookUp, ChoiceMulti, and Choice data types of the DefaultValue property. Read-only Variant.

## Behavior
Returns anArrayofStringvalues that contains the choices offered to the user by theListLookUp,ChoiceMulti, andChoicedata types of theDefaultValueproperty. Read-onlyVariant.

## Example Usage
```vba
Sub PrintChoices() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.ChoicesEnd Sub
```