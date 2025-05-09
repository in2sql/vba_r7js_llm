# Range SortSpecial Method

## Business Description
Uses East Asian sorting methods to sort the range, a PivotTable report, or uses the method for the active region if the range contains only one cell. For example, Japanese sorts in the order of the Kana syllabary.

## Behavior
Uses East Asian sorting methods to sort the range, a PivotTable report, or uses the method for the active region if the range contains only one cell. For example, Japanese sorts in the order of the Kana syllabary.

## Example Usage
```vba
Sub SpecialSort() 
 
 Application.Range("A1:A5").SortSpecialSortMethod:=xlPinYin 
 
End Sub
```