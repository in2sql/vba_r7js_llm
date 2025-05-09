# Range.ReadingOrder Property (Excel)

## Business Description
The `ReadingOrder` property in Excel determines the direction in which text is displayed within a cell or range—left-to-right, right-to-left, or based on the context. This is especially useful for international workbooks that include languages with different writing directions.

## Behavior
- **Get or Set**: You can read or change the reading order for a specified range.
- **Options**:
  - `xlRTL`: Right-to-left (used for languages like Arabic or Hebrew)
  - `xlLTR`: Left-to-right (used for most Western languages)
  - `xlContext`: Uses the context of the first strong character in the cell
- **Use Case**: Ensures that data is presented correctly for users regardless of language.

## Example Usage
```vba
' Set reading order to right-to-left for a range
Range("A1:B2").ReadingOrder = xlRTL

' Set reading order to left-to-right
Range("A1:B2").ReadingOrder = xlLTR

' Use context-based reading order
Range("A1:B2").ReadingOrder = xlContext
```

**Tip:** Adjust the reading order to make your spreadsheets accessible and user-friendly for international teams.
