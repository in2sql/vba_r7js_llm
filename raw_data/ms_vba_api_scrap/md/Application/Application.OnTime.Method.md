# Application.OnTime method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Application.OnTime Now + TimeValue("00:00:15"), "my_Procedure"
```

## Parameters
- **EarliestTime**: Required
- **Procedure**: Required
- **LatestTime**: Optional
- **Schedule**: Optional

## Remarks
Use Now + TimeValue(time) to schedule something to be run when a specific amount of time (counting from now) has elapsed. Use TimeValue(time) to schedule something to be run a specific time.

## Example
No VBA example available.
