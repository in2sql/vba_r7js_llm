# LeaderLines Object

## Business Description
Leader lines in Excel charts visually connect data labels to their corresponding data points, making it easier to understand which label belongs to which value, especially when labels are positioned away from the data points.

## Behavior
- **Purpose**: Automatically draws lines from data labels to their respective data points when needed, improving chart clarity.
- **Not a Collection**: The LeaderLines object represents all leader lines in a chart group; you cannot address individual lines.
- **Usage**: Commonly used in pie and doughnut charts, or any chart where data labels are moved for readability.

## Example Usage
```vba
' Enable leader lines for a chart series
ActiveChart.SeriesCollection(1).HasLeaderLines = True

' Access the LeaderLines object (for formatting, etc.)
Set leaderLines = ActiveChart.SeriesCollection(1).LeaderLines

' Note: You cannot manipulate individual leader lines.
```

**Tip:** Use leader lines to keep your charts easy to read when data labels are moved away from data points.
