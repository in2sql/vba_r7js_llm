# How to: Change the Color of the Horizontal Axis of a Sparkline

## Business Description
You can change the color of the horizontal axis of a sparkline by using the Color property of the SparkColor object.

## Behavior
You can change the color of the horizontal axis of a sparkline by using theColorproperty of theSparkColorobject. The following code example iterates through three sparkline groups and sets the color of the horizontal axis equal to the fill color in cell A8. This example requires three sparkline groups starting in cells A2, B2, and C2. Cell A8 must be filled with the color that you want to use for the color of the horizontal axis. This example usesColorproperty of theInteriorobject to get the color of cell A8.

## Example Usage
```vba
Sub AxisColor()
    'The sparkline group
    Dim oSparkGroup As SparklineGroup
    'Loop through the sparkline groups on the sheet
    For Each oSparkGroup In Range("A2:C2").SparklineGroups
        'Show the axis
        oSparkGroup.Axes.Horizontal.Axis.Visible = True
        'Set the color of the axis to the color of cell A8
        oSparkGroup.Axes.Horizontal.Axis.Color.Color = Range("A8").Interior.Color
    Next oSparkGroup
End Sub
```