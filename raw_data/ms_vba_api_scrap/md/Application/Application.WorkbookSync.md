# Application.WorkbookSync Event (Excel)

## Business Description
The `WorkbookSync` event in Excel was used to detect when a workbook finished synchronizing changes with a shared or server copy. This allowed for custom actions after syncs, but the feature is now deprecated and should not be used in new solutions.

## Behavior
- **Deprecated**: This event remains in the object model for backward compatibility only. It is not recommended for use in new applications.
- **Purpose**: Historically, it was triggered after a workbook completed synchronization, allowing you to automate post-sync tasks.
- **Modern Usage**: For new projects, consider alternative collaboration and sync features in Excel.

## Example Usage
```vba
' Example: Responding to the WorkbookSync event (not recommended for new code)
Private Sub Application_WorkbookSync(ByVal Wb As Workbook, ByVal SyncEventType As Long)
    MsgBox "Workbook sync completed."
End Sub
```

**Tip:** Avoid using deprecated events in new solutions. Explore Excel's current collaboration tools for modern alternatives.
