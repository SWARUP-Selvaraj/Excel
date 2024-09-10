Attribute VB_Name = "ResetUsedRange"
Sub resetUsedRanges()
    ' This subroutine resets the used ranges of all worksheets in the active workbook.
    ' It forces Excel to recalculate the last used cell in each worksheet, which can help
    ' fix issues where Excel incorrectly identifies the used range due to lingering formatting
    ' or data in cells that appear empty.

    Dim ws As Worksheet ' Declare a variable to represent each worksheet

    ' Loop through all worksheets in the active workbook
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate ' Activate the current worksheet
        
        ' Access the UsedRange property twice to reset the internal used range of the worksheet
        ' This forces Excel to recalculate the last used cell in the worksheet
        ActiveSheet.UsedRange
        ActiveSheet.UsedRange
    Next ws ' Move to the next worksheet
End Sub