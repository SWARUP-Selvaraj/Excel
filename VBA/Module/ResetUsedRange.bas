Attribute VB_Name = "ResetUsedRange"
Sub resetUsedRanges()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ActiveSheet.UsedRange
        ActiveSheet.UsedRange
    Next ws

End Sub
