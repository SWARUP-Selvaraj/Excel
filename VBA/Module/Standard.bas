' https://superuser.com/questions/217504/is-there-a-list-of-windows-special-directories-shortcuts-like-temp
Function tableExist(tableDict As Scripting.Dictionary) As Boolean
    ' Verifies if a Table Exist in a spreadsheet and returns a boolean.
    On Error Resume Next
    tableExist = CBool(tableDict("_Sheet_").Range(tableDict("_Table_")).Rows.Count)
    If Err.Number <> 0 Then
        tableExist = False
        Err.Number = 0
    End If
End Function
Function buildTable(tableDict As Scripting.Dictionary, _
    Optional ByVal tableFormat As String = "TableStyleMedium9")
    ' Creates a Table based on the template specified in the tableDict dictionary.
    ' Returns False if Table was created without encountering errors.
    Dim rng As Range
    Dim i, r, c As Long
    On Error Resume Next
    Set sht = tableDict("_Sheet_")
    If tableExist(tableDict) Then
        ClearTable tableDict
    End If
    r = tableDict("_Row_")
    c = tableDict("_Column_")
    For i = 5 To tableDict.Count - 1
        tableDict("_Sheet_").Cells(r, c + tableDict.Items(i)).Value = CStr(tableDict.Keys(i))
    Next
    Set rng = tableDict("_Sheet_").Range(tableDict("_Sheet_").Cells(r, c + 1), _
        tableDict("_Sheet_").Cells(r, c + tableDict.Count - 5))
    rng.Worksheet.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = tableDict("_Table_")
    rng.Worksheet.ListObjects(tableDict("_Table_")).TableStyle = tableFormat
    Set tableDict("_Range_") = sht.Range(tableDict("_Table_"))
    If Err.Number <> 0 Then
        buildTable = False
        Err.Number = 0
    Else
        buildTable = True
    End If
    If Not buildTable Then messageBox "Creation of Table " & tableDict("_Table_") & _
        " failed!" & Err.Number, "", vbCritical
End Function
Function messageBox(ByVal msg As String, Optional ByVal mtitle As String = "", _
    Optional ByVal msty As VbMsgBoxStyle = vbInformation) As Integer
    ' Displays a Standard Message Box
    messageBox = MsgBox(msg, msty, AppName & IIf(Len(mtitle) > 0, " > " & mtitle, ""))
End Function
Public Function item_lookup(ByVal target As Variant, lkuprng As Range, _
    ByVal lkupcol As Integer, Optional ByVal errval As Variant = "True") As Variant
    ' Performs VLOOKUP without raising error.  When item is not found returns the errval.
    On Error Resume Next
    If errval = "True" Then
        item_lookup = target
    Else
        item_lookup = errval
    End If
    item_lookup = Application.WorksheetFunction.VLookup(target, lkuprng, lkupcol, False)
End Function
Public Function item_match(ByVal target As Variant, lkuprng As Range, _
    Optional ByVal method As Integer = 0)
    ' Performs MATCH without raising error.  When item is not found returns -1.
    On Error Resume Next
    item_match = -1
    item_match = Application.WorksheetFunction.Match(target, lkuprng, method)
End Function
Sub applySort(srtrng As ListObject, Optional ByVal srtHead As Variant = xlYes, _
    Optional ByVal srtCase As Boolean = False, Optional ByVal srtOrient As Variant = _
    xlTopToBottom, Optional ByVal srtMethod As Variant = xlPinYin)
    srtrng.Sort.Header = srtHead
    srtrng.Sort.MatchCase = srtCase
    srtrng.Sort.Orientation = srtOrient
    srtrng.Sort.SortMethod = srtMethod
    srtrng.Sort.Apply
    srtrng.Sort.SortFields.Clear
End Sub
Sub ClearTable(tableDict As Scripting.Dictionary)
    ' Deletes a Table and resets the used range.
    On Error Resume Next
    tableDict("_Sheet_").Range(tableDict("_Table_") & "[#All]").Clear
    tableDict("_Sheet_").Activate
    ActiveSheet.UsedRange
End Sub
Function MakeUNC(ByVal path As String, Optional ByVal suffix As Boolean = False) As String
    ' Prepends Universal Naming Convention to any Path.
    MakeUNC = IIf(Left(path, 4) = "\\?\", path, "\\?\" & path)
    If suffix Then
        MakeUNC = IIf(Right(MakeUNC, 1) = "\", MakeUNC, MakeUNC & "\")
    Else
        MakeUNC = IIf(Right(MakeUNC, 1) = "\", Left(MakeUNC, Len(MakeUNC) - 1), MakeUNC)
    End If
End Function
Function UnMakeUNC(ByVal path As String, Optional ByVal suffix As Boolean = False) As String
    ' Removes Universal Naming Convention from any path.
    UnMakeUNC = IIf(Left(path, 4) = "\\?\", Right(path, Len(path) - 4), path)
    If suffix Then
        UnMakeUNC = IIf(Right(UnMakeUNC, 1) = "\", UnMakeUNC, UnMakeUNC & "\")
    Else
        UnMakeUNC = IIf(Right(UnMakeUNC, 1) = "\", Left(UnMakeUNC, Len(UnMakeUNC) - 1), UnMakeUNC)
    End If
End Function
Function FileSelect(ByVal title As String, Optional ByVal initial As String = "C:\", _
    Optional ByVal filter As String = "False", Optional ByVal extn As String = "*.*;") As String
    ' Pops up a File Picker and returns the path of the selected file.
    If initial = "" Then initial = "C:\"
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = initial
        .title = title
        .AllowMultiSelect = False
        If filter <> "False" Then
            .Filters.Add filter, extn, 1
        End If
        .Show
        If Not .SelectedItems.Count = 0 Then
            FileSelect = .SelectedItems(1)
        End If
    End With
End Function
Function FolderSelect(ByVal title As String, Optional ByVal initial As String = "C:\") As String
    ' Pops up a Folder Picker and returns the path of the selected folder.
    If initial = "" Then initial = "C:\"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = initial
        .title = title
        .Show
        If Not .SelectedItems.Count = 0 Then
            FolderSelect = .SelectedItems(1)
        End If
    End With
End Function
Sub RobustCopy(ByVal src, ByVal tar, ByVal nme)
    ' Initiates a file copy using Windows Robust Copier
    src = UnMakeUNC(src, False)
    tar = UnMakeUNC(tar, False)
    oShell.Run "cmd /c ""robocopy """ & src & """ """ & tar & """ """ & nme & """""", 0, True
End Sub
Function genTimeStamp(Optional ByVal sd As String = "_", Optional ByVal st As String = "_", _
    Optional ByVal s As String = "-", Optional ByVal prefix As String = "TS-", _
    Optional ByVal postfix as String = "") As String
    ' Returns a current timestamp string <prefix>DD<sd>MM<sd>YYYY<s>HH<st>MM<st>SS<postfix>
    genTimeStamp = prefix & Format(Year(Now), "0000") & sd & Format(Month(Now), "00") & sd & _
        Format(Day(Now), "00") & s & Format(Hour(Now), "00") & st & _
        Format(Minute(Now), "00") & st & Format(Second(Now), "00") & postfix
End Function
Sub BuildFullPath(ByVal fullpath)
    ' Recursively checks and builds the given path
    fullpath = MakeUNC(fullpath, True)
    If Not oFSO.FolderExists(fullpath) Then
        BuildFullPath oFSO.GetParentFolderName(fullpath)
        oFSO.CreateFolder fullpath
    End If
End Sub
Function getSize(ByVal filePath As String, _
    Optional ByVal searchString As String = "Size:") As Variant
    ' Returns the folder size by parsing the report generated by Disk Usage.
    Dim txtStream As Object
    Dim tmpLine As String
    getSize = -1
    Set txtStream = oFSO.OpenTextFile(filePath, 1, False)
    Do While Not txtStream.AtEndOfStream
        tmpLine = txtStream.ReadLine
        If InStr(1, tmpLine, searchString, vbBinaryCompare) > 0 Then
            getSize = CDbl(Replace(Replace(Replace(Replace(tmpLine, searchString, ""), _
                " ", ""), ",", ""), "bytes", "")) / 1048576
        End If
    Loop
    txtStream.Close
    If getSize > 0 Then
        On Error Resume Next
        Kill filePath
    End If
End Function
Function pathJoin(ByVal base As String, ByVal addon As String) As String
    ' Returns a joined path by appending an addon path string to a Base path.
    base = base & IIf(Right(base, 1) = "\", "", "\")
    If Len(addon) > 0 Then
        addon = IIf(Left(addon, 1) = "\", mid(addon, 2, Len(addon) - 1), addon)
    End If
    If Len(addon) > 0 Then
        addon = IIf(Right(addon, 1) = "\", addon, addon & "\")
    End If
    pathJoin = base & addon
End Function
Function GetRelativePath(ByVal basepath As String, ByVal abspath As String) As String
    ' Returns a relative path by removing the base path from the absolute path.
    basepath = MakeUNC(basepath, True)
    abspath = MakeUNC(abspath, False)
    GetRelativePath = Right(abspath, Len(abspath) - Len(basepath))
End Function
Sub exeCmd(ByVal cmd As String, Optional ByVal visible As Integer = 0, _
    Optional ByVal wait As Boolean = False, Optional ByVal quot As String = """")
    ' Executes the given cmd in windows command prompt.
    oShell.Run "cmd.exe /C " & quot & cmd & quot " & exit", visible, wait
End Sub
Sub openFolder(ByVal FolderName As String, Optional ByVal focus = vbNormalFocus)
    ' Opens the specified folder in windows explorer.
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO.FolderExists(FolderName) Then
        Shell Environ("SystemRoot") & "\explorer.exe """ & FolderName & """", focus
    Else
        messageBox """" & FolderName & """ does NOT EXISTS!", "Invalid Folder", vbCritical
    End If
End Sub
Sub RemoveTree(ByVal path As String, Optional ByVal warn As Integer = 2)
    ' !!! WARNING !!! Use with EXTREME CAUTION.
    ' Recursively remove all files and subfolders from the specified folder.
    ' File removal is PERMANANT and CANNOT be reversed or recovered later.
    path = UnMakeUNC(path, True)
    If warn <> 1 Then
        warn = messageBox("Are you sure you want to delete the following location?" & _
            vbNewLine & vbNewLine & path, "Confirm Folder Deletion", vbOKCancel)
    End If
    If warn = 1 Then exeCmd "cd """ & path & _
        """ & DEL /F/Q/S *.* > NUL & cd .. & RMDIR /Q/S """ & path & """", 0, True, ""
End Sub
Function InsertRecord(tableDict As Scripting.Dictionary) As Variant
    ' When the first cell of the Last Row is not empty, inserts a row to
    ' the end of the specified table.
    If Len(tableDict("_Sheet_").Range(tableDict("_Table_")).Cells(tableDict("_Sheet_" _
        ).Range(tableDict("_Table_")).Rows.Count, 1).Value) <> 0 Then
        tableDict("_Sheet_").Range(tableDict("_Table_")).ListObject.ListRows.Add _
            AlwaysInsert:=True
    End If
    Set tableDict("_Range_") = tableDict("_Sheet_").Range(tableDict("_Table_"))
    InsertRecord = tableDict("_Sheet_").Range(tableDict("_Table_")).Rows.Count
End Function
Sub applyLower(rng As Range)
    ' Converts all the text in the given range to Lower Case.
    Dim cel As Range
    For Each cel In rng.Cells
        cel.Value = LCase(cel.Value)
    Next
End Sub
Sub trimToChar(rng As Range, ByVal nChar As Integer)
    ' Truncates the text longer than nChar from all cells in the specified range.
    Dim cel As Range
    For Each cel In rng.Cells
        cel.Value = LCase(Left(cel.Value, xlFunct.Min(nChar, Len(cel.Value))))
    Next
End Sub
Function getTableRange(tableDict As Scripting.Dictionary) As Range
    ' Returns a Range object corresponding to a Table
    Set tableDict("_Range_") = tableDict("_Sheet_").Range(tableDict("_Table_"))
    Set getTableRange = tableDict("_Range_")
End Function
Function fetchFile(ByVal src As String, ByVal dest As String) As Boolean
    ' Downloads a file from the Internet.
    Dim hReq As Variant
    Dim oStream As Variant
    Set hReq = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    fetchFile = False
    With hReq
        .Open "POST", src, False
        .send
    End With
    If hReq.Status = 200 Then

        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write hReq.responseBody
        oStream.SaveToFile dest, 2
        oStream.Close
        If Err.Number = 0 Then fetchFile = True
    End If
End Function