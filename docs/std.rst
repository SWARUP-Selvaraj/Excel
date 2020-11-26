Standard Module
===============

Standard Module contains Functions and Sub Procedures for some of the frequently
used processes


Working with Tables
-------------------

tableExists
^^^^^^^^^^^

.. function:: tableExists(tableDict As Scripting.Dictionary, )

   Verifies if a Table Exist in a spreadsheet and returns a boolean

   :type: `Function Procedure`_
   :parameter:
      tableDict (`Dictionary`_) - Dictionary representing a table
   :returns: True if table exists else False.
   :return type: `Boolean`_


buildTable
^^^^^^^^^^

.. function:: buildTable(tableDict As Scripting.Dictionary, Optional ByVal tableFormat As String = "TableStyleMedium9")

   Creates a Table based on the specifications of tableDict

   :type: `Function Procedure`_
   :parameters:
      * tableDict (`Dictionary`_) - Dictionary representing a table
      * tableFormat (`String`_) - String representing a defined `Table Style <https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyles>`_
   :returns: True if encountered errors else False.
   :return type: `Boolean`_


ClearTable
^^^^^^^^^^

.. function:: ClearTable(tableDict As Scripting.Dictionary)

   Deletes a Table and resets the used range.

   :type: `Sub Procedure`_
   :parameters:
      tableDict (`Dictionary`_) - Dictionary representing a table

InsertRecord
^^^^^^^^^^^^

.. function:: InsertRecord(tableDict As Scripting.Dictionary)

   When the first cell of the Last Row is not empty, inserts a row to the end of the specified table.

   :type: `Function Procedure`_
   :parameters:
      tableDict (`Dictionary`_) - Dictionary representing a table
   :returns: Number of rows present in the table.
   :return type: `Variant`_

getTableRange
^^^^^^^^^^^^^

.. function:: getTableRange(tableDict As Scripting.Dictionary)

   Returns a Range object corresponding to a Table

   :type: `Function Procedure`_
   :parameters:
      tableDict (`Dictionary`_) - Dictionary representing a table
   :returns: Range object representing the table range.
   :return type: `Range`_

applySort
^^^^^^^^^

.. function:: applySort(srtrng As ListObject, Optional ByVal srtHead As Variant = xlYes, Optional ByVal srtCase As Boolean = False, Optional ByVal srtOrient As Variant = xlTopToBottom, Optional ByVal srtMethod As Variant = xlPinYin)

   Removes `Sort`_ from a given table .

   :type: `Sub Procedure`_
   :parameters:
      * srtrng (`ListObject`_) - Table to operate on
      * srtHead (`Variant`_) - Heading present in the range
      * srtCase (`Boolean`_) - Match case while sorting
      * srtOrient (`Variant`_) - Sorting order
      * srtMethod (`Variant`_) - Sorting Method

Theming
-------

messageBox
^^^^^^^^^^

.. function:: messageBox(ByVal msg As String, Optional ByVal mtitle As String = "", Optional ByVal msty As VbMsgBoxStyle = vbInformation)

   Displays a standard formatted Message Box.

   :type: `Function Procedure`_
   :parameters:
      * msg (`String`_) - Message to be displayed
      * mtitle (`String`_) - Message Heading to be displayed at Title bar
      * msty (`MsgBox Constant`_) - Message Box Constant
   :returns: True if encountered errors else False.
   :return type: `Boolean`_

Lookups
-------

item_lookup
^^^^^^^^^^^

.. function:: item_lookup(ByVal target As Variant, lkuprng As Range, ByVal lkupcol As Integer, Optional ByVal errval As Variant = "True")

   Performs VLOOKUP without raising error.  When item is not found returns the errval.

   :type: `Function Procedure`_
   :parameters:
      * target (`Variant`_) - Value to be lookedup in the range
      * lkuprng (`Range`_) - Range on which to perform the Lookup
      * lkupcol (`Integer`_) - Column to be returned from the lookup range
      * errval (`Variant`_) - Value to be returned when a match is not found
   :returns: True if encountered errors else False.
   :return type: `Boolean`_

item_match
^^^^^^^^^^

.. function:: item_match(ByVal target As Variant, lkuprng As Range, Optional ByVal method As Integer = 0)

   Performs MATCH without raising error.  When item is not found returns -1.

   :type: `Function Procedure`_
   :parameters:
      * target (`Variant`_) - Value to be matched in the range
      * lkuprng (`Range`_) - Range on which to perform the Lookup
      * method (`Integer`_) - Column to be returned from the lookup range
   :returns: True if encountered errors else False.
   :return type: `Boolean`_

Working with Paths
------------------

MakeUNC
^^^^^^^

.. function:: MakeUNC(ByVal path As String, Optional ByVal suffix As Boolean = False)

   Prepends `Universal Naming Convention`_ to any Path.

   :type: `Function Procedure`_
   :parameters:
      * path (`String`_) - Value to be matched in the range
      * suffix (`Boolean`_) - Boolean to require the presence of trailing \\ in the returned path
   :returns: UNC path corresponding to the given path.
   :return type: `String`_

UnMakeUNC
^^^^^^^^^

.. function:: UnMakeUNC(ByVal path As String, Optional ByVal suffix As Boolean = False)

   Removes `Universal Naming Convention`_ from any path.

   :type: `Function Procedure`_
   :parameters:
      * path (`String`_) - Value to be matched in the range
      * suffix (`Boolean`_) - Boolean to require the presence of trailing \\ in the returned path
   :returns: Removes UNC from the given path.
   :return type: `String`_

pathJoin
^^^^^^^^

.. function:: pathJoin(ByVal base As String, ByVal addon As String)

   Returns a joined path by appending an addon path string to a Base path.

   :type: `Function Procedure`_
   :parameters:
      * base (`String`_) - Base path (with or without trailing \\)
      * addon (`String`_) - Path to be added (with or without leading and trailing \\)
   :returns: Joined path (base + addon) with appropriate \\ .
   :return type: `String`_

GetRelativePath
^^^^^^^^^^^^^^^

.. function:: GetRelativePath(ByVal basepath As String, ByVal abspath As String)

   Returns a relative path by removing the base path from the absolute path.

   :type: `Function Procedure`_
   :parameters:
      * basepath (`String`_) - Base path to be removed from Full Path
      * abspath (`String`_) - Absolute Path to be made relative
   :returns: Relative path by removing basepath from abspath.
   :return type: `String`_

FileSelect
^^^^^^^^^^

.. function:: FileSelect(ByVal title As String, Optional ByVal initial As String = "C:\", Optional ByVal filter As String = "False", Optional ByVal extn As String = "*.*;")

   Pops up a File Picker and returns the path of the selected file.

   :type: `Function Procedure`_
   :parameters:
      * title (`String`_) - Titlebar content for the file selection dialogbox
      * initial (`String`_) - Default location
      * filter (`String`_) - Label of the Filter (Eg.: Excel Files)
      * extn (`String`_) - Extensions to filter (Eg.: \*.xlsx)
   :returns: Path of the selected file.
   :return type: `String`_

FolderSelect
^^^^^^^^^^^^

.. function:: FolderSelect(ByVal title As String, Optional ByVal initial As String = "C:\")

   Pops up a Folder Picker and returns the path of the selected folder.

   :type: `Function Procedure`_
   :parameters:
      * title (`String`_) - Titlebar content for the folder selection dialogbox
      * initial (`String`_) - Default location
   :returns: Path of the selected folder.
   :return type: `String`_

File System Operations
----------------------

RobustCopy
^^^^^^^^^^

.. function:: RobustCopy(ByVal src, ByVal tar, ByVal nme)

   Initiates a file copy using `Windows Robust Copier`_

   :type: `Sub Procedure`_
   :parameters:
      * src (`String`_) - Source Folder of the file to be copied
      * tar (`String`_) - Target Folder of the file to be copied
      * nme (`String`_) - Name of the file to be copied

BuildFullPath
^^^^^^^^^^^^^

.. function:: BuildFullPath(ByVal fullpath)

   Recursively checks and builds the given path.

   :type: `Sub Procedure`_
   :parameters:
      fullpath (`String`_) - The folder path to be built.

getSize
^^^^^^^

.. function:: getSize(ByVal filePath As String, Optional ByVal searchString As String = "Size:")

   Returns the folder size by parsing the report generated by `Disk Usage`_.

   :type: `Function Procedure`_
   :parameters:
      * filePath (`String`_) - Path of the report generated by `Disk Usage`_
      * searchString (`String`_) - The data to be extracted from the report
   :returns: Size of the Folder in MegaBytes.
   :return type: `Variant`_

Utilities
---------

genTimeStamp
^^^^^^^^^^^^

.. function:: genTimeStamp(Optional ByVal sd As String = "_", Optional ByVal st As String = "_", Optional ByVal s As String = "-", Optional ByVal prefix As String = "TS-", Optional ByVal postfix as String = "")

   Returns a current timestamp string in the following format:
   **<prefix>DD<sd>MM<sd>YYYY<s>HH<st>MM<st>SS<postfix>**

   :type: `Function Procedure`_
   :parameters:
      * sd (`String`_) - Separator for Date
      * st (`String`_) - Separator for Time
      * s (`String`_) - Separator for Date and Time
      * prefix (`String`_) - String to be prepended to the timestamp
      * postfix (`String`_) - String to be appended to the timestamp
   :returns: Current Timestamp.
   :return type: `String`_

applyLower
^^^^^^^^^^

.. function:: applyLower(rng As Range)

   Converts all the text in the given range to Lower Case.

   :type: `Sub Procedure`_
   :parameters:
      rng (`Range`_) - The cell range to be processed.

trimToChar
^^^^^^^^^^

.. function:: trimToChar(rng As Range, ByVal nChar As Integer)

   Truncates the text longer than nChar from all cells in the specified range.

   :type: `Sub Procedure`_
   :parameters:
      * rng (`Range`_) - The cell range to be processed.
      * nChar (`Integer`_) - Maximum number of characters allowed in each cell.

Advanced Utilities
----------------------

exeCmd
^^^^^^

.. function:: exeCmd(ByVal cmd As String, Optional ByVal visible As Integer = 0, Optional ByVal wait As Boolean = False, Optional ByVal quot As String = """")

   Executes the given cmd in windows command prompt.

   :type: `Sub Procedure`_
   :parameters:
      * cmd (`String`_) - Command to be executed
      * visible (`Integer`_) - Window Style integer as documented in `Shell`_
      * wait (`Boolean`_) - Wait till the command executes
      * quot (`String`_) - Enclose the command with quot string

openFolder
^^^^^^^^^^

.. function:: openFolder(ByVal FolderName As String, Optional ByVal focus = vbNormalFocus)

   Opens the specified folder in windows explorer.

   :type: `Sub Procedure`_
   :parameters:
      * FolderName (`String`_) - Path of the folder to be opened
      * focus (`Variant`_) - Focus mode as documented in `Shell Constants`_

RemoveTree
^^^^^^^^^^

.. warning:: File removal is **PERMANANT** and CANNOT be **reversed** or **recovered** later.

.. function:: RemoveTree(ByVal path As String, Optional ByVal warn As Integer = 2)

   Recursively remove all files and subfolders from the specified folder.

   :type: `Sub Procedure`_
   :parameters:
      * path (`String`_) - Path of the folder to be deleted
      * warn (`Integer`_) - Display a Warning prompt? (defaults to yes)

fetchFile
^^^^^^^^^

.. function:: fetchFile(ByVal src As String, ByVal dest As String)

   Downloads a file from the Internet.

   :type: `Function Procedure`_
   :parameters:
      * src (`String`_) - Download URL of the file
      * dest (`String`_) - Target folder path for the file
   :returns: Status of download.
   :return type: `Boolean`_


.. _Dictionary: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object
.. _Boolean: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/boolean-data-type
.. _String: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/string-data-type
.. _MsgBox Constant: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-constants
.. _Variant: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/variant-data-type
.. _Integer: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/integer-data-type
.. _Range: https://docs.microsoft.com/en-us/office/vba/api/excel.range(object)
.. _ListObject: https://docs.microsoft.com/en-us/office/vba/api/excel.listobject
.. _Sub Procedure: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement
.. _Function Procedure: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement
.. _Universal Naming Convention: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-dfsc/149a3039-98ce-491a-9268-2f5ddef08192
.. _Windows Robust Copier: https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
.. _Disk Usage: https://docs.microsoft.com/en-us/sysinternals/downloads/du
.. _Shell: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shell-function
.. _Shell Constants: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shell-constants
.. _Sort: https://docs.microsoft.com/en-us/office/vba/api/excel.range.sort