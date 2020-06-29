Option Explicit
'character set for non ascii non printable characters
Private Const REGEX_ASCII_NON_PRINTABLE_PATTERN = "[\u0007-\u001F]"
 
'character set for non-ascii characters
Private Const REGEX_UNICODE_PATTERN = "[^\u0000-\u007F]"
 
Sub copyVisibleCells(rng As Range, destWorksheet As Worksheet)
    'Select visible cells in a range and paste only the visible cells to another worksheet
    rng.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    'Copy Visible cells only in the range and paste in target sheet
    Selection.Copy
    destWorksheet.Select
    destWorksheet.Paste
End Sub
Sub copyVisibleCellsEnd(rng As Range, destWorksheet As Worksheet)
    'Select visible cells in a range and paste only the visible cells to last row of worksheet
    Dim rowIndex As Long
   
    If getVisibleRowCount(rng) = 1 Then
        'exit sub if there is only a header and then select the destination worksheet
        destWorksheet.Select
        Exit Sub
    End If
  
    Set rng = rng.Offset(1).Resize(rng.Rows.count - 1)
 
    rng.SpecialCells(xlCellTypeVisible).Select
    'Copy Visible cells only in the range and paste in target sheet
    Selection.Copy
    rowIndex = destWorksheet.Range("A1").CurrentRegion.Rows.count + 1
    destWorksheet.Select
    destWorksheet.Range("A" & rowIndex).Select
    destWorksheet.Paste
End Sub

Function getColumnCount(rng As Range) As Long
'return the number of columns from a range
    getColumnCount = rng.Columns.count
End Function
Function getRowCount(rng As Range) As Long
'returns the number of rows from a range
    getRowCount = rng.Rows.count
End Function

Function getVisibleColumnCount(rng As Range) As Long
'returns the number of visible columns from a range
    Dim cellItem As Range
    Dim count As Long
    count = 0
    For Each cellItem In rng.SpecialCells(xlCellTypeVisible).Columns
        count = count + 1
    Next cellItem
    getVisibleColumnCount = count
End Function

Function getVisibleRowCount(rng As Range) As Long
'return the number of visible rows from a range
    Dim cellItem As Range
    Dim count As Long
    count = 0
    For Each cellItem In rng.SpecialCells(xlCellTypeVisible).Rows
        count = count + 1
    Next cellItem
    getVisibleRowCount = count
End Function

Function isVisibleRowGreaterThan(rng As Range, rowCount) As Boolean
'return the number of visible rows from a range
    Dim cellItem As Range
    Dim count As Long
    Dim isGreater As Boolean
    count = 0
    isGreater = False
    For Each cellItem In rng.SpecialCells(xlCellTypeVisible).Rows
        count = count + 1
        If count > rowCount Then
            isGreater = True
            Exit For
        End If
    Next cellItem
    isVisibleRowGreaterThan = isGreater
End Function

Function fileExists(file As String) As Boolean
'check if a file exists returns true if yes and false if not
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = fso.fileExists(file)
    Set fso = Nothing
End Function

Function folderExists(Path As String) As Boolean
'check if a folder exists or not returns true if exisit and false if not
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderExists = fso.folderExists(Path)
    Set fso = Nothing
End Function

 Function getFileCount(psPath As String) As Long
'strive4peace
'uses Late Binding. Reference for Early Binding:
'  Microsoft Scripting Runtime
   'PARAMETER
   '  psPath is folder to get the number of files for
   '     for example, c:\myPath
   ' Return: Long
   '    -1 = path not valid
   '     0 = no files found, but path is valid
   '    99 = number of files where 99 is some number
  
   'inialize return value
   getFileCount = -1
   'skip errors
   On Error Resume Next
   'count files in folder of FileSystemObject for path
   With CreateObject("Scripting.FileSystemObject")
      getFileCount = .GetFolder(psPath).Files.count
   End With
End Function
 
Function getFileNamesFromPath(Path As String, Optional ext As String = "", Optional excludePrefix As String = "") As Collection
'returns filenames from a folder path
'if ext is not empty then filter file names by file extension. Example of ext parameter file extension strings docx, exe
'if excludePrefix is not empty exclude all files from folder that begins with the prefix string
    Dim col As New Collection
    Dim filename As String
       
    If ext <> "" Then
        filename = Dir(ThisWorkbook.Path & "\*." & ext, vbNormal & vbHidden)
    Else
        filename = Dir(ThisWorkbook.Path & "\", vbNormal & vbHidden)
    End If
   
    Do While filename <> ""
        If excludePrefix <> "" Then
            If InStr(1, filename, excludePrefix) = 0 Then
                col.Add filename
            End If
        Else
            col.Add filename
        End If
        filename = Dir
    Loop
   
    Set getFileNamesFromPath = col
End Function
 
Function deleteFolder(folderPath As String) As Boolean
'delete a folder from folder path
'this function deletes empty or non empty folder
'the function will failed if there is a permission access issue else returns true
'if the folder does not exists true is returned
    Dim fso As Object
    Dim tempPath As String
    tempPath = Trim(folderPath)
    If tempPath <> "" Then
        If Right(tempPath, 1) = "\" Then
            tempPath = Left(tempPath, Len(tempPath) - 1)
        End If
    End If
 
    On Error GoTo errHandler:
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.folderExists(tempPath) Then
        Call fso.deleteFolder(tempPath)
    End If
   
    deleteFolder = True
exitSuccess:
    Exit Function
errHandler:
    Debug.Print Err.Number, Err.Description
    GoTo exitSuccess
End Function
 
Function getFolderCount(psPath As String) As Long
'strive4peace
'uses Late Binding. Reference for Early Binding:
'  Microsoft Scripting Runtime
   'PARAMETER
   '  psPath is path to get the number of folders for
   '     for example, c:\myPath
   ' Return: Long
   '  -1 = path not valid
   '   0 = no folders found, but path is valid
   '  99 = number of folders where 99 is some number
  
   'inialize return value
   getFolderCount = -1
   'skip errors
   On Error Resume Next
   'count SubFolders in FileSystemObject for psPath
   With CreateObject("Scripting.FileSystemObject")
      getFolderCount = .GetFolder(psPath).SubFolders.count
   End With
End Function
 
Function columnNumToColumnLetter(colNum As Long) As String
'returns an excel column letter from the number number
'if the column letter cannot be determine returns vbnullstring
    Dim regex As Object
    Dim Matches As Object
    Dim addr As String
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "[A-Z]+"
    addr = Cells(1, colNum).Address(False, False)
    If regex.test(addr) Then
        Set Matches = regex.Execute(addr)
        columnNumToColumnLetter = Matches(0)
    Else
        columnNumToColumnLetter = ""
    End If
    Set regex = Nothing
    Set Matches = Nothing
End Function
Sub deleteRowIfCellBlank(rng As Range)
'delete the entire row if any cells are blank
    On Error Resume Next
    rng.Cells.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub
 
Function getColumnIndex(rng As Range, heading As String, Optional ColumnLetter As Boolean = False) As Variant
'returns heading column letter or number if the header is found else returns 0
'if ColumnLetter is true a letter is return if the column is found else 0
    Dim title As Range
    Dim header As Range
    Set title = rng.Rows(1)
    For Each header In title.Cells
        If StrComp(header.Value, heading, vbTextCompare) = 0 Then
            If ColumnLetter = False Then
                getColumnIndex = header.Column
            Else
                getColumnIndex = columnNumToColumnLetter(header.Column)
            End If
            Exit Function
        End If
  
    Next header
    getColumnIndex = 0
    Set title = Nothing
End Function
Function rangeToArray(rng As Range) As Variant
'returns a range of values as an array
 
    ' Declare dynamic array
    Dim tempArray As Variant
 
    ' tempArray values into array from first row
    rangeToArray = rng.Value
End Function
 
Sub arrayToRange(arr As Variant, rng As Range)
'copies array values to a range
'example Range("A1:C1] = Array[1,2,3]
    rng.Value = arr
End Sub
 
Function worksheetExists(sheetName As String) As Boolean
'checks active workbook if a worksheet exists
    Dim ws As Worksheet
      For Each ws In Application.ActiveWorkbook.Worksheets
        If sheetName = ws.Name Then
          worksheetExists = True
          Exit For
        End If
      Next ws
End Function
Function worksheetDelete(sheetName As String) As Boolean
'delete worksheet if the workseet exists in the active workbook by worksheet name
    If worksheetExists(sheetName) Then
        ActiveWorkbook.Worksheets(sheetName).Delete
    End If
    worksheetDelete = True
End Function
Function worksheetCreate(sheetName As String, Optional sheetIndex As Integer = 0) As Worksheet
' create a worksheet with provided sheetname in active workbook
    Dim objSheet As Object
    On Error GoTo errHandler
    If sheetIndex = 0 Then
        sheetIndex = Sheets.count
    End If
    Set objSheet = Sheets.Add(After:=Sheets(sheetIndex))
    objSheet.Name = sheetName
    Exit Function
errHandler:
    MsgBox "Error Number: " & Err.Number & vbCr, vbCr & " Description: & err.Description"
End Function
Function worksheetCopy(wsName As String, Optional wbPath = "", Optional newWsName = "") As Boolean
'copies a worksheet from within the same workbook or from an external workbook
'if newWsName is not an empty string the copied worksheet is renamed to the newWsName
    Dim tempActiveWorkbook As Workbook, wbExternal As Workbook
    On Error GoTo errHandler
    Set tempActiveWorkbook = ActiveWorkbook
    'delete sales force worksheet if it already exists
    If wbPath <> "" Then
        Set wbExternal = Workbooks.Open(filename:=wbPath)
        wbExternal.Sheets(wsName).Copy After:=Workbooks(tempActiveWorkbook.Name).Sheets(tempActiveWorkbook.Sheets.count)
        wbExternal.Close SaveChanges:=False
    Else
        tempActiveWorkbook.Sheets(wsName).Copy After:=Workbooks(tempActiveWorkbook.Name).Sheets(tempActiveWorkbook.Sheets.count)
    End If
  
    If newWsName <> "" Then
        tempActiveWorkbook.ActiveSheet.Name = newWsName
    End If
  
    worksheetCopy = True
exitSuccess:
    Set tempActiveWorkbook = Nothing
    Set wbExternal = Nothing
    Exit Function
errHandler:
    MsgBox Err.Description
    Resume exitSuccess
End Function
 
Sub worksheetUnhideAllRows(Optional ws As Worksheet)
'unhide all rows in a worksheet
'if no worksheet is provided then the active worksheet is used
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    ws.Rows.EntireRow.Hidden = False
End Sub
 
Sub worksheetUnhideAllColumns(Optional ws As Worksheet)
'unhide all columns in a worksheet
'if no worksheet is provided then the active worksheet is used
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    ws.Rows.EntireColumn.Hidden = False
End Sub
 
Sub worksheetUnhideAllRowsAndColumns(Optional ws As Worksheet)
'unhide all rows and columns in a worksheet
'if no worksheet is provided then the active worksheet is used
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    Call worksheetUnhideAllRows(ws)
    Call worksheetUnhideAllColumns(ws)
End Sub
 
Function worksheetIsFilterMode(Optional ws As Worksheet) As Boolean
'returns true if a worksheet has a filter applied else false
'if no worksheet is provided teh active worksheet is used
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    worksheetIsFilterMode = ws.FilterMode
End Function
 
Sub worksheetClearFilter(Optional ws As Worksheet)
'unfilter a worksheet if it worksheet is filtered
'if no worksheet is provided teh active worksheet is used
If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    If worksheetIsFilterMode(ws) Then
        ws.ShowAllData
    End If
End Sub
 
Sub worksheetShowAllData(Optional ws As Worksheet)
'unhides all rows, columns and remove filters from a worksheet
'if no worksheet is provided the active worksheet is used
    If worksheetIsFilterMode(ws) Then
        ws.ShowAllData
    End If
    Call worksheetUnhideAllRowsAndColumns
End Sub
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'             String Functions Section              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'ASCII char URL https://www.ibm.com/support/knowledgecenter/en/ssw_aix_72/com.ibm.aix.networkcomm/conversion_table.htm
 
 
Public Function regexTest(strData As String, pattern As String, Optional isGlobal As Boolean = True, Optional isIgnoreCase As Boolean = True, Optional isMultiLine As Boolean = True) As Boolean
'returns true if a pattern match else false
 
Dim objRegex As Object
 
On Error GoTo errHandler
 
Set objRegex = CreateObject("vbScript.regExp")
With objRegex
    .Global = isGlobal
    .ignoreCase = isIgnoreCase
    .MultiLine = isMultiLine
    .pattern = pattern
    'if the pattern is a match then replace the text else return the orginal string
    If .test(strData) Then
        regexTest = True
    Else
        regexTest = False
    End If
End With
exitSuccess:
    Set objRegex = Nothing
    Exit Function
errHandler:
    regexTest = False
    Debug.Print Err.Description
    Resume exitSuccess
End Function
 
Function regexMatches(data As String, pattern As String, Optional ignoreCase As Boolean = True, Optional globalMatches As Boolean = True) As Collection
'return a collection found from a pattern using regular expressions
 
    Dim regex As Object, theMatches As Object, match As Object
    Dim col As New Collection
    Set regex = CreateObject("vbScript.regExp")
    
    regex.pattern = pattern
    regex.Global = globalMatches
    regex.ignoreCase = ignoreCase
    
    Set theMatches = regex.Execute(data)
    
    For Each match In theMatches
      col.Add match.Value
    Next
   
    Set regexMatches = col
End Function
 
Function regexReplace(strData As String, pattern As String, Optional replace_with_str = vbNullString, Optional isGlobal As Boolean = True, Optional isIgnoreCase As Boolean = True, Optional isMultiLine As Boolean = True) As String
'returns string replacing data using a regex pattern
    Dim objRegex As Object
On Error GoTo errHandler
    Set objRegex = CreateObject("vbScript.regExp")
    With objRegex
        .Global = isGlobal
        .ignoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
        'if the pattern is a match then replace the text else return the orginal string
        If .test(strData) Then
            regexReplace = .Replace(strData, replace_with_str)
        Else
            regexReplace = strData
        End If
    End With
exitSuccess:
    Set objRegex = Nothing
    Exit Function
errHandler:
    regexReplace = strData
    Debug.Print Err.Description
    Resume exitSuccess
End Function
Function regexPatternCount(strData As String, pattern As String, Optional isGlobal As Boolean = True, Optional isIgnoreCase As Boolean = True, Optional isMultiLine As Boolean = True) As Long
'returns the number of matters matches in a string using regex
'-1 will return if there was an error
    Dim objRegex As Object
    Dim Matches As Object
 
On Error GoTo errHandler
    Set objRegex = CreateObject("vbScript.regExp")
    objRegex.pattern = pattern
    objRegex.Global = isGlobal
    objRegex.ignoreCase = isIgnoreCase
    objRegex.MultiLine = isMultiLine
    'Retrieve all matches
    Set Matches = objRegex.Execute(strData)
    'Return the pattern matches count
    regexPatternCount = Matches.count
exitSuccess:
    Set Matches = Nothing
    Set objRegex = Nothing
    Exit Function
errHandler:
    regexPatternCount = -1
    Resume exitSuccess
End Function
 
Function regexRemoveConcatDupChars(data As String) As String
'remove duplicates characters when concatenated together
    regexRemoveConcatDupChars = regexReplace(data, "(.)\1+", "$1")
End Function
Function regexContainsConcatDupChars(data As String) As Boolean
'returns true if there are concatenated characters of the same type in as string provided
    regexContainsConcatDupChars = regexPatternCount(data, "(.)\1+")
End Function
 
Function regexContainsNonAscii(data As String) As Boolean
'returns true if a string contains unicode characters else false
    If regexPatternCount(data, REGEX_UNICODE_PATTERN) > 0 Then
        regexContainsNonAscii = True
    Else
        regexContainsNonAscii = False
    End If
End Function
Function regexLeftTrim(data As String) As String
'returns a string removing spaces and tab characters from the beginning of a string only
    regexLeftTrim = regexReplace(data, "^[\s\t]+")
End Function
Function regexRightTrim(data As String) As String
'returns a string removing spaces and tab characters from the beginning and end of a string
    regexRightTrim = regexReplace(data, "[\s\t]+$")
End Function
Function regexTrim(data As String) As String
'returns a string removing spaces and tab characters from the beginning and end of a string
    data = regexLeftTrim(data)
    data = regexRightTrim(data)
    regexTrim = data
End Function
 
Function setFirstLetterCapitalized(data As String) As String
'returns a string with first letter capitialize
    If Len(data) = 0 Then
        setFirstLetterCapitalized = ""
    Else
        setFirstLetterCapitalized = UCase(Mid(data, 1, 1)) & Mid(data, 2, Len(data))
    End If
End Function
Function setProperCase(data As String) As String
'returns a string with all words starting with a capital letter and the rest lowercase
    setProperCase = StrConv(data, vbProperCase)
End Function
Function sqlStrFormat(data As String) As String
'returns a string replacing single quotes with double single quotes
    Const SINGLE_QUOTE_CHAR = "'"
    sqlStrFormat = Replace(data, SINGLE_QUOTE_CHAR, SINGLE_QUOTE_CHAR & SINGLE_QUOTE_CHAR)
End Function
 
Private Sub displayError(Optional toImmediateWindow As Boolean = True)
'display error code number and description in the immediate window by default
'if toImmediateWindow is false then the error is displayed in a messagebox
'this subroutine is used for ON ERROR GoTo statements error handler section
    If toImmediateWindow Then
        Debug.Print Err.Number, Err.Description
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical
    End If
End Sub
 
Function createDictionary(Optional ignoreCase As Boolean = False) As Object
'returns a dictionary object
'if ignore case is true the dictionary keys will not be case sensitive. The default is case sensitive
    Dim dict As Object
   
    Set dict = CreateObject("Scripting.Dictionary")
   
    If ignoreCase Then
        dict.comparemode = vbTextCompare
    End If
   
    Set createDictionary = dict
End Function
 
Function isValueInRange(rng As Range, search As String, Optional lookIn As XlFindLookIn = XlFindLookIn.xlFormulas, Optional lookAt As XlLookAt = XlLookAt.xlWhole, Optional matchCase As Boolean = False) As String
'returns string address of the cell where the value if found in the range
'if the value is not found than an empty string is returned
 
    Dim cell As Range
   
    Set cell = rng.Find(What:=search, lookIn:=lookIn, lookAt:=lookAt, matchCase:=matchCase)
   
    If cell Is Nothing Then
        isValueInRange = ""
    Else
        isValueInRange = cell.Address
    End If
 
End Function
 
Function countNumberOfNonBlankCells(rng As Range) As Long
'returns the count of cells that are not empty
    countNumberOfNonBlankCells = Application.WorksheetFunction.CountA(rng)
End Function
 
Sub quickSort(vArray As Variant, inLow As Long, inHi As Long)
'sort an array in ascending order
'example quickSort(arr, LBound(arr), UBound(arr))
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
 
  tmpLow = inLow
  tmpHi = inHi
 
  pivot = vArray((inLow + inHi) \ 2)
 
  While (tmpLow <= tmpHi)
 
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend
 
     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend
 
     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
 
  Wend
 
  If (inLow < tmpHi) Then quickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then quickSort vArray, tmpLow, inHi
 
End Sub
 
Function binarySearch(lookupArray As Variant, lookupValue As Variant) As Long
'binary search lookup for arrays
'the array must be sorted when using this function
'-1 is return if not found else the index of where the item is found
 
    Dim lngLower As Long
    Dim lngMiddle As Long
    Dim lngUpper As Long
 
    lngLower = LBound(lookupArray)
    lngUpper = UBound(lookupArray)
 
    Do While lngLower < lngUpper
       
        lngMiddle = (lngLower + lngUpper) \ 2
 
        If lookupValue > lookupArray(lngMiddle) Then
            lngLower = lngMiddle + 1
        Else
            lngUpper = lngMiddle
        End If
       
    Loop
   
    If lookupArray(lngLower) = lookupValue Then
        binarySearch = lngLower
    Else
        binarySearch = -1    'search does not find a match
    End If
End Function
 
Sub sort_search_example()
'example of sorting items in range, look for item in array with binary search, copy sorted results back to range as 2d array
    Dim arr  As Variant, arr2d As Variant
    Dim rng As Range
    Dim x As Long
    Set rng = Range("A2:A6")
   
    'excel ranges start from an index of 1
    ReDim arr(1 To rng.count)
    '2d to copy 1d array back into excel
    ReDim arr2d(1 To rng.count, 0)
   
    For x = LBound(arr) To UBound(arr)
        arr(x) = rng.Cells(x, 1).Value
    Next x
   
    'sort the array
    Call quickSort(arr, LBound(arr), UBound(arr))
   
    Debug.Print "Index of sorted array item -1 means not found: " & CStr(binarySearch(arr, "Jerome"))
   
    'copy the sorted results into a 2d array
    For x = LBound(arr) To UBound(arr)
        arr2d(x, 0) = arr(x)
    Next x
   
    'remove items from 1d array
    Erase arr
   
    'copy values of 2d array back into range without the needs of cell by cell update
    rng.Value = arr2d
End Sub