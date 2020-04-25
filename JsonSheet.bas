Attribute VB_Name = "JsonSheet"
''
' JsonExcelSheet v0.0.1
' (c) DKiyanov - https://github.com/DKiyanov/JsonExcelSheet
'
' Outputs and reading json to/from MS Excel sheet
'
' Errors:
' 10002 - JSON output/reading error
'
' @class JsonSheet
' @author DmKiyanov@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'
' =============================================
' Description of format output
' =============================================
' Elementary type values are data that is not a structure or an array.
' For each value of an elementary type in a json string, can be formed a string describing the "path" to it value, recorded by the following rules:
'   Recording the path starts from the beginning of the json string, from left to right.
'   "$" - indicates the root, located at the beginning of each path string.
'   "[" - descent to a new level, into the array.
'       Values of elementary type can be embedded in the json array, in this case after the symbol "[" should follow the symbol "e".
'   "{" - descent to a new level, into the structure
'       After the symbol "{" should follow the field name in quotation marks <">.
'
' The header of each Excel column contains a "path" to the value of the elementary data type in the json string.
' For each array in json string, additional Excel column is allocated. This column will contain the row numbers of the json array.
' Header name ("path") for this column, built as described above but the character "i" is added to the end of the string (as a result "$...[i").
'
' Json data is displayed below the Excel header row.
' Each elementary data type value from a json string is displayed in a column which header and data path match.
' For each row of the json array, a separate row of the Excel sheet is allocated, in the corresponding "$...[i" column, displays the row number of the json array (numbering starts from 1).
'
' Sometimes it’s necessary, to transpose the output of the array, corresponding columns are allocated for this.
' The header of these columns contain the path to the value, with the json array row number added to the end of path (no separators, just a number).
'
' =============================================
' Public Methods
' =============================================
'
' Public Function GetJsonHeader(json As Object) As collection
'   Returns a collection containing fields path/addresses
'   json object can be obtained by: Set json = JsonConverter.ParseJson(jsonString)
'
' Public Sub OutJsonHeader(json As Object, Sheet As Worksheet, Row As Integer, Col As Integer)
'   Forms on a sheet columns headings containing field path/addresses
'
' Public Sub OutJsonBody(json As Object, Sheet As Worksheet, Row As Integer, Col As Integer)
'   Output on a worksheet content of json. looking at the column headings
'   the location of the row of column headers is specified in the call parameters
'   data is output below this heading row
'
' Public Function ReadJsonFromSheet(Sheet As Worksheet, Row As Integer, Col As Integer) As Object
'   Reads data from a worksheet and fills an object representing Json (Dictionary or Collection in root)
'   further, this object can be converted to json string using JsonConverter.ConvertToJson(<object>, Whitespace:=2)
'   call parameters must point to the first cell of the heading row

Option Explicit

Private gColumns As collection

Private gColRowStack As Dictionary
Private gRowIndex As Long
Private gNewRow As Boolean
Private gNewArray As Boolean
Private gLastArrayToRowIndex As Long
Private g_iRowIndex As Long

Private gSheet As Worksheet
Private gSheetRow As Integer
Private gSheetCol As Integer

Private gRootObject As Object

Public Function GetJsonHeader(json As Object) As collection
    Set gColumns = New collection
    
    Call ExpandJsonHead("$", json)
    
    Set GetJsonHeader = gColumns
    Set gColumns = Nothing
End Function

Private Sub ExpandJsonHead(path As String, obj As Variant)
    Dim tName As String
    Dim lPath As String
    Dim subObj As Variant
    Dim tSubName As String
    
    tName = TypeName(obj)
    
    If tName = "Collection" Then
        lPath = path & "["
        AddColumn (lPath & "i")
        
        For Each subObj In obj
            Call ExpandJsonHead(lPath, subObj)
        Next
        
        Exit Sub
    End If
    
    If tName = "Dictionary" Then
        Dim Dict As Dictionary
        Dim i As Integer
        
        Set Dict = obj
        
        For i = 0 To Dict.Count - 1
            lPath = path & "{""" & Dict.Keys()(i) & """"
            
            tSubName = TypeName(Dict.Items(i))
            
            If tSubName = "Collection" Or tSubName = "Dictionary" Then
                Set subObj = Dict.Items(i)
                Call ExpandJsonHead(lPath, subObj)
            Else
                AddColumn (lPath)
            End If
        Next
        
        Exit Sub
    End If
    
    AddColumn (path & "e")
End Sub

Private Sub AddColumn(colName As String)
    On Error GoTo Exit_
        Call gColumns.Add(colName, colName)
Exit_:
End Sub

Public Sub OutJsonHeader(json As Object, Sheet As Worksheet, row As Integer, col As Integer)
    Dim Columns As collection
    Dim colIndex As Integer
    
    Set Columns = GetJsonHeader(json)
    
    For colIndex = 1 To Columns.Count
      Sheet.Cells(row, col + colIndex - 1).Value = Columns(colIndex)
    Next
End Sub

Private Sub ReadJsonHeader(Sheet As Worksheet, row As Integer, col As Integer)
    Set gColumns = New collection
    
    Dim LastColumn As Long
    LastColumn = Sheet.UsedRange.Columns(Sheet.UsedRange.Columns.Count).Column
    
    Dim colIndex As Long
    For colIndex = col To LastColumn
        gColumns.Add (Sheet.Cells(row, colIndex).Text)
    Next
End Sub

Public Sub OutJsonBody(json As Object, Sheet As Worksheet, row As Integer, col As Integer)
    Set gColRowStack = New Dictionary
    gRowIndex = 1
    gNewRow = True
    gNewArray = False
    
    Set gSheet = Sheet
    gSheetRow = row - 1
    gSheetCol = col - 1
    
    Call ReadJsonHeader(Sheet, row, col)
    
    Call ExpandJsonBody("$", json)
End Sub

Private Sub ExpandJsonBody(path As String, obj As Variant)
    Dim tName As String
    Dim lPath As String
    Dim subObj As Variant
    Dim tSubName As String
    Dim iRowIndex As Long
    
    Dim iCol As String
    
    tName = TypeName(obj)
    
    If tName = "Collection" Then
        lPath = path & "["
        iCol = lPath & "i"
        gColRowStack.item(iCol) = 0
        
        gNewArray = True
        For Each subObj In obj
            gNewRow = True
            iRowIndex = iRowIndex + 1
            gColRowStack.item(iCol) = iRowIndex
            g_iRowIndex = iRowIndex
            
            Call ExpandJsonBody(lPath, subObj)
        Next
        
        gColRowStack.Remove (iCol)
        
        Exit Sub
    End If
    
    If tName = "Dictionary" Then
        Dim Dict As Dictionary
        Dim i As Integer
        
        Set Dict = obj
        
        For i = 0 To Dict.Count - 1
            lPath = path & "{""" & Dict.Keys()(i) & """"
            
            tSubName = TypeName(Dict.Items(i))
            
            If tSubName = "Collection" Or tSubName = "Dictionary" Then
                Set subObj = Dict.Items(i)
                Call ExpandJsonBody(lPath, subObj)
            Else
                Call AddBodyValue(lPath, Dict.Items(i))
            End If
        Next
        
        Exit Sub
    End If
    
    Call AddBodyValue(path & "e", obj)
End Sub

Private Sub AddBodyValue(colName As String, val As Variant)
    Dim colIndex As Integer
    Dim colIndexI As Integer
    Dim colRow As Variant
    Dim col As Variant
    Dim cell As Range
       
    Dim colNameI As String
    colNameI = colName & CStr(g_iRowIndex)
    colIndexI = cIndexOff(gColumns, colNameI)
    If colIndexI >= 0 Then
        If gNewArray Then
            gNewArray = False
            
            If gNewRow Then
                gNewRow = False
                gRowIndex = gRowIndex + 1
                
                For Each colRow In gColRowStack.Keys
                    colIndex = cIndexOff(gColumns, colRow)
                    If colIndex >= 0 Then
                        gSheet.Cells(gSheetRow + gRowIndex, gSheetCol + colIndex).Value = gColRowStack.item(colRow)
                    End If
                Next
            End If
        
            gLastArrayToRowIndex = gRowIndex
        End If
        
        Call SetCellValue(gSheetRow + gLastArrayToRowIndex, gSheetCol + colIndexI, val)
        
        Exit Sub
    End If
        
    If gNewRow Then
        gNewRow = False
        gRowIndex = gRowIndex + 1
        
        For Each colRow In gColRowStack.Keys
            colIndex = cIndexOff(gColumns, colRow)
            If colIndex >= 0 Then
                gSheet.Cells(gSheetRow + gRowIndex, gSheetCol + colIndex).Value = gColRowStack.item(colRow)
            End If
        Next
    End If
    
    If gNewArray Then
        gNewArray = False
        gLastArrayToRowIndex = gRowIndex
    End If
    
    colIndex = cIndexOff(gColumns, colName)
    If colIndex >= 0 Then
        Call SetCellValue(gSheetRow + gRowIndex, gSheetCol + colIndex, val)
    End If
End Sub

Private Sub SetCellValue(row As Integer, col As Integer, val As Variant)
    Dim cell As Range
    Set cell = gSheet.Cells(row, col)
    
    If IsNull(val) Then
        cell.Value = "Null"
        Exit Sub
    End If
    
    Dim tName As String
    tName = TypeName(val)
    
    If tName = "String" Then
        cell.NumberFormat = "@"
        cell.Value = CStr(val)
        Exit Sub
    End If
    
    cell.NumberFormat = "General"
    cell.Value = val
End Sub

Private Function cIndexOff(collection As collection, val As Variant) As Integer
    Dim index As Long
    Dim item As Variant
    
    cIndexOff = -1
    
    For Each item In collection
        index = index + 1
        If item = val Then
            cIndexOff = index
            Exit For
        End If
    Next
End Function

Public Function ReadJsonFromSheet(Sheet As Worksheet, row As Integer, col As Integer) As Object
    Call ReadJsonHeader(Sheet, row, col)
    
    Dim LastRow As Long
    LastRow = Sheet.UsedRange.Rows(Sheet.UsedRange.Rows.Count).row
    
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim cell As Range
    Dim cellValue As Variant
    
    Set gRootObject = Nothing
    
    For rowIndex = row + 1 To LastRow
    For colIndex = col To gColumns.Count
        Set cell = Sheet.Cells(rowIndex, colIndex)
        cellValue = cell.Value
        If cell.NumberFormat = "General" And UCase(cellValue) = "NULL" Then
            cellValue = Null
        End If
        If Not IsEmpty(cellValue) Then
            Call SetValue(gColumns(colIndex), rowIndex, cellValue)
        End If
    Next
    Next
    
    Set ReadJsonFromSheet = gRootObject
End Function

Private Sub SetValue(path As String, rowIndex As Long, val As Variant)
    If Right$(path, 1) = "i" Then
        Exit Sub
    End If
    
    Dim pathLen As Long
    pathLen = Len(path)
    
    Dim pos As Long
    Dim ch As String
    Dim str As String
    
    Dim quotePosStart As Long
    
    Dim curObjType As String ' $ [ {
    
    Dim curCltn As collection
    Dim curCltnRow As Long
    
    Dim curDict As Dictionary
    Dim curDictField As String
    Dim colIndex As Integer
    
    quotePosStart = -1
    curObjType = "$"
    
    For pos = 2 To pathLen
        ch = Mid$(path, pos, 1)
        Select Case ch
            Case "["
                Select Case curObjType
                    Case "$"
                        If gRootObject Is Nothing Then
                            Set gRootObject = New collection
                        End If
                        Set curCltn = gRootObject
                    Case "["
                        Set curCltn = GetJsonCollectionItem(curCltn, curCltnRow, ch)
                    Case "{"
                        Set curCltn = GetJsonDictionaryItem(curDict, curDictField, ch)
                End Select
                
                str = Mid$(path, 1, pos) & "i"
                colIndex = cIndexOff(gColumns, str)
                If colIndex < 0 Then
                    Err.Raise 10002, "JsonSheet", "column with name '" & str & "' not found"
                End If
                curCltnRow = Cells(rowIndex, colIndex)
                curObjType = ch
            Case "{"
                Select Case curObjType
                    Case "$"
                        If gRootObject Is Nothing Then
                            Set gRootObject = New Dictionary
                        End If
                        Set curDict = gRootObject
                    Case "["
                        Set curDict = GetJsonCollectionItem(curCltn, curCltnRow, ch)
                    Case "{"
                        Set curDict = GetJsonDictionaryItem(curDict, curDictField, ch)
                End Select
                
                curDictField = "?"
                curObjType = ch
            Case """"
                If curDictField <> "?" Then
                    Err.Raise 10002, "JsonSheet", "not a suitable place for symbol <""> in " & path
                End If
                
                If quotePosStart < 0 Then
                    quotePosStart = pos + 1
                Else
                    curDictField = Mid$(path, quotePosStart, pos - quotePosStart)
                    quotePosStart = -1
                End If
        End Select
    Next
    
    Dim numstr As String
    Dim i As Long
    numstr = "0123456789"
    
    If InStr(numstr, ch) > 0 Then
        For i = 1 To pathLen
            If InStr(numstr, Mid$(path, pathLen - i, 1)) = 0 Then
                curCltnRow = CLng(Right$(path, i))
                Exit For
            End If
        Next
    End If
    
    Select Case curObjType
        Case "$"
            Err.Raise 10002, "JsonSheet", "parse error " & path
        Case "["
            Call GetJsonCollectionItem(curCltn, curCltnRow, "V", val)
        Case "{"
            Call GetJsonDictionaryItem(curDict, curDictField, "V", val)
    End Select
End Sub

Private Function GetJsonCollectionItem(Cltn As collection, row As Long, xObjType As String, Optional val As Variant) As Object
    Dim dCount As Integer
    Dim vEmpty As Variant
    Dim i As Long
    Dim obj As Object
    Dim tName As String
    
    vEmpty = Empty
    
    dCount = row - Cltn.Count
    If dCount > 1 Then
        For i = 1 To dCount
            Cltn.Add (vEmpty)
        Next
        dCount = 1
    End If
    
    If dCount = 1 Then
        Select Case xObjType
            Case "V"
                Call Cltn.Add(val)
            Case "["
                Set GetJsonCollectionItem = New collection
                Call Cltn.Add(GetJsonCollectionItem)
            Case "{"
                Set GetJsonCollectionItem = New Dictionary
                Call Cltn.Add(GetJsonCollectionItem)
        End Select
        
        Exit Function
    End If
       
    If xObjType <> "V" Then
        Set GetJsonCollectionItem = Cltn.item(row)
        
        If Not IsEmpty(GetJsonCollectionItem) Then
            tName = TypeName(GetJsonCollectionItem)
            If tName = "Collection" And xObjType <> "[" Then
                Err.Raise 10002, "JsonSheet", "Error on build json"
            End If
            If tName = "Dictionary" And xObjType <> "{" Then
                Err.Raise 10002, "JsonSheet", "Error on build json"
            End If
            
            Exit Function
        End If
    End If
    
    Cltn.Remove (row)
    
    Select Case xObjType
        Case "V"
            Call Cltn.Add(val, Before:=row)
        Case "["
            Set GetJsonCollectionItem = New collection
            Call Cltn.Add(GetJsonCollectionItem, Before:=row)
        Case "{"
            Set GetJsonCollectionItem = New Dictionary
            Call Cltn.Add(GetJsonCollectionItem, Before:=row)
    End Select
End Function

Private Function GetJsonDictionaryItem(Dict As Dictionary, Key As String, xObjType As String, Optional val As Variant) As Object
    Dim tName As String
    
    If xObjType = "V" Then
        Dict.item(Key) = val
        Exit Function
    End If
    
    If Dict.Exists(Key) Then
        Set GetJsonDictionaryItem = Dict.item(Key)
        
        tName = TypeName(GetJsonDictionaryItem)
        If tName = "Collection" And xObjType <> "[" Then
            Err.Raise 10002, "JsonSheet", "Error on build json"
        End If
        If tName = "Dictionary" And xObjType <> "{" Then
            Err.Raise 10002, "JsonSheet", "Error on build json"
        End If
        
        Exit Function
    End If
    
    Select Case xObjType
        Case "["
            Set GetJsonDictionaryItem = New collection
            Set Dict.item(Key) = GetJsonDictionaryItem
        Case "{"
            Set GetJsonDictionaryItem = New Dictionary
            Set Dict.item(Key) = GetJsonDictionaryItem
    End Select
End Function
