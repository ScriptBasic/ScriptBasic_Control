Attribute VB_Name = "modIntellisense"

Public modules As New Collection
Public IncludeFiles() As String

Function isIncludeFile(ByVal nameSpace As String) As Boolean
    On Error Resume Next
    nameSpace = Trim(nameSpace)
    If Len(nameSpace) = 0 Then Exit Function
    For i = 0 To UBound(IncludeFiles)
        If LCase(nameSpace) = LCase(GetBaseName(IncludeFiles(i))) Then
            isIncludeFile = True
            Exit Function
        End If
    Next
End Function

'this is overly simplistic..it doesnt account for multiple spaces or tabs..
'should use regular expression really...
Function isFileIncluded(ByVal fName As String, ByVal script As String) As Boolean
    On Error Resume Next
    fName = Trim(fName)
    If Len(fName) = 0 Then Exit Function
    
    If InStr(1, script, "import " & fName, vbTextCompare) > 0 Then
        isFileIncluded = True
        Exit Function
    End If
    
    If InStr(1, script, "include " & fName, vbTextCompare) > 0 Then
        isFileIncluded = True
        Exit Function
    End If

    
End Function

Function InitIntellisense(includeDir As String) As Boolean

    On Error Resume Next
    
    Set modules = New Collection
    If Not FolderExists(includeDir) Then Exit Function
    
    Dim curList As String
    Dim inBlock As Boolean
    Dim curModule As String
    Dim f
    
    IncludeFiles() = GetFolderFiles(includeDir, ".bas", False)
    For Each fPath In IncludeFiles
        tmp = ReadFile(includeDir & "\" & fPath)
        tmp = Replace(tmp, Chr(&HD), Empty)
        tmp = Split(tmp, vbLf)
        For Each X In tmp
            X = Replace(X, vbTab, " ")
            While InStr(X, "  ") > 0
                X = Replace(X, "  ", " ")
            Wend
            X = Trim(X)
            If Len(X) = 0 Then GoTo skipLine
            If Left(X, 1) = "#" Or Left(X, 1) = "'" Then GoTo skipLine 'its a comment ignore this line..
            
            words = Split(X, " ")
            If LCase(words(0)) = "module" Then curModule = LCase(words(1))
            If Len(curModule) = 0 Then GoTo skipLine 'dont start recording till were in the module declare..
            
            If LCase(words(0)) = "const" Then curList = curList & " :" & words(1)
            
            If LCase(words(0)) = "declare" Then
                words(2) = Replace(words(2), "::", Empty)
                curList = curList & " :" & words(2)  'declare sub xxxxx
            End If
            
            If LCase(words(0)) = "end" And LCase(words(1)) = "module" Then Exit For
skipLine:
        Next
        If Len(curModule) > 0 And Len(curList) > 0 Then
            'Debug.Print modules.Count & ") " & curModule & ":" & curList
            modules.Add Trim(curList), Trim(curModule)
        End If
        curList = Empty
        curModule = Empty
    Next

    InitIntellisense = (Err.Number = 0)
    
End Function


'old Format
'
'#comment
'module
'{
'    name1 name2
'}
'
'Function InitIntellisense(fpath As String) As Boolean
'
'    On Error Resume Next
'
'    If Not FileExists(fpath) Then Exit Function
'
'    Dim curList As String
'    Dim inBlock As Boolean
'    Dim curModule As String
'
'    tmp = Split(ReadFile(fpath), vbCrLf)
'    For Each x In tmp
'        x = Replace(x, vbTab, " ")
'        x = Replace(x, "  ", " ")
'        x = Trim(x)
'        If Len(x) = 0 Then GoTo skipLine
'        If Left(x, 1) = "#" Or Left(x, 1) = "'" Then GoTo skipLine 'its a comment ignore this line..
'
'        If curModule = "" Then
'            curModule = x
'            GoTo skipLine
'        End If
'
'        If inBlock And x = "{" Then
'            MsgBox "Error parsing " & fpath & " you can not nest {} blocks", vbInformation
'            Exit Function
'        End If
'
'        If x = "{" Then
'            inBlock = True
'            GoTo skipLine
'        End If
'
'        If x = "}" Then
'            inBlock = False
'            If curModule = Empty Then
'                MsgBox "CurModule has not been named? error parsing " & fpath, vbInformation
'                Exit Function
'            End If
'            If curList <> Empty Then
'                modules.Add Trim(curList), curModule
'            End If
'            curList = Empty
'            curModule = Empty
'            GoTo skipLine
'        End If
'
'        AddLine x, curList
'
'
'skipLine:
'    Next
'
'    InitIntellisense = (Err.Number = 0)
'
'End Function
'
'Private Function AddLine(x, ByRef curList As String) As Boolean
'
'    On Error Resume Next
'
'    tmp = Split(x, " ")
'    For Each Y In tmp
'        If Len(Y) > 0 Then
'            curList = curList & " :" & Y
'        End If
'    Next
'
'    AddLine = (Err.Number = 0)
'
'End Function
'
'
'
'
'
'
'
