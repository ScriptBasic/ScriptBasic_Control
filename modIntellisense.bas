Attribute VB_Name = "modIntellisense"

Public modules As New Collection
Public IncludeFiles() As String
Public FunctionPrototypes() As String

Function GetAutoCompleteStringForIncludes(partial As String) As String()
    Dim x, matches() As String, i As Long
    
    If AryIsEmpty(IncludeFiles) Then Exit Function
    If Len(partial) = 0 Then Exit Function
    
    For i = 0 To UBound(IncludeFiles)
        x = LCase(IncludeFiles(i))
        If Left(x, Len(partial)) = LCase(partial) Then
            push matches, IncludeFiles(i)
        End If
    Next
    
    GetAutoCompleteStringForIncludes = matches()
    
End Function


Function GetAutoCompleteString(partial As String) As String()
    Dim x, matches() As String, i As Long
    
    If AryIsEmpty(FunctionPrototypes) Then Exit Function
    If Len(partial) = 0 Then Exit Function
    
    For i = 0 To UBound(FunctionPrototypes)
        x = LCase(FunctionPrototypes(i))
        If Left(x, Len(partial)) = LCase(partial) Then
            push matches, FunctionPrototypes(i)
        End If
    Next
    
    GetAutoCompleteString = matches()
    
End Function

Function GetAutoCompleteStringForModule(methods As String, partial As String) As String()
    Dim x, matches() As String, i As Long
    Dim m() As String
    
    If Len(methods) = 0 Then Exit Function
    If Len(partial) = 0 Then Exit Function
    
    m() = Split(methods, ":")
    If AryIsEmpty(m) Then Exit Function

    For i = 0 To UBound(m)
        If Len(m(i)) > 0 Then
            x = LCase(Trim(m(i)))
            If Left(x, Len(partial)) = LCase(partial) Then
                push matches, m(i)
            End If
        End If
    Next
    
    GetAutoCompleteStringForModule = matches()
    
End Function

Function LoadFunctionPrototypes(fpath As String)

  Erase FunctionPrototypes
  If Not FileExists(fpath) Then Exit Function
  
  Dim tmp() As String, x
  Const endMarker = "#modules"
  
  tmp = Split(ReadFile(fpath), vbCrLf)
  For Each x In tmp
        x = Trim(x)
        If Left(x, Len(endMarker)) = endMarker Then Exit For
        If Len(x) > 0 And Left(x, 1) <> "#" And Left(x, 1) <> "'" Then
            a = InStr(x, "(")
            If a > 2 Then x = Mid(x, 1, a - 1)
            push FunctionPrototypes, x
        End If
        
  Next
  
  LoadFunctionPrototypes = UBound(FunctionPrototypes)
  
End Function


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
    For Each fpath In IncludeFiles
        tmp = ReadFile(includeDir & "\" & fpath)
        tmp = Replace(tmp, Chr(&HD), Empty)
        tmp = Split(tmp, vbLf)
        For Each x In tmp
            x = Replace(x, vbTab, " ")
            While InStr(x, "  ") > 0
                x = Replace(x, "  ", " ")
            Wend
            x = Trim(x)
            If Len(x) = 0 Then GoTo skipLine
            If Left(x, 1) = "#" Or Left(x, 1) = "'" Then GoTo skipLine 'its a comment ignore this line..
            
            words = Split(x, " ")
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
