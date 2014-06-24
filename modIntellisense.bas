Attribute VB_Name = "modIntellisense"

Public modules As New Collection

Function InitIntellisense(fpath As String) As Boolean

    On Error Resume Next
    
    If Not FileExists(fpath) Then Exit Function
    
    Dim curList As String
    Dim inBlock As Boolean
    Dim curModule As String
    
    tmp = Split(ReadFile(fpath), vbCrLf)
    For Each x In tmp
        x = Replace(x, vbTab, " ")
        x = Replace(x, "  ", " ")
        x = Trim(x)
        If Len(x) = 0 Then GoTo skipLine
        If Left(x, 1) = "#" Or Left(x, 1) = "'" Then GoTo skipLine 'its a comment ignore this line..
        
        If curModule = "" Then
            curModule = x
            GoTo skipLine
        End If
        
        If inBlock And x = "{" Then
            MsgBox "Error parsing " & fpath & " you can not nest {} blocks", vbInformation
            Exit Function
        End If
        
        If x = "{" Then
            inBlock = True
            GoTo skipLine
        End If
        
        If x = "}" Then
            inBlock = False
            If curModule = Empty Then
                MsgBox "CurModule has not been named? error parsing " & fpath, vbInformation
                Exit Function
            End If
            If curList <> Empty Then
                modules.Add Trim(curList), curModule
            End If
            curList = Empty
            curModule = Empty
            GoTo skipLine
        End If
        
        AddLine x, curList
        
        
skipLine:
    Next
    
    InitIntellisense = (Err.Number = 0)
    
End Function

Private Function AddLine(x, ByRef curList As String) As Boolean

    On Error Resume Next

    tmp = Split(x, " ")
    For Each Y In tmp
        If Len(Y) > 0 Then
            curList = curList & " :" & Y
        End If
    Next
    
    AddLine = (Err.Number = 0)
    
End Function




