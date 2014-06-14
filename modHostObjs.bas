Attribute VB_Name = "modHostObjs"

Private objs As New Collection 'of CVariable

Sub AddObject(name As String, value As Object)
    Dim v As New CVariable
    v.name = name
    v.value = ObjPtr(value)
    v.varType = "long"
    On Error Resume Next
    objs.Add v, name
End Sub

Sub AddLong(name As String, value As Long)
    Dim v As New CVariable
    v.name = name
    v.value = value
    v.varType = "long"
    On Error Resume Next
    objs.Add v, name
End Sub

Sub AddString(name As String, value As String)
    Dim v As New CVariable
    If Len(value) > 1024 Then Exit Sub
    v.name = name
    v.value = value
    v.varType = "string"
    On Error Resume Next
    objs.Add v, name
End Sub


Public Function HostResolver(ByVal buf As Long, ByVal strlen As Long, ByVal totalBufSz As Long) As Long

    Dim b() As Byte
    Dim name As String
    Dim v As CVariable
    
    ReDim b(strlen)
    CopyMemory b(0), ByVal buf, strlen
    name = StrConv(b, vbUnicode)
    If Right(name, 1) = Chr(0) Then name = Left(name, Len(name) - 1)
    
    On Error Resume Next
    Set v = objs(name)
    
    If v Is Nothing Then Exit Function 'returns 0
    
    If v.varType = "long" Then
        HostResolver = v.value
    Else
        If Len(v.value) + 1 < totalBufSz Then
            b() = StrConv(v.value & Chr(0), vbFromUnicode)
            CopyMemory ByVal buf, b(0), UBound(b) + 1
        End If
    End If
    
End Function

