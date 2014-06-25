Attribute VB_Name = "modHostObjs"
'this module is used for script to host app object integration.. (for embedded clients ie ActiveX control builds..)

Private objs As New Collection 'of CVariable

Sub AddObject(name As String, Value As Object)
    Dim v As New CVariable
    v.name = name
    v.Value = ObjPtr(Value)
    v.varType = "long"
    On Error Resume Next
    objs.Add v, name
End Sub

Sub AddLong(name As String, Value As Long)
    Dim v As New CVariable
    v.name = name
    v.Value = Value
    v.varType = "long"
    On Error Resume Next
    objs.Add v, name
End Sub

Sub AddString(name As String, Value As String)
    Dim v As New CVariable
    If Len(Value) > 1024 Then Exit Sub
    v.name = name
    v.Value = Value
    v.varType = "string"
    On Error Resume Next
    objs.Add v, name
End Sub

'this is used for script to host app object integration.. (for embedded clients)
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
        HostResolver = v.Value
    Else
        If Len(v.Value) + 1 < totalBufSz Then
            b() = StrConv(v.Value & Chr(0), vbFromUnicode)
            CopyMemory ByVal buf, b(0), UBound(b) + 1
        End If
    End If
    
End Function

