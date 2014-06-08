Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal Length As Long)



Public Sub vb_stdout(ByVal lpMsg As Long, ByVal sz As Long)
    Dim b() As Byte
    Dim msg As String
    ReDim b(sz)
    CopyMemory b(0), ByVal lpMsg, sz
    msg = StrConv(b, vbUnicode)
    With Form1.txtOut
        .Text = .Text & Replace(msg, vbLf, vbCrLf)
    End With
End Sub
