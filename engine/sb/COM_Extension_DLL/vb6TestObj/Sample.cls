VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Public Function longTest(v As Long)
    MsgBox "4 byte Long value: " & v & " received"
End Function

Public Function intTest(v As Integer)
    MsgBox "2 Byte Integer value: " & v & " received"
End Function

Public Function ByteTest(v As Byte)
    MsgBox "1 Byte value: " & v & " received"
End Function

Public Function GetString(prompt As String, title, def) As String
    GetString = InputBox(prompt, title, def)
End Function

Public Function LaunchUI() As Long
    Set f = New Form1
    ShowWindow f.hwnd, 1 'this gets around can not display non modal form from dll thing..
    LaunchUI = ObjPtr(f)
End Function

Public Sub BlockUntilFormCloses()
    While Not f Is Nothing
        DoEvents
        Sleep 20
    Wend
End Sub
