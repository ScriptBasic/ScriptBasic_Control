VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public isGlobal As Boolean
Private m_name As String
Private m_value As Variant
Public varType As String
Public index As Long 'for arrays
Public pAryElement As Long

Property Let name(v As String)
    
    m_name = Replace(v, "main::", Empty)
    
'    If InStr(v, "::") > 0 Then
'        tmp = Split(v, "::")
'        m_name = tmp(1)
'        namespace = tmp(0)
'    Else
'        m_name = v
'    End If
End Property

Property Get name() As String
    name = m_name
End Property

Property Let value(v As String)
    If Left(v, 5) = "ARRAY" And InStr(v, "@ 0x") > 0 Then
        tmp = Split(v, "@ 0x")
        m_value = tmp(0)
        pAryElement = CLng("&h" & tmp(1))
    Else
        m_value = v
    End If
End Property

Property Get value() As String
    value = m_value
End Property
