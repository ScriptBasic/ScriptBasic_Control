VERSION 5.00
Begin VB.Form frmCallBack 
   Caption         =   "CallBack Test"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form2"
   ScaleHeight     =   1935
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRetNow 
      Caption         =   "Return now"
      Height          =   420
      Left            =   2970
      TabIndex        =   4
      Top             =   1350
      Width           =   1320
   End
   Begin VB.CommandButton cmdOp2 
      Caption         =   "Operation 2"
      Height          =   420
      Left            =   3015
      TabIndex        =   3
      Top             =   720
      Width           =   1275
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   855
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   315
      Width           =   1590
   End
   Begin VB.CommandButton cmdOp1 
      Caption         =   "Operation 1"
      Height          =   420
      Left            =   3060
      TabIndex        =   0
      Top             =   135
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   420
      Left            =   225
      TabIndex        =   5
      Top             =   900
      Width           =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   690
   End
End
Attribute VB_Name = "frmCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ext_SBCallBack Lib "COM.dll" Alias "SBCallBack" (ByVal EntryPoint As Long, ByVal arg As Long) As Long
Private Declare Function eng_SBCallBack Lib "sb_engine.dll" Alias "SBCallBack" (ByVal EntryPoint As Long, ByVal arg As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private m_owner As Sample

Private Function TriggerCallBack(nodeID As Long, argValue As Long)

    'this little trick is so this works for both the standard COM extension,
    'as well as the embedded sb_engine.dll project i am working on..
    If GetModuleHandle("COM.dll") <> 0 Then
        TriggerCallBack = ext_SBCallBack(nodeID, argValue)
    ElseIf GetModuleHandle("sb_engine.dll") <> 0 Then
        TriggerCallBack = eng_SBCallBack(nodeID, argValue)
    Else
        MsgBox "Could not find the extension dll or the sb_engine.dll to get the sbcallback export from?", vbExclamation
    End If
    
End Function

Function ShowCallBackForm(defVal As Long, owner As Sample) As Long
    On Error Resume Next
    Set m_owner = owner
    txtValue = defVal
    Me.Show 1
    Set m_owner = Nothing
    ShowCallBackForm = CLng(txtValue)
End Function

Private Sub cmdOp1_Click()

    Dim nodeID As Long
    Dim arg As Long
    
    On Error Resume Next
    nodeID = m_owner.CallBackHandlers("frmCallBack.cmdOp1_Click")
    
    If nodeID = 0 Then
        MsgBox "Script writer forgot to register a callback handler for me..", vbInformation
        Exit Sub
    End If
    
    arg = CLng(txtValue)
    txtValue = TriggerCallBack(nodeID, arg)
    
End Sub

Private Sub cmdOp2_Click()

    Dim nodeID As Long
    Dim arg As Long
    
    On Error Resume Next
    nodeID = m_owner.CallBackHandlers("frmCallBack.cmdOp2_Click")
    
    If nodeID = 0 Then
        MsgBox "Script writer forgot to register a callback handler for me..", vbInformation
        Exit Sub
    End If
    
    arg = CLng(txtValue)
    txtValue = TriggerCallBack(nodeID, arg)
    
End Sub

Private Sub cmdRetNow_Click()
    Me.Visible = False 'this will break the me.show modal lock
End Sub

Private Sub Form_Load()
    Label2.Caption = "Registered Handlers: " & m_owner.CallBackHandlers.Count
End Sub
