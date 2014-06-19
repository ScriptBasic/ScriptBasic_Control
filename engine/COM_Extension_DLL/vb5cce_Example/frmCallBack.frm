VERSION 5.00
Begin VB.Form frmCallBack 
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return Now"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOp1 
      Caption         =   "Operation 1"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Value 
      Caption         =   "Value"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is the simple callback it takes one long arg and returns a long
Private Declare Function ext_SBCallBack Lib "COM.dll" Alias "SBCallBack" (ByVal EntryPoint As Long, ByVal arg As Long) As Long
Private m_owner As CCESample

Function ShowCallBackForm(defVal As Long, owner As CCESample) As Long
    On Error Resume Next
    Set m_owner = owner
    txtValue = defVal
    Me.Show 1
    Set m_owner = Nothing
    ShowCallBackForm = CLng(txtValue)
    Unload Me
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
    txtValue = ext_SBCallBack(nodeID, arg)
    
End Sub

Private Sub cmdReturn_Click()
     Me.Visible = False 'this will break the me.show modal lock
End Sub
