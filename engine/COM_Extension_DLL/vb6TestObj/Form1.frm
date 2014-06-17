VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Try me"
      Height          =   375
      Left            =   3105
      TabIndex        =   1
      Top             =   2610
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddItem(msg As String)
    List1.AddItem msg
End Sub

Private Sub Command1_Click()
    List1.Clear
    List1.AddItem "I am still active and processing as normal"
End Sub

'we have to manually center form because of the way we showed it
'setting formpos center default wont work in this instance..
'but it will for frmSelectDate
Private Sub Form_Load()
    On Error Resume Next
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set f = Nothing
    Unload Me
End Sub
