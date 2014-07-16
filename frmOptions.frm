VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Paths"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About Scivb"
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   1170
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6795
      TabIndex        =   6
      Top             =   1125
      Width           =   1230
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   330
      Index           =   1
      Left            =   7560
      TabIndex        =   5
      Top             =   585
      Width           =   420
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   330
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   135
      Width           =   420
   End
   Begin VB.TextBox txtExtDir 
      Height          =   330
      Left            =   1530
      TabIndex        =   3
      Top             =   585
      Width           =   5865
   End
   Begin VB.TextBox txtIncDir 
      Height          =   330
      Left            =   1530
      TabIndex        =   2
      Top             =   135
      Width           =   5865
   End
   Begin VB.Label Label2 
      Caption         =   "Extension Directory"
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "Include Directory"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1410
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
    frmMain.scivb.ShowAbout
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim path  As String
    path = dlg.FolderDialog()
    If Len(path) = 0 Then Exit Sub
    If Index = 0 Then txtIncDir = path Else txtExtDir = path
End Sub

Private Sub cmdSave_Click()
    includeDir = txtIncDir
    moduleDir = txtExtDir
    If Right(includeDir, 1) <> "\" Then includeDir = includeDir & "\"
    If Right(moduleDir, 1) <> "\" Then moduleDir = moduleDir & "\"
    Call SaveMySetting("includeDir", includeDir)
    Call SaveMySetting("moduleDir", moduleDir)
    InitIntellisense includeDir
    SetDefaultDirs includeDir, moduleDir
    Unload Me
End Sub

Private Sub Form_Load()
    txtIncDir = includeDir
    txtExtDir = moduleDir
End Sub
