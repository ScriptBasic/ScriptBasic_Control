VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAryDump 
   Caption         =   "Form2"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lv 
      Height          =   4515
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   7964
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAryDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selVariable As ListItem
Dim g_varName As String

Private Sub Form_Load()

    With lv
        .ColumnHeaders(.ColumnHeaders.Count).Width = .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - 100
    End With
    
End Sub

Sub DumpArrayValues(varName As String, c As Collection)
    Dim v As CVariable
    Dim li As ListItem
    
    g_varName = varName
    lv.ListItems.Clear
    
    For Each v In c
        Set li = lv.ListItems.Add(, , v.index)
        li.SubItems(1) = v.varType
        li.SubItems(2) = v.value
        If v.varType = "array" Then li.Tag = v.pAryElement
    Next
    
    Me.Caption = "Array Dump: " & varName
    Me.Show
    
End Sub

Private Sub lv_DblClick()

    If selVariable Is Nothing Then Exit Sub
    If selVariable.SubItems(1) <> "array" Then Exit Sub
    
    Dim c As Collection
    Dim varName As String
    Dim f As frmAryDump
    
    varName = g_varName & "[" & selVariable.Text & "]"
    Set c = EnumArrayVariables(selVariable.Tag)
    If c.Count > 0 Then
        Set f = New frmAryDump
        f.DumpArrayValues varName, c
        f.Move f.Left + 300, f.Top + 300
    End If
    
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
     Set selVariable = Item
End Sub
