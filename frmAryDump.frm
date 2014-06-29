VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAryDump 
   Caption         =   "Form2"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAssociative 
      Caption         =   "Display as Associative Array"
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   2355
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4110
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   7250
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
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuModifyValue 
         Caption         =   "Modify Value"
      End
   End
End
Attribute VB_Name = "frmAryDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'int __stdcall dbg_SetAryValByPointer(pDebuggerObject pDO, VARIABLE v, int index, int isLong, char* bbuf)
Private Declare Function dbg_SetAryValByPointer Lib "sb_engine" (ByVal hDebug As Long, ByVal pAry As Long, ByVal Index As Long, ByVal isLong As Long, ByVal strValue As String) As Long

Dim selVariable As ListItem
Dim g_varName As String
Dim pAryPtr As Long
Dim cVars As Collection

Private Sub chkAssociative_Click()
    DumpArrayValues g_varName, cVars, pAryPtr
End Sub

Private Sub Form_Load()

    mnuPopup.Visible = False
    
    With lv
        .ColumnHeaders(.ColumnHeaders.count).Width = .Width - .ColumnHeaders(.ColumnHeaders.count).Left - 100
    End With
    
End Sub

Sub DumpArrayValues(varName As String, c As Collection, mAryPtr As Long)
    Dim v As CVariable
    Dim li As ListItem
    
    If mAryPtr = 0 Then mAryPtr = VariableFromName(varName)
    
    pAryPtr = mAryPtr
    g_varName = varName
    Set cVars = c
    
    lv.ListItems.Clear
    
    If chkAssociative.Value = 0 Then
        For Each v In c
            Set li = lv.ListItems.Add(, , v.Index)
            li.SubItems(1) = v.varType
            li.SubItems(2) = v.Value
            Set li.Tag = v
        Next
    Else
        For i = 1 To c.count
            Set v = c(i)
            If (i - 1) Mod 2 = 0 Then
                Set li = lv.ListItems.Add(, , v.Value)
            Else
                 li.SubItems(1) = v.varType
                 li.SubItems(2) = v.Value
                 Set li.Tag = v
            End If
        Next
    End If
    
    Me.Caption = "Array Dump: " & varName
    Me.Show
    
End Sub

Private Sub lv_DblClick()

    On Error Resume Next
    
    If selVariable Is Nothing Then Exit Sub
    
    If selVariable.SubItems(1) <> "array" Then
        mnuModifyValue_Click
        Exit Sub
    End If
    
    Dim c As Collection
    Dim varName As String
    Dim f As frmAryDump
    Dim v As CVariable
    
    varName = g_varName & "[" & selVariable.Text & "]"
    Set v = selVariable.Tag
    Set c = EnumArrayVariables(v.pAryElement)
    If c.count > 0 Then
        Set f = New frmAryDump
        f.DumpArrayValues varName, c, v.pAryElement
        f.Move f.Left + 300, f.Top + 300
    End If
    
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
     Set selVariable = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuModifyValue_Click()
    
    On Error Resume Next
    
    If selVariable Is Nothing Then Exit Sub
    
    Dim v As CVariable
    Dim isNumeric As Long
    Dim c As Collection
    Dim Value As String
    Dim newVal As String
    
    Set v = selVariable.Tag
    
    If v.varType = "array" Then
        MsgBox "Can not modify a parent array object directly", vbInformation
        Exit Sub
    End If

    If v.varType = "ref" Then
        MsgBox "Can not edit ref variables edit the parent variable directly..", vbInformation
        Exit Sub
    End If
    
    Value = v.Value
    If Left(Value, 1) = """" Then Value = Mid(Value, 2)
    If Right(Value, 1) = """" Then Value = Mid(Value, 1, Len(Value) - 1)
    
    newVal = InputBox("Modify variable " & v.name, , Value)
    If Len(newVal) = 0 Or newVal = Value Then Exit Sub
    
    If Left(newVal, 2) = "0x" Then
        x = CLng("&h" & Mid(newVal, 3))
        If Err.Number = 0 Then
            newVal = x
            isNumeric = 1
        End If
    Else
        x = CLng(newVal)
        If Err.Number = 0 Then isNumeric = 1
    End If
        
    dbg_SetAryValByPointer hDebugObject, pAryPtr, v.Index, isNumeric, newVal
     
    'If InStr(g_varName, "[") > 0 Then
        'its not the top level array var, but a recirsive one..
        Set c = EnumArrayVariables(pAryPtr)
    'Else
    '    Set c = EnumArrayVariables(g_varName)
    'End If
    
    If c.count > 0 Then DumpArrayValues g_varName, c, pAryPtr
    
End Sub

