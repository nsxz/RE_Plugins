VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProcesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Process"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "pid"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "process"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Processes"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuAnalyze 
         Caption         =   "Analyze"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Kill"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnuSignatureScanner 
         Caption         =   "Signature Scanner"
      End
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim selli As ListItem
Dim ready2return As Boolean
Dim proc As New CProcessInfo

Function getProcess() As Long
    
    Dim d As New CProcess
    Dim li As ListItem
    Dim pid As Long
    Dim cmd As String
    Dim c As Collection
    Dim f

    lv.ColumnHeaders(2).Width = lv.Width - lv.ColumnHeaders(2).Left - 300
    
    On Error Resume Next
    
    Set p = proc.GetRunningProcesses
         
    For Each d In p
        Set li = lv.ListItems.Add(, , d.pid)
        li.SubItems(1) = d.path
        li.SubItems(2) = d.User
        li.Tag = d.pid
        If pid = d.pid And pid > 0 Then Set liProc = li
    Next
        
    Me.Show
    
     ready2return = False
    While Not ready2return
        DoEvents
    Wend
    
    On Error Resume Next
    getProcess = selli.Tag
    Unload Me
    
End Function

Private Sub Command1_Click()
     ready2return = True
     
End Sub

Private Sub Command2_Click()
    Text1_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
     ready2return = True
End Sub

Private Sub lv_DblClick()
     ready2return = True
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
End Sub



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub




Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Sub Text1_Change()
    Dim li As ListItem
    For Each li In lv.ListItems
        If Not li.Selected And LCase(VBA.Left(li.SubItems(1), Len(Text1))) = Text1 Then
            li.Selected = True
            li.EnsureVisible
            Exit For
        End If
    Next
End Sub
