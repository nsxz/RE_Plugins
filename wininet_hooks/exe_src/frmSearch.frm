VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Match"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Count"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sample Text"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Phrase"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abort As Boolean


Private Sub cmdAbort_Click()
    abort = True
End Sub

Private Sub Command1_Click()
Dim li As ListItem
Dim li2 As ListItem
On Error Resume Next

        abort = False
        
        If frmFileViewer.active_lv Is Nothing Then
            MsgBox "File viewer active listview is nothing"
            Exit Sub
        End If
        
         If Len(Text1) = 0 Then
            MsgBox "Enter search phrase"
            Exit Sub
         End If
        
        lv.ListItems.Clear
        b = IIf(frmFileViewer.active_lv Is frmFileViewer.lvPages, "c:\pages\", "c:\posts\")
        pb.value = 0
        pb.Max = frmFileViewer.active_lv.ListItems.Count
        
        For Each li In frmFileViewer.active_lv.ListItems
               tmp = fso.ReadFile(b & li.Text)
               a = InStr(1, tmp, Text1, vbTextCompare)
               If a > 0 Then
                     Set li2 = lv.ListItems.Add(, , li.Text)
                     li2.SubItems(1) = GetCount(tmp, Text1)
                     li2.SubItems(2) = Mid(tmp, a, 15)
                End If
                If abort Then Exit Sub
                pb.value = pb.value + 1
        Next
               
End Sub

Function GetCount(haystack, needle)
    On Error Resume Next
    tmp = Split(haystack, needle)
    GetCount = UBound(tmp)
End Function

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
        If frmFileViewer.active_lv Is Nothing Then
            MsgBox "File viewer active listview is nothing"
            Exit Sub
        End If

        b = IIf(frmFileViewer.active_lv Is frmFileViewer.lvPages, "c:\pages\", "c:\posts\")
        frmFileViewer.LoadFile b & Item.Text
        
End Sub
