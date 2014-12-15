VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFileViewer 
   Caption         =   "Wininet_Hooks File Viewer"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUnix2Dos 
      Caption         =   "Unix to Dos for text"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main LogFile"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin APiLogger.CmnDlg CmnDlg1 
      Left            =   4560
      Top             =   120
      _ExtentX        =   582
      _ExtentY        =   503
   End
   Begin APiLogger.ucHexEdit ucHexEdit1 
      Height          =   5775
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10186
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   600
      Width           =   12495
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6255
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   11033
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Hex"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvPosts 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvPages 
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Posts"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Pages"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch 
         Caption         =   "Search for string"
      End
      Begin VB.Menu mnuCombineSelected 
         Caption         =   "Combine Selected"
      End
      Begin VB.Menu mnuDeleteSelected 
         Caption         =   "Delete Selected Files"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete All Logged Data"
      End
   End
End
Attribute VB_Name = "frmFileViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public active_lv As ListView
Dim baseCaption As String
Const logfile = "c:\wininet_log.txt"

Private Sub Command1_Click()
    LoadFile logfile
End Sub

Private Sub Form_Load()
   On Error Resume Next
    
   With ucHexEdit1
        .Move Text1.Left, Text1.Top
        .Visible = False
   End With
    
   baseCaption = Me.Caption
   
   Dim li As ListItem
   Dim a() As String
   Dim b() As String
   Dim x
   
   a = fso.GetFolderFiles("c:\pages")
   b = fso.GetFolderFiles("c:\posts")
   
   For Each x In a
        Set li = lvPages.ListItems.Add(, , x)
   Next
   
   For Each x In b
        Set li = lvPosts.ListItems.Add(, , x)
   Next
   
   Command1_Click
   
    
End Sub

Private Sub lvPages_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Dim x
        x = "c:\pages\" & Item.Text
        LoadFile x
End Sub

Private Sub lvPages_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set active_lv = lvPages
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lvPosts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set active_lv = lvPosts
        PopupMenu mnuPopup
    End If
End Sub

Public Sub LoadFile(x)
    Dim buf() As Byte
    Dim ret As String
    
    If fso.FileExists(x) Then
            tmp = fso.ReadFile(x)
            
            If chkUnix2Dos.value = 1 And InStr(tmp, Chr(0)) < 1 Then
                tmp = Replace(tmp, vbCrLf, Chr(0))
                tmp = Replace(tmp, Chr(&HA), vbCrLf)
                tmp = Replace(tmp, Chr(0), vbCrLf)
            End If
            
            If InStr(tmp, Chr(0)) > 0 Then
                buf = StrConv(tmp, vbFromUnicode)
                For i = 0 To UBound(buf)
                    If i > 200 Then Exit For
                    DoEvents
                    x = buf(i)
                    ret = ret & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
                Next
                Text1 = ret
            Else
                Text1 = tmp
            End If
    
            ucHexEdit1.LoadFile x
            Me.Caption = baseCaption & " - Size: " & Len(Text1) & " bytes (0x" & Hex(Len(Text1)) & ")"
        End If
End Sub

Private Sub lvPosts_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Dim x
        x = "c:\posts\" & Item.Text
        LoadFile x
End Sub

Private Sub mnuCombineSelected_Click()
    Dim li As ListItem
    Dim bp As String
    Dim tmp As String
    
    If active_lv Is Nothing Then Exit Sub
    
    bp = IIf(active_lv Is lvPages, "c:\pages\", "c:\posts\")
    For Each li In active_lv.ListItems
        If li.selected = True Then
            tmp = tmp & fso.ReadFile(bp & li.Text)
        End If
    Next
    
                
    bp = CmnDlg1.ShowSave(App.path)
    If Len(bp) > 0 Then
        fso.WriteFile bp, tmp
        MsgBox "Combine FIle saved as: " & bp, vbInformation
    End If
             
End Sub

Private Sub mnuDeleteAll_Click()
    On Error Resume Next
    
    If MsgBox("Are you sure you want to clear all logged data?", vbYesNo + vbCritical) = vbYes Then
        fso.DeleteFile logfile
        fso.DeleteFolder "c:\pages", True
        fso.DeleteFolder "c:\posts", True
        Unload Me
    End If
    
End Sub

Private Sub mnuDeleteSelected_Click()
    Dim b As String
    
    If active_lv Is Nothing Then Exit Sub
    If MsgBox("Are you sure you want to delete the selected files ?", vbYesNo + vbCritical) = vbNo Then Exit Sub
        
    b = IIf(active_lv Is lvPages, "c:\pages\", "c:\posts\")
        Dim li As ListItem
        For i = active_lv.ListItems.Count To 1 Step -1
            Set li = active_lv.ListItems(i)
            Kill b & li.Text
            active_lv.ListItems.Remove li.Index
        Next
    

End Sub

Private Sub mnuSearch_Click()
        If active_lv Is Nothing Then Exit Sub
        Dim li As ListItem
        Dim s As String
        
       frmSearch.Show
                
            
End Sub

Private Sub TabStrip1_Click()
    Dim i
    With TabStrip1
        i = .SelectedItem.Index
        Text1.Visible = IIf(i = 1, True, False)
        ucHexEdit1.Visible = Not IIf(i = 1, True, False)
    End With
End Sub
