VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "IDASrvr - OllySync UDP Bridge (listens on port 3333)"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   540
      TabIndex        =   5
      Top             =   3735
      Width           =   7890
   End
   Begin VB.CommandButton Command2 
      Caption         =   "rebind port"
      Height          =   420
      Left            =   405
      TabIndex        =   3
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear list"
      Height          =   420
      Left            =   2025
      TabIndex        =   2
      Top             =   4320
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reconnect IDA"
      Height          =   330
      Left            =   8505
      TabIndex        =   1
      Top             =   3735
      Width           =   1410
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9870
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   4140
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Idb"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   3735
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ida As New cIDAClient
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub cmdclear_Click()
    List1.Clear
End Sub

Private Sub Command1_Click()

    ida.EnumIDAWindows
    If ida.ActiveServers.Count = 0 Then
        List1.Clear
        List1.AddItem "No open IDA instances found. Do you have IDASrvr plugin installed?"
    ElseIf ida.ActiveServers.Count = 1 Then
        ida.ActiveIDA = ida.ActiveServers(1)
    Else
        ida.ActiveIDA = ida.SelectServer(True)
    End If
    Text1 = ida.LoadedFile
    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    
    sck.Close
    Err.Clear
    
    sck.LocalPort = 3333
    sck.Bind
    
    If Err.Number <> 0 Then
        List1.AddItem "Failed to bind to udp 3333"
    Else
        List1.AddItem "Now listening for IDA commands on udp 3333"
    End If
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Dim c As String, a As Long, autoConnectHWND As Long, t As String
    
    sck.LocalPort = 3333
    sck.Bind
    
    If Err.Number <> 0 Then
        List1.AddItem "Failed to bind to udp 3333"
    Else
        List1.AddItem "Now listening for IDA commands on udp 3333"
    End If
    
    ida.Listen Me.hwnd

    c = Command
    a = InStr(c, "/hwnd=")
    If a > 0 Then
        t = Mid(c, a)
        c = Trim(Replace(c, t, Empty))
        t = Trim(Replace(t, "/hwnd=", Empty))
        autoConnectHWND = CLng(t)
        If IsWindow(autoConnectHWND) = 0 Then autoConnectHWND = 0
    End If
    
    If autoConnectHWND <> 0 Then
        ida.ActiveIDA = autoConnectHWND
        Text1 = ida.LoadedFile
    Else
        Command1_Click
    End If
        
End Sub



Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Private Sub sck_DataArrival(ByVal bytesTotal As Long)
    
    On Error Resume Next
    
    Dim tmp As String
    Dim args() As String
    
    sck.GetData tmp
    List1.AddItem tmp
    
    args = Split(tmp, " ")
    
    Select Case args(0)
        Case "jmp": ida.Jump CLng(args(1))
                    'ida.QuickCall qcmSetFocusSelectLine
                    
        Case "jmpfunc": ida.Jump ida.FuncVAByName(args(1))
                        'ida.QuickCall qcmSetFocusSelectLine
                        
        Case "jmp_rva": ida.JumpRVA CLng(args(1))
                        'ida.QuickCall qcmSetFocusSelectLine
'        Case "curidb":
'                        sck.RemoteHost = sck.RemoteHostIP
'                        sck.RemotePort = 4444
'                        sck.SendData "curidb " & ida.LoadedFile & vbCrLf
    End Select
    
End Sub

