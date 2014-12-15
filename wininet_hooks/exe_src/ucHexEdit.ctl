VERSION 5.00
Begin VB.UserControl ucHexEdit 
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12705
   ScaleHeight     =   297.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   635.25
   Begin VB.VScrollBar scroll 
      Height          =   5820
      Left            =   10800
      Max             =   10
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDisp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5820
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   288
      ScaleMode       =   2  'Point
      ScaleWidth      =   531.75
      TabIndex        =   3
      Top             =   0
      Width           =   10695
   End
   Begin VB.TextBox txtChr 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ucHexEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'http://geocities.com/dzzie
'------------------------------------------------------------------------


Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const EM_GETSEL = &HB0

Dim fpath As String
Dim file() As Byte
Dim pageChanges() As String
Dim SelStart As Long
Dim SelLength As Long
Dim Dirty As Boolean


Private Sub UserControl_Initialize()
    SelStart = -1
    'loadfile App.path & "\flp.tmp"
    
    picDisp.FontSize = 9 '<--really really needs to be 9 !
                         'all my math depends on all characters being 9 high and 6 wide
                         'and the hexdump format being exactly the way it is..dont say i
                         'didnt warn yaaaaa :P
    
End Sub


Private Sub mnuOption_Click(Index As Integer)
    Select Case Index
        Case 0: SaveChanges fpath
        Case 1: SaveChanges 'will prompt for path
        Case 2:
                a = InputBox("Enter Hex Offset to jump to, note it will be stepped to &H200 boundry")
                a = RoundUp(cHex(a), &H200)
                a = a / &H200
                If a > scroll.Max Then scroll.value = scroll.Max Else scroll.value = a
    End Select
End Sub

'Private Sub picDisp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If fpath = Empty Then Exit Sub
'
'    sx = SnapX(x)
'    sy = snapY(y)
'    curoffset = GetOffsetFromEitherGrid(sx, sy)
'
'    If Button = 1 Then               'left click =edit
'         If Shift = 0 Then ResetAll  'meant for internal uses
'         If sx = Empty Then Exit Sub 'click outside hexdata area
'         If curoffset = -1 Then ResetAll True: Exit Sub 'at end of file
'         If x < 350 Then EditByte curoffset Else EditChar sx, sy
'         SelectByte curoffset
'    Else
'        'PopupMenu mnuPopup
'    End If
'
'End Sub

Private Sub picDisp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Caption = SnapX(x) & " " & snapY(y) & " " & x & " " & y
End Sub

Private Sub picDisp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    f = Data.Files(1)
    If FileExists(f) Then LoadFile Data.Files(1)
End Sub

Private Sub scroll_Change()
   On Error GoTo oops
    picDisp.Cls
    txtEdit.Visible = False
    txtChr.Visible = False
    sv = scroll.value 'scroll.value * &h200 = overflow if >= 64 !
    picDisp.Print HexDumpByteArray(file(), sv * &H200, &H1FF)
    ShowChanges
   Exit Sub
oops: 'yes i am a big fat cheater !
      If scroll.value = scroll.Max Then
        scroll.value = scroll.value - 1
        scroll.Max = scroll.value
      End If
End Sub

Private Sub txtchr_Change()
    i = SendMessageLong(txtChr.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    If i = 1 Then
        Dirty = True
        tmp = GetGridFromOffset(SelStart) 'set when box shown
        x = txtChr.Left
        y = snapY(txtChr.Top)
        h = Hex(Asc(txtChr))
        If Len(h) = 1 Then h = "0" & h
        RemberChange SelStart
        OverWrite x, y, txtChr, vbYellow, vbBlack
        OverWrite tmp(0), tmp(1), h & " ", vbYellow, vbBlack
        ChangeByteFromGrid tmp(0), tmp(1), h
        tmp = GetCharGridFromOffset(SelStart + 1)
        'picDisp_MouseDown 1, 1, CSng(tmp(0)), CSng(tmp(1))
    End If
End Sub

Private Sub txtEdit_Change()
    i = SendMessageLong(txtEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    If i = 2 Then
        Dirty = True
        x = txtEdit.Left
        y = snapY(txtEdit.Top)
        RemberChange SelStart 'SelStart set when box shown
        tmp = GetCharGridFromOffset(SelStart)
        OverWrite tmp(0), tmp(1), GetDisplayChar(txtEdit), vbYellow, vbBlack
        OverWrite x, y, txtEdit.Text & " ", vbYellow, vbBlack
        ChangeByteFromGrid x, y, txtEdit
        tmp = GetGridFromOffset(SelStart + 1)
        'picDisp_MouseDown 1, 1, CSng(tmp(0)), CSng(tmp(1))
    End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 33: 'page up
                 If scroll.value > 0 Then scroll.value = scroll.value - 1
        Case 34: 'pagedown
                 If scroll.value <> scroll.Max Then scroll.value = scroll.value + 1
        Case 38: 'uparrow
        Case 40: 'downarrow
    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    KeyAscii = FilterHexKey(KeyAscii)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
        (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")) Then
        i = SendMessageLong(txtEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
        If i > 1 Then i = 1
        txtEdit.SelStart = i
        txtEdit.SelLength = 1
    End If
End Sub

Private Sub txtchr_KeyPress(KeyAscii As Integer)
    i = SendMessageLong(txtEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    If i > 1 Then i = 1
    txtChr.SelStart = i
    txtChr.SelLength = 1
End Sub

Function FilterHexKey(mInkey) As Integer
    If mInkey < Asc("0") Or mInkey > Asc("9") Then
        If Not (mInkey >= Asc("A") And mInkey <= Asc("F")) Then
            If Not (mInkey >= Asc("a") And mInkey <= Asc("f")) Then
                 If mInkey <> 8 Then
                      mInkey = 0
                 End If
            End If
        End If
    End If
    If mInkey >= Asc("a") And mInkey <= Asc("f") Then
        mInkey = mInkey - 32
    End If
    FilterHexKey = mInkey
End Function

Private Sub ResetAll(Optional andSelStart = False)
    txtEdit.Visible = False: txtChr.Visible = False
    If andSelStart Then SelStart = -1 Else _
    If SelStart >= 0 Then SelectByte SelStart, False, True
End Sub

Sub SaveChanges(Optional path = Empty)
    Dirty = False
    ReDim pageChanges(scroll.Max)
    scroll_Change
    
    If path = Empty Then
        path = CmnDlg.ShowSave(App.path, AllFiles, "Save File As")
        If path = Empty Then Exit Sub
        fpath = path
        'Caption = "Editing " & fpath
    End If
    
    f = FreeFile
    Open path For Binary Access Write As f
    Put f, , file()
    Close f
End Sub

Sub LoadFile(path)
    If Dirty Then If MsgBox("File been changed since last save, would you like to save it now?", vbYesNo + vbInformation) = vbYes Then SaveChanges fpath
    picDisp.Picture = LoadPicture()
    fpath = path
    scroll.Visible = True
    'Caption = "Editing " & fpath
    
    f = FreeFile
    Open path For Binary As f
        ReDim file(1 To LOF(f))
        Get f, , file()
    Close f
    
    pages = UBound(file) / &H200
    scroll.Max = IIf(InStr(pages, "."), RoundUp(pages, 1), pages - 1)
    ReDim pageChanges(scroll.Max)
    
    'Visible = True
    picDisp.Cls
    picDisp.Print HexDumpByteArray(file(), 0, &H1FF)
End Sub

'--------------------------------------------------------------------
'Editor Api functions
'--------------------------------------------------------------------

Private Function GetByteFromGrid(x, y) As String
    'If x = Empty Then Exit Function 'used to indicate click out of bounds
    'rember file() = 1 based ! editor = 0 based !
    Dim ret As String
    offset = GetOffsetFromGrid(x, y) + 1
    ret = Hex(file(offset))
    GetByteFromGrid = IIf(Len(ret) = 1, "0" & ret, ret)
End Function

Private Function GetOffsetFromEitherGrid(x, y) As Long
    'wrapped again because i have 3 versions to correct ugh
    Dim offset As Long
    If x > 350 Then offset = CLng(GetOffsetFromCharGrid(x, y)) _
    Else offset = GetOffsetFromGrid(x, y)
    GetOffsetFromEitherGrid = IIf(offset < UBound(file), offset, -1)
End Function

Private Function GetOffsetFromGrid(x, y) As Long
    Dim offset As Long
    a = (x - 54) / 18                '54 points (9chars) before hex data starts
    topoffset = GetTopOffset()       'what page are we viewing?
    b = topoffset + ((y / 9) * 16)   '16 characters per line each 9 y points = one line
    offset = b + a                   'editor view is 0 based ! file() = 1 base
    GetOffsetFromGrid = offset
End Function

Private Function GetGridFromOffset(offset)
    topoffset = GetTopOffset()
    If offset > (topoffset + &H200) Then MsgBox "Ughh sel off page?!": Exit Function
    linesdown = (offset - topoffset) / 16
    x = ((offset Mod 16) * 18) + 54
    y = linesdown * 9
    Dim ret()
    push ret(), SnapX(x)
    push ret(), snapY(y)
    GetGridFromOffset = ret()
End Function

Private Function GetCharGridFromOffset(offset)
    topoffset = GetTopOffset()
    If offset > (topoffset + &H200) Then MsgBox "Ughh sel off page?!": Exit Function
    linesdown = (offset - topoffset) / 16
    x = ((offset Mod 16) * 6) + 360
    y = linesdown * 9
    Dim ret()
    push ret(), x
    push ret(), snapY(y)
    GetCharGridFromOffset = ret()
End Function

Private Function GetOffsetFromCharGrid(x, y)
    topoffset = GetTopOffset()
    modulus = (x - 360) / 6
    linesdown = y / 9
    base = topoffset + (linesdown * 16)
    GetOffsetFromCharGrid = base + modulus
End Function

Private Sub ChangeByteFromGrid(x, y, hexStrNewVal)
    file(GetOffsetFromGrid(x, y) + 1) = CByte("&H" & hexStrNewVal)
End Sub

Sub ChangeByteFromOffset(offset, hexstrValue)
    file(offset + 1) = CByte("&h" & hexstrValue)
End Sub

Private Sub OverWrite(x, y, Data, Optional bc = -1, Optional fc = -1)
    With picDisp
        If x = Empty Then Exit Sub
        orig = .ForeColor
         c = Array("M", "Z", "T") 'these 3 will overwrite all areas of block
         For i = 0 To 2
            .CurrentX = x
            .CurrentY = y
            .ForeColor = IIf(bc = -1, .BackColor, bc)
            picDisp.Print String(Len(Data), c(i))
         Next
        .CurrentX = x
        .CurrentY = y
        .ForeColor = IIf(fc = -1, orig, fc)
        picDisp.Print Data
        .ForeColor = orig
    End With
End Sub

Private Function snapY(it)
    snapY = RoundUp(it, 9)
End Function

Private Function GetTopOffset()
    sv = scroll.value     'bastard overflows on high offsets cause of this ! even when just multiplying and not savign to it !
    GetTopOffset = sv * &H200        'scrolls at &h200 pages
End Function

Private Function SnapX(it)
    Dim x As Integer 'characters are fixed at 6 points wide..hexbytes=2 +space
    x = CInt(it)     ' > 360 = char edit mode
    If x < 54 Or (x > 340 And x < 360) Then Exit Function 'x=empty = marker for some actions
    If x < 340 Then SnapX = RoundUp(x, 18) Else SnapX = RoundUp(x, 6)
End Function

Private Sub RemberChange(offset)
    'array has as many elements as there are pages
    'each element = comma delimited list of offsets changed
    pageChanges(scroll.value) = pageChanges(scroll.value) & offset & ","
End Sub

Sub ShowChanges()
    If pageChanges(scroll.value) = Empty Then Exit Sub
    t = Split(pageChanges(scroll.value), ",")
    For i = 0 To UBound(t) - 1
        SelectByte CLng(t(i)) 'boy is that clng necessary! through in 5min bug hunt cause of datatype dont even mumble option explicit bug hunting is half the ughh fun :P
    Next
End Sub

Sub SelectByte(offset, Optional selected = True, Optional char2 = True)
    If selected Then SelStart = offset 'Else SelStart = -1
    byteval = HexString(file(), offset + 1)
    If selected Then bc = vbYellow: fc = vbBlack Else bc = -1: fc = -1
    tmp = GetGridFromOffset(offset)
    OverWrite tmp(0), tmp(1), byteval & " ", bc, fc
    If char2 Then
        tmp = GetCharGridFromOffset(offset)
        OverWrite tmp(0), tmp(1), GetDisplayChar(byteval), bc, fc
    End If
End Sub

Private Sub EditByte(offset)
    If (offset + 1) > UBound(file) Then Exit Sub
    SelectByte offset
    tmp = GetGridFromOffset(offset)
    txtEdit.Move tmp(0), tmp(1) + picDisp.Top + 1, 15, 9
    txtEdit = HexString(file(), offset + 1)
    txtEdit.Visible = True
    txtEdit.SetFocus
End Sub

Private Sub EditChar(x, y)
        offset = GetOffsetFromCharGrid(x, y)
        If (offset + 1) > UBound(file) Then Exit Sub
        SelectByte offset 'this sets selstart
        txtChr.Move x, y + picDisp.Top + 1, 6, 9
        txtChr = GetDisplayChar(HexString(file(), offset + 1))
        txtChr.Visible = True
        txtChr.SetFocus
End Sub

'-------------------------------------------------------------------
'hex editor formatting functions
'-------------------------------------------------------------------
Private Function HexDumpByteArray(ary() As Byte, offset, length) As String
    Dim strArray() As String, x As Variant
    'ary = base 1 byte array offset and length assume base 0 numbers!
    length = length + 1 ' editor display is base 0 need base1 for array
    If offset = UBound(ary) Then MsgBox "Ughh one page to far man"
    If offset + length > UBound(ary) Then length = UBound(ary) - offset
    
    ReDim strArray(1 To length)
    For i = (offset + 1) To (offset + length)
      x = x + 1
      strArray(x) = Hex(ary(i))
      If Len(strArray(x)) = 1 Then strArray(x) = "0" & strArray(x)
    Next
    HexDumpByteArray = HexDump(strArray, offset)
End Function

Private Function HexDump(ary, ByVal offset) As String
    Dim s() As String, chars As String, tmp As String
    
    If offset > 0 And offset Mod 16 <> 0 Then MsgBox "Hexdump isnt being used right! Offset not on boundry"

    'i am lazy and simplicity rules, make sure offset read
    'starts at standard mod 16 boundry or all offsets will
    'be wrong ! it is okay to read a length that ends off
    'boundry though..that was easy to fix...
    
    chars = "   "
    For i = 1 To UBound(ary)
        tmp = tmp & ary(i) & " "
        x = CInt("&h" & ary(i))
        chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            push s, h & "   " & tmp & chars
            offset = offset + 16:  tmp = Empty: chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        h = Hex(offset)
        While Len(h) < 6: h = "0" & h: Wend
        h = h & "   " & tmp
        While Len(h) <= 56: h = h & " ": Wend
        push s, h & chars
    End If
    
    HexDump = Join(s, vbCrLf)
End Function

Private Function GetDisplayChar(hIt)
    x = CLng("&h" & hIt)
    GetDisplayChar = IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
End Function

Private Function HexString(it() As Byte, offset) As String
    Dim ret As String
    ret = Hex(it(offset))
    If Len(ret) = 1 Then ret = "0" & ret
    HexString = ret
End Function


'-------------------------------------------------------------------
' Library functions
'-------------------------------------------------------------------
Private Function cHex(v) As Long
    On Error Resume Next
    cHex = CLng("&h" & v)
End Function

Private Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

'ok so this one is misnamed and my math is stinkey poo poo :(
Private Function RoundUp(s, step)
    Dim r As Long
    r = CLng(s)
    If r Mod step <> 0 Then
        If r < step Then
            r = 0
        Else
            r = r - step
            While r Mod step <> 0: r = r + 1: Wend
        End If
    End If
    RoundUp = r
End Function


