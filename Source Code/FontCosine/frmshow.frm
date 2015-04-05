VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmshow 
   Caption         =   "Forrmshow"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   3615
   Icon            =   "frmshow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4470
   ScaleWidth      =   3615
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Min             =   8
      Max             =   500
      SelStart        =   8
      TickStyle       =   3
      Value           =   8
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000002&
      Caption         =   "Check2"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmshow.frx":030A
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin FontCosine.Charset Charset1 
      Height          =   2175
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3836
   End
   Begin VB.CheckBox lock1 
      BackColor       =   &H00FF0000&
      Caption         =   "S"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      MouseIcon       =   "frmshow.frx":0310
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmshow.frx":061A
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   1800
         Width           =   105
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Single"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Charecter Set"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sample Text"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private read_only As Boolean
Private con As New Collection
Private texta As String, _
cur_con As Integer
Public my_id As String
Private my_mdi As mdifrmmain
Private font_name As String, font_size As Integer
Private font_bold As Boolean, font_italics As Boolean
Private font_strikethru As Boolean, font_under As Boolean
Private my_fore_col As OLE_COLOR, my_back_col As OLE_COLOR
Dim moving As Boolean, xx As Long, yy As Long
Private Sub Charset1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If my_mdi.do_cloning Then Call Picture1_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub Form_Activate()
Dim i As frmshow
If my_mdi.do_copy Then
Me.fontbold_ = my_mdi.clu.fontbold_
Me.fontname_ = my_mdi.clu.fontname_
Me.fontitalic_ = my_mdi.clu.fontitalic_
Me.fontstrikethru_ = my_mdi.clu.fontstrikethru_
Me.fontsize_ = my_mdi.clu.fontsize_
Me.fontunderline_ = my_mdi.clu.fontunderline_
Me.forecolor_ = my_mdi.clu.forecolor_
Me.backcolor_ = my_mdi.clu.backcolor_
my_mdi.do_copy = False
my_mdi.MousePointer = 1
my_mdi.Toolbar1.Buttons("styleCopy").Value = tbrUnpressed
End If
Call Form_GotFocus
If my_mdi.do_cloning Then
Set i = my_mdi.create_new_view
i.fontbold_ = my_mdi.clu.fontbold_
i.fontname_ = my_mdi.clu.fontname_
i.fontitalic_ = my_mdi.clu.fontitalic_
i.fontstrikethru_ = my_mdi.clu.fontstrikethru_
i.fontsize_ = my_mdi.clu.fontsize_
i.fontunderline_ = my_mdi.clu.fontunderline_
i.forecolor_ = my_mdi.clu.forecolor_
i.backcolor_ = my_mdi.clu.backcolor_
i.TabStrip1.Tabs(cur_con).Selected = True
'.SelectedItem.Index = cur_con
my_mdi.class_col.Add i, i.my_id
Set my_mdi.clu = i
my_mdi.do_cloning = False
i.Show
my_mdi.MousePointer = 1
my_mdi.Toolbar1.Buttons("Clone").Value = tbrUnpressed
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = TabStrip1.SelectedItem.Index
        If i = TabStrip1.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
        Else
            'increment the tab
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(i + 1)
        End If
    End If
If Shift = vbAltMask And KeyCode = 83 Then
my_mdi.ListView1(0).SetFocus
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If cur_con <> 3 Then
Label1.Caption = Chr(KeyAscii)
Charset1.text = Chr(KeyAscii)
Else
Text1.SetFocus
End If
End Sub
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 And my_mdi.do_cloning = True Then Call Form_Activate
'End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then moving = True: xx = X: yy = Y
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If moving Then Label1.Move Label1.Left + X - xx, Label1.Top + Y - yy
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.refresh
Picture1.refresh
If Button = 1 Then moving = False
End Sub
'************
Property Get utext() As String
If cur_con = 1 Then
utext = Label1.Caption
Else
utext = con(cur_con).text
End If
End Property
Property Let readonly(new_va As Boolean)
Check1.Value = Abs(CLng(new_va))
End Property
Property Get readonly() As Boolean
readonly = read_only
End Property
Property Get text() As String
If cur_con = 1 Then
RichTextBox1.text = Label1.Caption
Else
RichTextBox1.text = con(cur_con).text
End If
Set RichTextBox1.Font = con(cur_con).Font
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.text)
RichTextBox1.SelColor = con(cur_con).forecolor
text = RichTextBox1.TextRTF
End Property
Property Let text(new_text As String)
If Not (read_only) Then
Label1.Caption = Left(new_text, 1)
Charset1.text = Left(new_text, 1)
Text1.text = new_text
Call relocate
End If
End Property
Property Set mdi(md As mdifrmmain)
Set my_mdi = md
End Property
Property Get mdi() As mdifrmmain
Set mdi = my_mdi
End Property
Public Property Get locked() As Boolean
locked = CBool(lock1.Value)
End Property
Public Sub refresh_()
Dim i As frmshow
Call con(cur_con).refresh
If cur_con = 1 Then _
Picture1.refresh
If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.refresh
i.Tag = ""
End If
End If
Next
End If
End If
End Sub
Public Property Let fontsize_(ByVal new_size As Long)
If read_only Then Exit Property
font_size = new_size
Slider1.Value = new_size
End Property
Public Property Let fontbold_(ByVal new_val As Boolean)
Dim i As frmshow
If read_only Then Exit Property
con(cur_con).fontbold = new_val
If cur_con = 1 Then Call relocate
font_bold = new_val

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.fontbold_ = Me.fontbold_
i.Tag = ""
End If
End If
Next
End If
End If

End Property
Public Property Let fontitalic_(ByVal new_val As Boolean)
Dim i As frmshow
If read_only Then Exit Property
con(cur_con).fontitalic = new_val
If cur_con Then relocate
font_italics = new_val

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.fontitalic_ = Me.fontitalic_
i.Tag = ""
End If
End If
Next
End If
End If
End Property
Public Property Let fontname_(ByVal New_Font As String)
Dim i As frmshow
On Error Resume Next
If read_only Then Exit Property
'*****
font_name = New_Font
If cur_con = 1 Then
Call relocate(2)
Else
con(cur_con).fontname = New_Font
con(cur_con).refresh
End If
Check1.Caption = font_name & _
"-" & Str(font_size)
'*****
'con(cur_con).fontname = New_Font
'If cur_con = 1 Then Call relocate
'con(cur_con).refresh
'font_name = New_Font
'Check1.Caption = font_name & _
'"-" & Str(font_size)

'Call relocate

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.fontname_ = Me.fontname_
i.Tag = ""
End If
End If
Next
End If
End If
End Property
Public Property Let fontunderline_(ByVal new_val As Boolean)
Dim i As frmshow
If read_only Then Exit Property
con(cur_con).fontunderline = new_val
If cur_con = 1 Then Call relocate
font_under = new_val

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.fontunderline_ = Me.fontunderline_
i.Tag = ""
End If
End If
Next
End If
End If
End Property
Public Property Let fontstrikethru_(ByVal New_Font As Boolean)
Dim i As frmshow
If read_only Then Exit Property
con(cur_con).fontstrikethru = New_Font
If cur_con = 1 Then Call relocate
font_strikethru = New_Font
'Call relocate

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.fontstrikethru_ = Me.fontstrikethru_
i.Tag = ""
End If
End If
Next
End If
End If
End Property
'************
Public Property Get fontsize_() As Long
fontsize_ = con(cur_con).fontsize
'Check1.Caption = fontname_ & "-" & Str(font_size)
End Property
Public Property Get fontbold_() As Boolean
fontbold_ = con(cur_con).fontbold
End Property
Public Property Get fontitalic_() As Boolean
fontitalic_ = con(cur_con).fontitalic
End Property
Public Property Get fontname_() As String
fontname_ = con(cur_con).fontname
Check1.Caption = fontname_ & "-" & Str(font_size)
End Property
Public Property Get fontunderline_() As Boolean
fontunderline_ = con(cur_con).fontunderline
End Property
Public Property Get fontstrikethru_() As Boolean
fontstrikethru_ = con(cur_con).fontstrikethru
End Property
'Public Property Let caption_(ByVal new_caption As String)
'    If Not (read_only) Then
'    Me.Caption = new_caption
'    End If
'End Property
'Public Property Get caption_() As String
'    caption_ = Me.Caption
'End Property
Public Property Let forecolor_(ByVal new_col As OLE_COLOR)
Dim i As frmshow
If read_only Then Exit Property
con(cur_con).forecolor = new_col
con(cur_con).refresh
If cur_con = 1 Then Picture1.refresh
my_fore_col = new_col

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.forecolor_ = Me.forecolor_
i.Tag = ""
End If
End If
Next
End If

End If
End Property
Public Property Let backcolor_(ByVal new_col As OLE_COLOR)
    Dim i As frmshow
    If read_only Then Exit Property
    'con(cur_con).backcolor = new_col
     If cur_con = 1 Then
     Picture1.backcolor = new_col
     'Picture1.refresh
     'con(cur_con).refresh
     Else
    con(cur_con).backcolor = new_col
     End If
     my_back_col = new_col

If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.backcolor_ = Me.backcolor_
i.Tag = ""
End If
End If
Next
End If
End If
End Property
Public Property Get forecolor_() As OLE_COLOR
forecolor_ = my_fore_col
End Property
Public Property Get backcolor_() As OLE_COLOR
backcolor_ = my_back_col
End Property
'*************
Private Sub Charset1_DblClick()
Label1.Caption = Charset1.text
TabStrip1.Tabs(1).Selected = True
Call TabStrip1_Click
End Sub
Private Sub Check1_Click()
'Me.readonly = CBool(Check1.Value)
read_only = CBool(Check1.Value)
Slider1.enabled = Not (CBool(Check1.Value))
If Check1.Value = 1 Then lock1.Value = 0
If cur_con = 1 Then Picture1.SetFocus Else con(cur_con).SetFocus
lock1.enabled = Not (CBool(Check1.Value))
End Sub
Private Sub Check1_GotFocus()
'Call Form_GotFocus
End Sub
Private Sub Form_GotFocus()
Set my_mdi.clu = Me
my_mdi.refresh_view_settings
'RaiseEvent cluchanged
End Sub
Private Sub Form_Initialize()
cur_con = 1
my_fore_col = QBColor(0)
my_back_col = QBColor(15)
con.Add Label1, "single"
con.Add Charset1, "charset"
con.Add Text1, "text"
End Sub
Private Sub Form_Load()
If my_mdi.WindowState = 0 Then Me.WindowState = 2
'Dim i As Integer
With Label1
.fontsize = 875
.fontsize = Slider1.Value
font_size = Slider1.Value
font_name = Label1.fontname
End With
With Charset1
.fontsize = 875
.fontsize = Slider1.Value
End With
With Text1
.fontsize = 875
.fontsize = Slider1.Value
End With
Check1.Caption = Label1.Font.name & "-" & Str(Slider1.Value)
With TabStrip1
.Top = 0
.Left = 0
Picture1.Left = .ClientLeft
Slider1.Left = .ClientLeft
Check1.Left = .ClientLeft
lock1.Height = Slider1.Height
lock1.Width = lock1.Height
End With
'font_name = Label1.fontname
End Sub
Private Sub Form_Resize()
On Error GoTo errhand
With TabStrip1
.Height = Me.ScaleHeight
.Width = Me.ScaleWidth
Check1.Width = .ClientWidth + 40
Picture1.Width = .ClientWidth
Slider1.Width = .ClientWidth - (lock1.Width + 20)
Picture1.Height = .ClientHeight _
- (Slider1.Height + Check1.Height + 60)
Check1.Top = .ClientTop
End With
Picture1.Top = Check1.Top + Check1.Height
Slider1.Top = Picture1.Top + Picture1.Height + 20
With Picture1
Charset1.Top = .Top
Charset1.Height = .Height
Charset1.Width = .Width
Charset1.Left = .Left
Text1.Top = .Top
Text1.Height = .Height
Text1.Width = .Width
Text1.Left = .Left
End With
lock1.Top = Slider1.Top
lock1.Left = Slider1.Left + Slider1.Width + 20
Call relocate
errhand:
If Err.Number = 380 Then Err.Clear: Resume Next
End Sub
Private Sub Label1_Click()
'Call Form_GotFocus
End Sub
Private Sub lock1_Click()
If lock1.Value = 0 Then lock1.backcolor = RGB(0, 0, 255) Else lock1.backcolor = RGB(0, 255, 0)
lock1.fontbold = CBool(lock1.Value)
If lock1.Value = 1 Then Slider1.SetFocus
'RaiseEvent locker(lock1.Value)
End Sub
Private Sub lock1_GotFocus()
'Call Form_GotFocus
End Sub
Private Sub Picture1_GotFocus()
'Call Form_GotFocus
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And my_mdi.do_cloning Then Call Form_Activate
End Sub

Private Sub Slider1_Change()
Call Slider1_Scroll
End Sub
Private Sub Slider1_GotFocus()
'Call Form_GotFocus
End Sub
Private Sub Slider1_Scroll()
Dim i As frmshow
font_size = Slider1.Value
'****
If cur_con = 1 Then
Call relocate(1)
Else
con(cur_con).fontsize = Slider1.Value
End If
'******
'con(cur_con).fontsize = Slider1.Value
'If cur_con = 1 Then
'Call relocate
'End If
Check1.Caption = _
fontname_ & "-" & Str(font_size)
'Check1.Caption = con(cur_con).fontname & _
'"-" & Str(Slider1.Value)
If (lock1.Value) Then
If Len(Me.Tag) = 0 Then
For Each i In my_mdi.class_col
If Not (i Is Me) Then
If i.locked Then
i.Tag = "121"
i.fontsize_ = Me.fontsize_
i.Tag = ""
End If
End If
Next
End If
End If
'RaiseEvent sizecha(Slider1.Value)
End Sub
Private Sub TabStrip1_Click()
Static pre As Object
If pre Is Nothing Then Set pre = Picture1
Dim sele As Long
sele = TabStrip1.SelectedItem.Index
pre.Visible = False
Select Case sele
Case 1
Picture1.Visible = True
Set pre = Picture1
Picture1.backcolor = my_back_col
cur_con = 1
'Call relocate
Case 2
Charset1.Visible = True
Set pre = Charset1
cur_con = 2
Case 3
Text1.Visible = True
Set pre = Text1
cur_con = 3
End Select
With con(cur_con)
.fontsize = font_size
.fontname = font_name
.fontitalic = font_italics
.fontstrikethru = font_strikethru
.fontbold = font_bold
.fontunderline = font_under
.backcolor = my_back_col
.forecolor = my_fore_col
End With
If cur_con = 1 Then Call relocate
End Sub
Private Sub TabStrip1_GotFocus()
If cur_con <> 1 Then con(cur_con).SetFocus
'Call Form_GotFocus
End Sub
Private Sub relocate(Optional i As Long = 0)
With Label1
.Visible = False
If Not (i = 1) Then .fontname = font_name
If Not (i = 2) Then .fontsize = font_size
.Top = (Picture1.ScaleHeight - .Height) / 2
.Left = (Picture1.ScaleWidth - .Width) / 2
.Visible = True
'.refresh '*******
Picture1.refresh '**********
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Variant
i = Me.Caption
my_mdi.class_col.Remove (my_id)
If my_mdi.class_col.Count <> 0 Then
Set my_mdi.clu = my_mdi.class_col(1)
Else
my_mdi.Chk_style(0).enabled = False
my_mdi.Chk_style(1).enabled = False
my_mdi.Chk_style(2).enabled = False
my_mdi.Chk_style(3).enabled = False
my_mdi.Color_pallete1.enabled = False
For Each i In my_mdi.ListView1(0).ListItems
i.Ghosted = True
Next
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If my_mdi.do_cloning Then Call Picture1_MouseDown(Button, Shift, X, Y)
End Sub
