VERSION 5.00
Begin VB.UserControl Charset 
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   1140
   ScaleWidth      =   1455
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   0
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Tag             =   "0"
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   210
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1095
      LargeChange     =   8
      Left            =   1200
      Max             =   235
      Min             =   1
      SmallChange     =   4
      TabIndex        =   0
      Top             =   0
      Value           =   20
      Width           =   225
   End
End
Attribute VB_Name = "Charset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Dblclick()
Dim my_forecolor As OLE_COLOR, my_backcolor As OLE_COLOR
Dim cur_index As Long
Dim font_name As String, m_text As String
Dim focused As Long, font_size As Long
Dim font_italic As Boolean, font_underline As Boolean
Dim font_strikethru As Boolean, font_bold As Boolean
Public Sub refresh()
Dim i As Long
For i = 0 To 19
Label1(i).refresh
Picture1(i).refresh
Next i
End Sub
Public Property Let forecolor(new_color As OLE_COLOR)
Dim i As Integer
my_forecolor = new_color
For i = 0 To 19
If i = cur_index Then
Picture1(i).backcolor = new_color
Else
Label1(i).forecolor = new_color
End If
Next
End Property
Public Property Let backcolor(new_color As OLE_COLOR)
Dim i As Integer
my_backcolor = new_color
For i = 0 To 19
If i = cur_index Then
Label1(i).forecolor = new_color
Else
Picture1(i).backcolor = new_color
End If
Next
End Property
Public Property Let fontsize(New_Font As Long)
Dim i As Integer
font_size = New_Font
For i = 0 To 19
Label1(i).Visible = False
'Label1(i).fontsize = New_Font
Call aloalo(i, New_Font)
Label1(i).refresh
Label1(i).Visible = True
'Picture1(i).refresh
Next
'Call relo
End Property
Public Property Let fontname(New_Font As String)
Dim i As Integer
font_name = New_Font
For i = 0 To 19
Label1(i).Visible = False
Label1(i).fontname = New_Font
Label1(i).refresh
Label1(i).Visible = True
'Picture1(i).refresh
Next
Call relo
End Property
Public Property Let fontitalic(New_Font As Boolean)
Dim i As Integer
font_italic = New_Font
For i = 0 To 19
Label1(i).Visible = False
Label1(i).fontitalic = New_Font
Label1(i).Visible = True
Next
Call relo
End Property
Public Property Let fontunderline(New_Font As Boolean)
Dim i As Integer
font_underline = New_Font
For i = 0 To 19
Label1(i).Visible = False
Label1(i).fontunderline = New_Font
Label1(i).Visible = True
Next
Call relo
End Property
Public Property Let fontstrikethru(New_Font As Boolean)
Dim i As Integer
font_strikethru = New_Font
For i = 0 To 19
Label1(i).Visible = False
Label1(i).fontstrikethru = New_Font
Label1(i).Visible = True
Next
Call relo
End Property
Public Property Let fontbold(New_Font As Boolean)
Dim i As Integer
font_bold = New_Font
For i = 0 To 19
Label1(i).Visible = False
Label1(i).fontbold = New_Font
Label1(i).Visible = True
Next
Call relo
End Property
'*********
Public Property Get forecolor() As OLE_COLOR
forecolor = my_forecolor
End Property
Public Property Get backcolor() As OLE_COLOR
backcolor = my_backcolor
End Property
Public Property Get fontsize() As Long
fontsize = font_size
End Property
Public Property Get fontname() As String
fontname = font_name
End Property
Public Property Get fontitalic() As Boolean
fontitalic = font_italic
End Property
Public Property Get fontunderline() As Boolean
fontunderline = font_underline
End Property
Public Property Get fontstrikethru() As Boolean
fontstrikethru = font_strikethru
End Property
Public Property Get fontbold() As Boolean
fontbold = font_bold
End Property
'*********
Private Sub resi()
On Error Resume Next
Dim i As Integer
Dim ii As Integer
Picture1(0).Height = UserControl.Height / 5
Picture1(0).Width = (UserControl.Width - VScroll1.Width - 30) / 4
For i = 0 To 4
With Picture1(4 * i)
.Top = i * Picture1(0).Height
.Left = Picture1(0).Left
.Height = Picture1(0).Height
.Width = Picture1(0).Width
End With
For ii = ((4 * i) + 1) To ((4 * i) + 3)
With Picture1(ii)
.Top = Picture1(ii - 1).Top
.Left = Picture1(ii - 1).Left + Picture1(ii - 1).Width
.Height = Picture1(0).Height
.Width = Picture1(0).Width
End With
Next ii
Next i
Call relo
With VScroll1
.Top = Picture1(0).Top
.Left = Picture1(0).Left + (Picture1(0).Width * 4) + 30
.Height = 5 * Picture1(0).Height
End With
End Sub
Private Sub relo()
Dim i As Integer
For i = 0 To (Label1.Count - 1)
With Label1(i)
Label1(i).Visible = False
.Top = (Picture1(i).ScaleHeight - .Height) / 2
.Left = (Picture1(i).ScaleWidth - .Width) / 2
Label1(i).Visible = True
End With
Next
End Sub
Private Sub label1_DblClick(Index As Integer)
Call Picture1_DblClick(Index)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Picture1_MouseDown(Index, Button, Shift, X, Y)
End Sub
Private Sub Picture1_Click(Index As Integer)
Call Label1_Click(Index)
End Sub
Private Sub Picture1_DblClick(Index As Integer)
RaiseEvent Dblclick
End Sub
Private Sub Picture1_KeyPress(Index As Integer, KeyAscii As Integer)
Me.text = Chr(KeyAscii)
End Sub
Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_Initialize()
Dim i As Integer
For i = 1 To 19
Load Picture1(i)
Load Label1(i)
Set Label1(i).Container = Picture1(i)
Label1(i).Visible = True
Picture1(i).Visible = True
Next
Label1(0).forecolor = Picture1(2).backcolor
Picture1(0).backcolor = Label1(2).forecolor
'Call resi*******************
VScroll1.Value = 65
'Call populate(65)**********************
my_forecolor = Label1(2).forecolor
my_backcolor = Picture1(2).backcolor
End Sub
Public Property Get text() As String
text = Label1(cur_index).Caption
End Property
Public Property Let text(new_text As String)
If Trim(new_text) = "" Then new_text = "A"
VScroll1.Value = Asc(Left(new_text, 1)) - cur_index
End Property
Private Sub populate(FR As Integer)
Dim j As Integer
For j = FR To (FR + Label1.Count - 1)
Label1(j - FR).Caption = Chr(j)
Next
End Sub
Private Sub Label1_Click(Index As Integer)
Static pre As Integer
Label1(pre).forecolor = my_forecolor
Picture1(pre).backcolor = my_backcolor
Label1(Index).forecolor = my_backcolor
Picture1(Index).backcolor = my_forecolor
m_text = Label1(Index).Caption
cur_index = Index
pre = Index
RaiseEvent click
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
Me.text = Chr(KeyAscii)
End Sub
Private Sub UserControl_resize()
Call resi
End Sub
Private Sub VScroll1_Change()
Call VScroll1_Scroll
End Sub
Private Sub VScroll1_Scroll()
Call populate(VScroll1.Value)
End Sub
Public Property Get Font() As Font
        Set Font = Label1(0).Font
End Property
Private Sub aloalo(who As Integer, siza As Long)
With Label1(who)
'Label1(i).Visible = False
.fontsize = siza
.Top = (Picture1(who).ScaleHeight - .Height) / 2
.Left = (Picture1(who).ScaleWidth - .Width) / 2
'Label1(i).Visible = True
End With
End Sub
