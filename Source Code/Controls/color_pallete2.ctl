VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl Color_pallete 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ScaleHeight     =   2910
   ScaleWidth      =   3885
   ToolboxBitmap   =   "color_pallete2.ctx":0000
   Begin VB.PictureBox Color_box 
      Height          =   2055
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   900
      TabIndex        =   1
      Top             =   0
      Width           =   960
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000014&
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   0
      Width           =   600
      Begin VB.Label lbl_fore 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog com 
      Left            =   840
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Color_pallete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Event click(what As String)
Dim num_col As Long, num_rows As Long, first_time As Boolean
'Private Function getmatlen() As Long
'Dim j As Variant
'j = GetAllSettings("Color Pallete", "Matrix")
'getmatlen = UBound(j, 1)
'End Function
Private Sub label1_DblClick(Index As Integer)
com.Flags = cdlCCRGBInit
com.Color = Label1(Index).backcolor
com.ShowColor
Label1(Index).backcolor = com.Color
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lbl_fore(0).backcolor = Label1(Index).backcolor: RaiseEvent click("fore")
If Button = 2 Then Picture1.backcolor = Label1(Index).backcolor: RaiseEvent click("Back")
End Sub
Private Sub lbl_fore_Click(Index As Integer)
Call Picture1_Click
End Sub
Private Sub Picture1_Click()
Dim i As OLE_COLOR
i = lbl_fore(0).backcolor
lbl_fore(0).backcolor = Picture1.backcolor
Picture1.backcolor = i
RaiseEvent click("both")
End Sub
Private Sub get_colours()
Dim j As Variant, jj As Long, my_col(2), i As Long
j = GetAllSettings("Color Pallete", "Matrix")
For i = 0 To UBound(j, 1)
If i = Label1.Count Then Exit For
Label1(i).backcolor = j(i, 1)
'Label1(i).backcolor = RGB(CLng(Left(j(i, 1), 3)), CLng(Mid(j(i, 1), 3, 3)), CLng(Right(j(i, 1), 3)))
Next
For jj = i To (Label1.Count - 1)
my_col(0) = Int((255) * Rnd)
my_col(1) = Int((255) * Rnd)
my_col(2) = Int((255) * Rnd)
Label1(jj).backcolor = RGB(my_col(0), my_col(1), my_col(2))
Next jj
End Sub
Private Sub UserControl_Initialize()
first_time = CBool(Val(GetSetting("Color Pallete", "options", "first_time", "1")))
num_col = 1
num_rows = 1
Call lo
End Sub
Sub lo()
'On Error GoTo errhand
Dim i As Integer
For i = Label1.Count To (num_col * num_rows) - 1
Load Label1(i)
Next i
'errhand:
'If Err.Number = 360 Then Err.Clear: Resume Next Else MsgBox "error"
End Sub
Private Sub UserControl_resize()
With UserControl
Picture1.Width = .Height
Picture1.Height = .Height
Color_box.Width = .Width - (Picture1.Width)
Color_box.Height = .Height - 30
Color_box.Left = Picture1.Width
lbl_fore(0).Height = Picture1.Height / 2
lbl_fore(0).Width = Picture1.Width / 2
lbl_fore(0).Top = (Picture1.ScaleHeight - lbl_fore(0).Height) / 2
lbl_fore(0).Left = (Picture1.ScaleWidth - lbl_fore(0).Width) / 2
End With
'Picture1.Left = 0
'Picture1.Left = Color_box.Width + 30
Label1(0).Left = 0
Label1(0).Top = 0
Call aling
End Sub
Private Sub aling()
Label1(0).Height = Color_box.Height / num_rows
Label1(0).Width = Color_box.Width / num_col
Dim my_id As Integer, i As Integer, j As Integer
For i = 0 To num_rows - 1
For j = 0 To num_col - 1
Let my_id = (i * num_col) + j
With Label1(my_id)
.Top = (i) * Label1(0).Height
.Left = j * Label1(0).Width
.Height = Label1(0).Height
.Width = Label1(0).Width
.Visible = True
End With
Next j
Next i
If first_time = True Then
Dim my_col(2) As Long
For i = 0 To (Label1.Count - 1)
my_col(0) = Int((255) * Rnd)
my_col(1) = Int((255) * Rnd)
my_col(2) = Int((255) * Rnd)
'Label1(my_id).backcolor = QBColor(my_col)
Label1(i).backcolor = RGB(my_col(0), my_col(1), my_col(2))
Next
Else
Call get_colours
End If
End Sub
Public Property Let forecolor(new_color As OLE_COLOR)
lbl_fore(0).backcolor = new_color
PropertyChanged "Forecolor"
End Property
Public Property Get forecolor() As OLE_COLOR
forecolor = lbl_fore(0).backcolor
'PropertyChanged "Forecolor"
End Property
Public Property Let backcolor(new_color As OLE_COLOR)
Picture1.backcolor = new_color
PropertyChanged "Backcolor"
End Property
Public Property Get backcolor() As OLE_COLOR
backcolor = Picture1.backcolor
'PropertyChanged "Forecolor"
End Property
Public Sub dimentions(n_rows As Long, n_cols As Long)
num_rows = n_rows
num_col = n_cols
Call lo
Call aling
'PropertyChanged "dimentions"
End Sub
Public Property Let Rows(new_rows As Long)
num_rows = new_rows
Call lo
Call aling
PropertyChanged "Rows"
End Property
Public Property Let columns(new_col As Long)
num_col = new_col
Call lo
Call aling
PropertyChanged "Columns"
End Property
Public Property Get Rows() As Long
Rows = num_rows
End Property
Public Property Get columns() As Long
columns = num_col
End Property
Private Sub UserControl_Terminate()
Dim i As Long
If first_time = True Then SaveSetting "Color Pallete", "options", "first_time", "0"
If first_time = False Then DeleteSetting "Color Pallete", "Matrix"
For i = 0 To (Label1.Count - 1)
SaveSetting "Color Pallete", "Matrix", Str(i), Str(Label1(i).backcolor)
Next
End Sub
Public Property Let enabled(new_val As Boolean)
Dim i As Long
For i = 0 To Label1.Count - 1
Label1(i).enabled = new_val
Next i
Picture1.enabled = new_val
lbl_fore(0).enabled = new_val
'PropertyChanged "enabled"
End Property
Public Property Get enabled() As Boolean
enabled = Picture1.enabled
End Property

