VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmshow 
   Caption         =   "Forrmshow"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   6000
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin FontCosine.Charset Charset1 
      Height          =   2175
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3836
      fontname        =   "MS Sans Serif"
      fontsize        =   13.5
      fontbold        =   -1  'True
      font            =   "frmshow.frx":0000
      fontname        =   "MS Sans Serif"
      fontsize        =   13.5
      fontbold        =   -1  'True
      font            =   "frmshow.frx":002C
      fontname        =   "MS Sans Serif"
      fontsize        =   13.5
      fontbold        =   -1  'True
      font            =   "frmshow.frx":0058
      fontname        =   "MS Sans Serif"
      fontsize        =   13.5
      fontbold        =   -1  'True
      font            =   "frmshow.frx":0084
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   270
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   476
      _Version        =   327682
      Min             =   8
      Max             =   875
      SelStart        =   8
      Value           =   8
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   960
      Width           =   975
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   2280
         Width           =   105
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4260
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Single"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Charecter Set"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Text"
            Object.Tag             =   ""
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
Public my_font As Font, text As String , _

Event igotfocus()
Private Sub labelreset()
With Label1
.Top = (Picture1.Height - Label1.Height) / 2
.Left = (Picture1.Width - Label1.Width) / 2
End With
End Sub
Private Sub Charset1_DblClick()
Label1.caption = Charset1.text
'TabStrip1.SelectedItem (1)
TabStrip1.Tabs(1).Selected = True
Call TabStrip1_Click
End Sub
Private Sub Check1_GotFocus()
Call Form_GotFocus
End Sub
Private Sub Form_GotFocus()
RaiseEvent igotfocus
End Sub
Private Sub Form_Load()
Dim i As Integer
With Label1
.FontSize = 875
.FontSize = Slider1.Value
End With
With Charset1
.FontSize = 875
.FontSize = Slider1.Value
End With
With Text1
.FontSize = 875
.FontSize = Slider1.Value
End With
Check1.caption = Label1.Font.Name & "-" & Str(Slider1.Value)
With TabStrip1
.Top = 0
.Left = 0
Picture1.Left = .ClientLeft
Slider1.Left = .ClientLeft
Check1.Left = .ClientLeft
End With
End Sub
Private Sub Form_Resize()
On Error GoTo errhand
With TabStrip1
.Height = Me.ScaleHeight
.Width = Me.ScaleWidth
Check1.Width = .ClientWidth + 40
Picture1.Width = .ClientWidth
Slider1.Width = .ClientWidth
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
Call labelreset
errhand:
If Err.Number = 380 Then Err.Clear: Resume Next
End Sub
Private Sub Label1_Click()
Call Form_GotFocus
End Sub
Private Sub Picture1_GotFocus()
Call Form_GotFocus
End Sub
Private Sub Slider1_GotFocus()
Call Form_GotFocus
End Sub
Private Sub Slider1_Scroll()
With Label1
.Visible = False
.FontSize = Slider1.Value
.Top = (Picture1.Height - Label1.Height) / 2
.Left = (Picture1.Width - Label1.Width) / 2
.Visible = True
Picture1.Refresh
Check1.caption = Label1.FontName & "-" & Str(Slider1.Value)
Charset1.FontSize = .FontSize
Text1.FontSize = .FontSize
End With
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
Case 2
Charset1.Visible = True
Set pre = Charset1
Case 3
Text1.Visible = True
Set pre = Text1
End Select
End Sub
Private Sub TabStrip1_GotFocus()
Call Form_GotFocus
End Sub
