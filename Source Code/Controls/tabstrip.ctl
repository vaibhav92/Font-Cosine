VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   6435
   ScaleWidth      =   5040
   Begin VB.CheckBox Check2 
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4740
      Value           =   2  'Grayed
      Width           =   1155
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Stike"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4740
      Width           =   1155
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Italics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4740
      Width           =   1155
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Colours"
      Height          =   270
      Index           =   3
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4410
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Height          =   4305
      Left            =   90
      TabIndex        =   9
      Top             =   390
      Width           =   4185
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000E&
         Height          =   4215
         Left            =   0
         ScaleHeight     =   4155
         ScaleWidth      =   4125
         TabIndex        =   16
         Top             =   90
         Width           =   4185
         Begin VB.TextBox Text2 
            Height          =   4215
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "tabstrip.ctx":0000
            Top             =   -30
            Visible         =   0   'False
            Width           =   4185
         End
         Begin Project1.Charset Charset1 
            Height          =   4185
            Left            =   -60
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   7382
            FontName        =   "MS Sans Serif"
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            Height          =   195
            Left            =   1980
            TabIndex        =   17
            Top             =   1920
            Width           =   105
         End
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   225
      Index           =   2
      Left            =   4380
      TabIndex        =   8
      Text            =   "1"
      Top             =   6030
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   225
      Index           =   1
      Left            =   4380
      TabIndex        =   7
      Text            =   "1"
      Top             =   5580
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   225
      Index           =   0
      Left            =   4380
      TabIndex        =   6
      Text            =   "8"
      Top             =   5160
      Width           =   435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Size"
      Height          =   435
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5085
      Width           =   795
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Charecter"
      Height          =   435
      Index           =   2
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5940
      Width           =   795
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Weight"
      Height          =   435
      Index           =   1
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   795
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   270
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   5595
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   476
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   270
      Index           =   2
      Left            =   900
      TabIndex        =   0
      Top             =   6030
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   476
      _Version        =   393216
      Min             =   1
      Max             =   255
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   300
      Index           =   0
      Left            =   900
      TabIndex        =   4
      Top             =   5145
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   529
      _Version        =   393216
      Min             =   8
      Max             =   875
      SelStart        =   8
      TickStyle       =   3
      Value           =   8
   End
   Begin Project1.Color_pallete Col1 
      Height          =   3915
      Left            =   4380
      TabIndex        =   10
      Top             =   480
      Width           =   550
      _ExtentX        =   979
      _ExtentY        =   6906
      BackColor       =   0
      ForeColor       =   16777215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6375
      Left            =   30
      TabIndex        =   15
      Top             =   90
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   11245
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
            Caption         =   "Sample"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim f_name As String
Dim contro As Control
Private Sub Check1_Click(index As Integer)
If Check1(index).Value = 1 Then
Select Case index
Case 0
Label1.fontsize = Slider1(index).Value
Call resi
Charset1.fontsize = Slider1(index).Value
Text2.fontsize = Slider1(index).Value
Case 1
Label1.font.Weight = Slider1(index).Value
Charset1.font.Weight = Slider1(index).Value
Text2.font.Weight = Slider1(index).Value
Case 2
Label1.caption = Chr(Slider1(index).Value)
Charset1.caption = Chr(Slider1(index).Value)
End Select
End If
End Sub
Private Sub Check2_Click(index As Integer)
If Check2(index).Value = 1 Then
If index = 0 Then
contro.FontItalic = True
ElseIf index = 2 Then
contro.FontUnderline = True
Else
contro.FontStrikethru = True
End If
Else
If index = 0 Then
contro.FontItalic = False
ElseIf index = 2 Then
contro.FontUnderline = False
Else
contro.FontStrikethru = False
End If
End If
End Sub
Private Sub Color_pallete1_Click()
j = change_font(f_name)
End Sub
Private Sub Col1_Click()
contro.ForeColor = Col1.ForeColor
contro.BackColor = Col1.BackColor
If Check1(3).Value = 1 Then
Picture1.BackColor = Col1.BackColor
Label1.ForeColor = Col1.ForeColor
Text2.ForeColor = Col1.ForeColor
Text2.BackColor = Col1.BackColor
Charset1.ForeColor = Col1.ForeColor
Charset1.BackColor = Col1.BackColor
End If
End Sub
Private Sub Slider1_Scroll(index As Integer)
Text1(index).Text = Slider1(index).Value
contro.fontsize = Slider1(0).Value
contro.font.Weight = Slider1(1).Value
If Slider1(2).Enabled = True Then
contro.caption = Chr(Slider1(2).Value)
End If
If Check1(index).Value = 1 Then
Select Case index
Case 0
Label1.fontsize = Slider1(index).Value
Call resi
Charset1.fontsize = Slider1(index).Value
Text2.fontsize = Slider1(index).Value
Case 1
Label1.font.Weight = Slider1(index).Value
Charset1.font.Weight = Slider1(index).Value
Text2.font.Weight = Slider1(index).Value
Case 2
Label1.caption = Chr(Slider1(index).Value)
Charset1.caption = Chr(Slider1(index).Value)
End Select
End If
If TabStrip1.SelectedItem.index = 1 Then Call resi
End Sub
Private Sub TabStrip1_Click()
Static pre_index, i As Control
If TabStrip1.SelectedItem.index = pre_index Then Exit Sub
contro.Visible = False
Select Case TabStrip1.SelectedItem.index
Case 1
Set i = Label1
Slider1(2).Enabled = True
Case 2
Set i = Charset1
Slider1(2).Enabled = True
Case 3
Set i = Text2
Slider1(2).Enabled = False
Case 4
Exit Sub
End Select
i.Visible = True
Set contro = i
pre_index = TabStrip1.SelectedItem.index
End Sub
Private Sub Text1_Change(index As Integer)
Slider1(index).Value = Val(Text1(index).Text)
End Sub
Public Property Let fname(new_fname As String) ' i.e the font name
f_name = new_fname
j = change_font(f_name)
End Property
Private Sub UserControl_Initialize()
Static i As Boolean
If i = False Then
Set contro = Label1
i = True
End If
End Sub
Private Sub resi()
Static i As Integer, j As Integer
If i = 0 And j = 0 Then
i = Label1.Top
j = Label1.Left
End If
ii = Label1.fontsize
Label1.Top = i - 12 * ii
Label1.Left = j - 6 * ii
End Sub
Public Property Let font(new_font As String)
Label1.font.Name = new_font
Charset1.FontName = new_font
Text1(1).font.Name = new_font

PropertyChanged "font"
End Property


