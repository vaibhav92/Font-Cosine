VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdifrmmain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "FontCosine"
   ClientHeight    =   6060
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6210
   HelpContextID   =   21005
   Icon            =   "mdifrmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   11
      Left            =   5040
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   3480
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":290A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":307A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":35BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":3B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":4046
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":458A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":46E6
            Key             =   "clone"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":4A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":4F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":548A
            Key             =   "paint"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":59CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":5F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":60AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":65F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":674E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3960
      Top             =   1680
   End
   Begin MSComDlg.CommonDialog Cda1 
      Left            =   5160
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpFile        =   "help.hlp"
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6150
      TabIndex        =   2
      Top             =   5565
      Width           =   6210
      Begin FontCosine.Color_pallete Color_pallete1 
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   767
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   5145
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   2190
      TabIndex        =   1
      Top             =   420
      Width           =   2250
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         HelpContextID   =   1
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Font Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   5055
         HelpContextID   =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   8916
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Font Box"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   5055
         Left            =   2160
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   45
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copyfont"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copyanscii"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bold"
            Style           =   4
            Object.Width           =   345
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "italic"
            Style           =   4
            Object.Width           =   345
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "underline"
            Style           =   4
            Object.Width           =   345
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "strike"
            Style           =   4
            Object.Width           =   345
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clone"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "styleCopy"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reset"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            ImageIndex      =   15
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            ImageIndex      =   17
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "backward"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clock"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.CheckBox Chk_style 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox Chk_style 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox Chk_style 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox Chk_style 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4440
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":6C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrmmain.frx":9446
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu f_ 
         Caption         =   "&New View"
         Index           =   0
      End
      Begin VB.Menu f_ 
         Caption         =   "&Close"
         Index           =   1
      End
      Begin VB.Menu f_ 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu f_vaot 
         Caption         =   "&Always on Top"
      End
      Begin VB.Menu wwtf_ 
         Caption         =   "-"
      End
      Begin VB.Menu f_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu e_ 
         Caption         =   "&Copy"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu e_ 
         Caption         =   "&Paste"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu e_ 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu e_ 
         Caption         =   "Copy &Font Name"
         Index           =   5
      End
      Begin VB.Menu e_ 
         Caption         =   "Copy &Unformatted"
         Index           =   6
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu v_box 
         Caption         =   "&Font Box"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu v_box 
         Caption         =   "&Color Box"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu v_box 
         Caption         =   "&ToolBar"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu v_w1 
         Caption         =   "-"
      End
      Begin VB.Menu font_pop 
         Caption         =   "Font &Box"
         Begin VB.Menu f_p 
            Caption         =   "&Large Icons"
            Index           =   0
         End
         Begin VB.Menu f_p 
            Caption         =   "&Small Icons"
            Index           =   1
         End
         Begin VB.Menu f_p 
            Caption         =   "Lis&t"
            Index           =   2
         End
         Begin VB.Menu f_p 
            Caption         =   "&Details"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu f_p 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu f_p 
            Caption         =   "&Arrange"
            Index           =   5
         End
         Begin VB.Menu f_p 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu f_p 
            Caption         =   "&Refresh"
            Index           =   7
         End
      End
      Begin VB.Menu v_w2 
         Caption         =   "-"
      End
      Begin VB.Menu v_refresh 
         Caption         =   "&Refresh View"
         Shortcut        =   {F5}
      End
      Begin VB.Menu reset_view 
         Caption         =   "Re&set View"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu ch_ 
      Caption         =   "&Charecters"
      Begin VB.Menu c_ 
         Caption         =   "&Upper Case"
         Index           =   0
      End
      Begin VB.Menu c_ 
         Caption         =   "&Lower Case"
         Index           =   1
      End
      Begin VB.Menu c_ 
         Caption         =   "&My chars"
         Index           =   2
         Begin VB.Menu mychars 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu c_ 
         Caption         =   "&Numbers"
         Index           =   3
      End
      Begin VB.Menu c_ 
         Caption         =   "&Accents"
         Index           =   4
      End
      Begin VB.Menu c_ 
         Caption         =   "&Symbols"
         Index           =   5
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu syms 
            Caption         =   ""
            Index           =   10
         End
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnutool_pre 
         Caption         =   "&Options"
         Index           =   0
      End
      Begin VB.Menu tw1 
         Caption         =   "-"
      End
      Begin VB.Menu prelaunch 
         Caption         =   "Charmap"
         Index           =   0
      End
      Begin VB.Menu prelaunch 
         Caption         =   "Fonts Folder"
         Index           =   1
      End
      Begin VB.Menu prelaunch 
         Caption         =   "Control Panel"
         Index           =   2
      End
      Begin VB.Menu prelaunch 
         Caption         =   "Explorer"
         Index           =   3
      End
      Begin VB.Menu tw2 
         Caption         =   "-"
      End
      Begin VB.Menu t_lpad 
         Caption         =   "&Launch Pad"
         Begin VB.Menu Launchpad 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu tw3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_fs 
         Caption         =   "Font&Show"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu w_ 
         Caption         =   "Tile &Horizontly"
         Index           =   0
      End
      Begin VB.Menu w_ 
         Caption         =   "Tile &Verticaly"
         Index           =   1
      End
      Begin VB.Menu w_ 
         Caption         =   "&Cascade"
         Index           =   2
      End
      Begin VB.Menu w_ 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuh 
         Caption         =   "FontCosine Help Topics"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuh 
         Caption         =   "Search Help On.."
         Index           =   1
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuh 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuh 
         Caption         =   "&Intellective &Homepage"
         Index           =   3
      End
      Begin VB.Menu mnuh 
         Caption         =   "&Product Support"
         Index           =   4
      End
      Begin VB.Menu mnuh 
         Caption         =   "Submit a &Bug"
         Index           =   5
      End
      Begin VB.Menu mnuh 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuh 
         Caption         =   "About FontCosine"
         Index           =   7
      End
   End
End
Attribute VB_Name = "mdifrmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
Public do_cloning As Boolean, do_copy As Boolean
Dim default_fore_color As OLE_COLOR, default_back_color As OLE_COLOR
Dim default_font_name As String, default_font_size As Long, default_char As String
Dim default_font_strike As Boolean, default_font_under As Boolean
Dim default_font_style As Long
Dim mbmoving As Boolean
Public clu As frmshow
Attribute clu.VB_VarHelpID = -1
Public class_col As New Collection, locked_clu As New Collection
Sub reset(Frm As frmshow)
Frm.forecolor_ = default_fore_color
Frm.backcolor_ = default_back_color
Frm.fontname_ = default_font_name
Frm.fontsize_ = default_font_size
Frm.fontstrikethru_ = default_font_strike
Frm.fontunderline_ = default_font_under
Frm.text = default_char
Select Case (default_font_style)
Case 0
Frm.fontbold_ = False
Frm.fontitalic_ = False
Case 1
Frm.fontbold_ = False
Frm.fontitalic_ = True
Case 2
Frm.fontbold_ = True
Frm.fontitalic_ = False
Case 3
Frm.fontitalic_ = True
Frm.fontbold_ = True
End Select
End Sub
Sub refresh_view_settings()
Color_pallete1.forecolor = clu.forecolor_
Color_pallete1.backcolor = clu.backcolor_
'Chk_style(0).Value = False
Chk_style(0).Value = Abs(CInt(clu.fontbold_))
Chk_style(1).Value = Abs(CInt(clu.fontitalic_))
Chk_style(2).Value = Abs(CInt(clu.fontunderline_))
Chk_style(3).Value = Abs(CInt(clu.fontstrikethru_))
End Sub
Public Function get_my_id() As String
Static id As Long
get_my_id = "view" & Trim(Str(id))
id = id + 1
End Function
Private Sub refresh_fonts()
Dim i As Integer
ListView1(0).ListItems.Clear
For i = 1 To Screen.FontCount
If Trim(Screen.Fonts(i)) <> "" Then ListView1(0).ListItems.Add , , Screen.Fonts(i), 1, 2
Next
End Sub
Private Sub c__Click(Index As Integer)
Select Case Index
Case 0
clu.text = UCase(clu.utext)
Case 1
clu.text = LCase(clu.utext)
Case 3
clu.text = Chr(48)
Case 4
clu.text = Chr(198)
End Select
End Sub
Private Sub Chk_style_Click(Index As Integer)
With Chk_style(Index)
Select Case Index
Case 0
clu.fontbold_ = .Value
Case 1
clu.fontitalic_ = .Value
Case 2
clu.fontunderline_ = .Value
Case 3
clu.fontstrikethru_ = .Value
End Select
End With
End Sub
Private Sub Color_pallete1_click(what As String)
Select Case LCase(what)
Case "fore"
clu.forecolor_ = Color_pallete1.forecolor
Case "back"
clu.backcolor_ = Color_pallete1.backcolor
Case "both"
clu.forecolor_ = Color_pallete1.forecolor: clu.backcolor_ = Color_pallete1.backcolor
End Select
End Sub
Private Sub f__Click(Index As Integer)
Dim i As frmshow
Select Case Index
Case 0
Set i = create_new_view()
If i Is Nothing Then Exit Sub
Set clu = i
class_col.Add clu, clu.my_id
clu.Show
Color_pallete1.enabled = True
ListView1(0).enabled = True
Chk_style(0).enabled = True
Chk_style(1).enabled = True
Chk_style(2).enabled = True
Chk_style(3).enabled = True
'clu.forecolor_ = QBColor(0)
'clu.backcolor_ = QBColor(15)
Case 1
    If Not (clu Is Nothing) Then
    Unload clu
    If class_col.Count >= 1 Then
    Set clu = class_col(1)
    Else
    ListView1(0).enabled = False
    Color_pallete1.enabled = False
    End If
    End If
Case 3
Call v_aot_Click
Case 5
Unload Me
'*******
'Unload mdifrmmain
Unload frmSplash
'Unload frmshow
Unload frmOptions
Unload frmAbout
'Set Startup.myappa = Nothing
'*******
Set Startup.myappa = Nothing
End Select
End Sub
Function create_new_view() As frmshow
On Error GoTo errhand
Dim Frm As New frmshow
Let Frm.my_id = get_my_id
Set Frm.mdi = Me
Frm.forecolor_ = default_fore_color
Frm.backcolor_ = default_back_color
Frm.fontname_ = default_font_name
Frm.fontsize_ = default_font_size
Frm.fontstrikethru_ = default_font_strike
Frm.fontunderline_ = default_font_under
Frm.text = default_char
Select Case (default_font_style)
Case 0
Frm.fontbold_ = False
Frm.fontitalic_ = False
Case 1
Frm.fontbold_ = False
Frm.fontitalic_ = True
Case 2
Frm.fontbold_ = True
Frm.fontitalic_ = False
Case 3
Frm.fontitalic_ = True
Frm.fontbold_ = True
End Select
Frm.Caption = Frm.my_id
Set create_new_view = Frm
errhand:
If Err.Number <> 0 Then MsgBox "Out Of Memory. Cannot Create More View": Set create_new_view = Nothing
End Function
Private Sub do_check(Index As Integer)
Dim i As Long
For i = 0 To 3
f_p(i).Checked = False
If i = Index Then f_p(Index).Checked = True
Next i
End Sub
Private Sub f_exit_Click()
Call f__Click(5)
End Sub
Private Sub f_p_Click(Index As Integer)
Select Case (Index)
Case 0
ListView1(0).View = lvwIcon
Case 1
ListView1(0).View = lvwSmallIcon
Case 2
ListView1(0).View = lvwList
Case 3
ListView1(0).View = lvwReport
Case 5
ListView1(0).Arrange = lvwAutoLeft
Case 7
Call refresh_fonts
End Select
Call do_check(Index)
End Sub
Private Sub f_vaot_Click()
f_vaot.Checked = Not f_vaot.Checked
Call v_aot_Click
End Sub
Private Sub Image1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static sglpos As Single
If mbmoving Then
sglpos = Image1.Left + X
If sglpos > 1500 And sglpos < 2900 Then Picture1.Width = sglpos
With TabStrip1
.Width = Picture1.ScaleWidth - Image1.Width
Image1.Left = .Left + .Width
ListView1(0).Width = .ClientWidth
ListView1(0).Top = .ClientTop
ListView1(0).Left = .ClientLeft
End With
End If
End Sub
Private Sub Image1_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbmoving = True
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbmoving = False
End Sub
Private Sub Launchpad_Click(Index As Integer)
On Error GoTo errhand
Dim i As Variant
i = Shell(Launchpad(Index).Tag, vbNormalFocus)
errhand:
If Err.Number = 53 Then MsgBox "Executable not found. Please check is path", vbExclamation: Err.Clear
End Sub
Private Sub ListView1_DblClick(Index As Integer)
If class_col.Count = 0 Then
Set clu = create_new_view()
'Set i = create_new_view()
'Set clu = i
class_col.Add clu, clu.my_id
clu.Show
Color_pallete1.enabled = True
ListView1(0).enabled = True
Chk_style(0).enabled = True
Chk_style(1).enabled = True
Chk_style(2).enabled = True
Chk_style(3).enabled = True
clu.fontname_ = ListView1(0).SelectedItem.text
End If
End Sub
Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
On Error GoTo errhand
'If class_col.Count = 0 Then
'Set clu = create_new_view()
''Set i = create_new_view()
''Set clu = i
'class_col.Add clu, clu.my_id
'clu.Show
'Color_pallete1.enabled = True
'ListView1(0).enabled = True
'Chk_style(0).enabled = True
'Chk_style(1).enabled = True
'Chk_style(2).enabled = True
'Chk_style(3).enabled = True
'End If
clu.fontname_ = Item.text
errhand:
If Err.Number = 380 Then Item.Ghosted = True: _
MsgBox "Font " & Item.text & " appears to be corrupt or damaged." & Chr(13) & _
"Removing font " & Item.text & " & Switching to default font": _
clu.fontname_ = default_font_name: ListView1(0).ListItems.Remove (Item.Index): Err.Clear
If Err.Number <> 390 And Err.Number <> 0 Then MsgBox "error": Stop
End Sub
Public Sub refresh_settings()
''On Error GoTo errhand
'Dim i As Integer, setta As Variant, ero As Boolean
'Dim j As Long, kl(10) As String
'kl(0) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "fontname")
'kl(1) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "fontsize")
'kl(2) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "forecolor")
'kl(3) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "backcolor")
'kl(4) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "fontstrike")
'kl(5) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "fontunderline")
'kl(6) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "fontstyle")
'kl(7) = EllReg.ReadRegistry(HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Options", "char")
'If kl(0) = "Not Found" Then default_font_name = "Times New Roman" _
'Else: default_font_name = kl(0)
'If kl(1) = "Not Found" Then default_font_size = 30 _
'Else default_font_size = Val(kl(1))
'If kl(2) = "Not Found" Then default_fore_color = CLng("0") _
'Else default_fore_color = CLng(kl(2))
'If kl(3) = "Not Found" Then default_back_color = CLng(" 16777215") _
'Else default_back_color = CLng(kl(3))
'If kl(4) = "Not Found" Then default_font_strike = CBool("0") _
'Else default_font_strike = CBool(kl(4))
'If kl(5) = "Not Found" Then default_font_under = CBool("0") _
'Else default_font_under = CBool(kl(5))
'If kl(6) = "Not Found" Then default_font_style = 0 _
'Else default_font_style = Val(kl(6))
'If kl(7) = "Not Found" Then default_char = "S" _
'Else default_char = kl(7)
't_lpad.enabled = True
'Dim kkl As Long
'setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Launchpad", kkl)
'Do Until setta(2) = "Not Found"
'    If kkl > (Launchpad.Count - 1) Then Load Launchpad(kkl)
'    Launchpad(kkl).Caption = setta(1)
'    Launchpad(kkl).Tag = setta(2)
'    kkl = kkl + 1
'    setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Launchpad", kkl)
'Loop
'If kkl = 0 And setta(2) = "Not Found" Then
't_lpad.enabled = False
'Else
'For j = kkl To (Launchpad.Count - 1)
'Unload Launchpad(j)
'Next
'End If
'kkl = 0
'c_(2).enabled = True
'setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Mychars", kkl)
'Do Until setta(2) = "Not Found"
'    If kkl > (mychars.Count - 1) Then Load mychars(kkl)
'    mychars(kkl).Caption = setta(1)
'    mychars(kkl).Tag = setta(2)
'    kkl = kkl + 1
'    setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\mychars", kkl)
'Loop
'If kkl = 0 And setta(2) = "Not Found" Then
'c_(2).enabled = False
'Else
'For j = kkl To (mychars.Count - 1)
'Unload mychars(j)
'Next
'End If
'************
On Error GoTo errhand
Dim i As Integer, setta As Variant, ero As Boolean
Dim j As Long
default_font_name = GetSetting(App.ProductName, "options", "fontname", "Times New Roman")
default_font_size = Val(GetSetting(App.ProductName, "options", "fontsize", "8"))
default_fore_color = CLng(GetSetting(App.ProductName, "options", "forecolor", "986895"))
default_back_color = CLng(GetSetting(App.ProductName, "options", "backcolor", "16777215"))
default_font_strike = CBool(GetSetting(App.ProductName, "options", "fontstrike", "0"))
default_font_under = CBool(GetSetting(App.ProductName, "options", "fontunderline", "0"))
default_font_style = Val(GetSetting(App.ProductName, "options", "fontstyle", "0"))
default_char = GetSetting(App.ProductName, "options", "char", "S")
t_lpad.enabled = True
setta = GetAllSettings(App.ProductName, "Launchpad")
If IsArray(setta) Then
Launchpad(0).Caption = setta(0, 0)
Launchpad(0).Tag = setta(0, 1)
For i = 1 To UBound(setta, 1)
        If ero Then Err.Clear: Exit For
        If i > (Launchpad.Count - 1) Then Load Launchpad(i)
        Launchpad(i).Caption = setta(i, 0)
        Launchpad(i).Tag = setta(i, 1)
Next
Else
t_lpad.enabled = False
End If
For j = i To (Launchpad.Count - 1)
Unload Launchpad(j)
Next
ero = False
c_(2).enabled = True
setta = GetAllSettings(App.ProductName, "mychars")
If IsArray(setta) Then
mychars(0).Caption = setta(0, 0)
'syms(0).Tag = setta(0, 1)
For i = 1 To UBound(setta, 1)
        If ero Then Err.Clear: Exit For
        If i > (mychars.Count - 1) Then Load mychars(i)
        mychars(i).Caption = setta(i, 0)
        'Launchpad(i).Tag = setta(i, 1)
Next
Else
c_(2).enabled = False
End If
For j = i To (mychars.Count - 1)
Unload mychars(j)
Next
errhand:
If Err.Number = 13 Then
ero = True
Resume Next
Else: If Err.Number = 362 Then Resume Next
End If

End Sub
Private Sub MDIForm_Initialize()
Dim i As Long
i = Val(GetSetting("FontCosine", "mainview", "State", "2"))
Me.WindowState = i
If i = 0 Then
Me.Height = Val(GetSetting("FontCosine", "mainview", "Height"))
Me.Top = Val(GetSetting("FontCosine", "mainview", "Top"))
Me.Width = Val(GetSetting("FontCosine", "mainview", "Width"))
Me.Left = Val(GetSetting("FontCosine", "mainview", "Left"))
'Me.clu.WindowState = 2
End If
End Sub
Private Sub MDIForm_Load()
'frmSplash.Show , Me
Dim i As Integer
Call refresh_settings
'********
Picture1.Width = Val(GetSetting("FontCosine", "mainview", "fontboxwidth", "2250"))
Picture1.Visible = CBool(Val(GetSetting("FontCosine", "mainview", "fontbox", "1")))
Picture2.Visible = CBool(Val(GetSetting("FontCosine", "mainview", "colorbox", "1")))
Toolbar1.Visible = CBool(Val(GetSetting("FontCosine", "mainview", "toolbar", "1")))
v_box(0).Checked = CBool(Val(GetSetting("FontCosine", "mainview", "fontbox", "1")))
v_box(1).Checked = CBool(Val(GetSetting("FontCosine", "mainview", "colorbox", "1")))
v_box(2).Checked = CBool(Val(GetSetting("FontCosine", "mainview", "toolbar", "1")))
mnu_fs.Checked = CBool(Val(GetSetting("FontCosine", "mainview", "fontshow", "0")))
f_vaot.Checked = CBool(Val(GetSetting("FontCosine", "mainview", "ontop", "0")))
If f_vaot.Checked Then
Timer3.enabled = True
'Call v_aot_Click
End If
If mnu_fs.Checked Then
Timer1.enabled = True
Toolbar1.Buttons("play").Value = tbrPressed
End If
'*******
For i = 0 To 3
With Chk_style(i)
.Top = Toolbar1.Buttons(Toolbar1.Buttons("bold").Index + i).Top
.Left = Toolbar1.Buttons(Toolbar1.Buttons("bold").Index + i).Left
.Height = Toolbar1.Buttons(Toolbar1.Buttons("bold").Index + i).Height
.Width = Toolbar1.Buttons(Toolbar1.Buttons("bold").Index + i).Width
End With
Next
Call refresh_fonts
Cda1.HelpFile = App.HelpFile
End Sub
Private Sub MDIForm_Resize()
Static loa As Boolean
Call v_box_Click(1)
Call v_box_Click(1)
TabStrip1.Height = Picture1.ScaleHeight
TabStrip1.Width = Picture1.ScaleWidth - Image1.Width
Image1.Left = TabStrip1.Left + TabStrip1.Width
ListView1(0).Width = TabStrip1.ClientWidth
ListView1(0).Height = TabStrip1.ClientHeight
ListView1(0).Top = TabStrip1.ClientTop
ListView1(0).Left = TabStrip1.ClientLeft
If loa = False Then
Dim kmi As Long, kmni As Long
kmi = Picture2.ScaleHeight
kmni = Picture2.ScaleWidth
Color_pallete1.Height = kmi
Color_pallete1.Width = kmni
Color_pallete1.dimentions 2, ((kmni / 202) - 2)
'Color_pallete1.dimentions 2, 50
syms(0).Caption = Chr(33): syms(0).Tag = Chr(33):
syms(1).Caption = Chr(63): syms(1).Tag = Chr(63)
syms(2).Caption = Chr(34): syms(2).Tag = Chr(34)
syms(3).Caption = Chr(36): syms(3).Tag = Chr(36)
syms(4).Caption = Chr(163): syms(4).Tag = Chr(163)
syms(5).Caption = Chr(153): syms(5).Tag = Chr(153)
syms(6).Caption = Chr(169): syms(6).Tag = Chr(169)
syms(7).Caption = Chr(174): syms(7).Tag = Chr(174)
syms(8).Caption = "Bullets": syms(8).Tag = Chr(149)
syms(9).Caption = "Quotes": syms(9).Tag = Chr(147)
syms(10).Caption = "Fractions": syms(10).Tag = Chr(188)
loa = True
End If
End Sub
Private Sub mnu_fs_Click()
mnu_fs.Checked = Not (mnu_fs.Checked)
If mnu_fs.Checked Then
Timer1.enabled = True
Toolbar1.Buttons("play").Value = tbrPressed
Else
Timer1.enabled = False
Toolbar1.Buttons("stop").Value = tbrPressed
End If
End Sub
Private Sub mnuh_Click(Index As Integer)
Dim success As Integer
Dim site As String
Select Case Index
Case 0
If Cda1.HelpFile = "" Then Cda1.HelpFile = App.HelpFile
Cda1.HelpCommand = cdlHelpContents
Cda1.ShowHelp
Case 1
If Cda1.HelpFile = "" Then Cda1.HelpFile = App.HelpFile
Cda1.HelpCommand = cdlHelpIndex
Cda1.ShowHelp
Case 3
site = GetSetting("FontCosine", "Urls", "homepage")
If site = "" Then site = "http://business.dencity.com/intellective"
success% = ShellToBrowser(Me, site, 0)
Case 4
site = GetSetting("FontCosine", "Urls", "support")
If site = "" Then site = "http://business.dencity.com/intellective"
success% = ShellToBrowser(Me, site, 0)
Case 5
site = GetSetting("FontCosine", "Urls", "bugs")
If site = "" Then site = "mailto:intellective@hotmail.com"
success% = ShellToBrowser(Me, site, 0)
Case 7
frmAbout.Show 1
End Select
End Sub
Private Sub mnutool_pre_Click(Index As Integer)
If f_vaot.Checked Then Call f_vaot_Click
frmOptions.Show 1, Me
Call refresh_settings
End Sub
Private Sub mychars_Click(Index As Integer)
clu.text = mychars(Index).Caption
End Sub
Private Sub prelaunch_Click(Index As Integer)
'53
On Error GoTo errhand
Dim font_folder As String
font_folder = ReadRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Fonts")
Dim i As Variant
Select Case Index
Case 0
i = Shell("charmap", vbNormalFocus)
Case 1
i = Shell("explorer.exe " & font_folder, vbNormalFocus)
Case 2
i = Shell("Control", vbNormalFocus)
Case 3
i = Shell("Explorer.EXE /n,/e,C:\", vbNormalFocus)
End Select
errhand:
If Err.Number = 53 Then MsgBox "Executable not found. Please check is path", vbExclamation: Err.Clear
End Sub
Private Sub reset_view_Click()
reset clu
End Sub
Private Sub syms_Click(Index As Integer)
clu.text = syms(Index).Tag
End Sub
Private Sub TabStrip1_GotFocus()
If ListView1(0).enabled Then ListView1(0).SetFocus
End Sub
Private Sub Timer1_Timer()
Dim i As Long
i = ListView1(0).SelectedItem.Index + 1
If i >= (ListView1(0).ListItems.Count - 1) Then i = 1
clu.fontname_ = ListView1(0).ListItems(i).text
ListView1(0).ListItems(i).Selected = True
End Sub

Private Sub Timer3_Timer()
Timer1.enabled = False
Call v_aot_Click
'Unload Timer3
'Unload Timer2
End Sub
Private Sub v_box_Click(Index As Integer)
Select Case Index
Case 0
v_box(Index).Checked = Not (v_box(Index).Checked)
Picture1.Visible = v_box(Index).Checked
Case 1
v_box(Index).Checked = Not (v_box(Index).Checked)
Picture2.Visible = v_box(Index).Checked
TabStrip1.Height = Picture1.ScaleHeight
ListView1(0).Height = TabStrip1.ClientHeight
Case 2
v_box(Index).Checked = Not (v_box(Index).Checked)
Toolbar1.Visible = v_box(Index).Checked
TabStrip1.Height = Picture1.ScaleHeight
ListView1(0).Height = TabStrip1.ClientHeight
End Select
End Sub
Private Sub v_refresh_Click()
clu.refresh_
End Sub
Private Sub v_aot_Click()
Dim xx1 As Single, yy1 As Single
xx1 = Screen.TwipsPerPixelX
yy1 = Screen.TwipsPerPixelY
    'f_vaot.Checked = Not f_vaot.Checked
    If f_vaot.Checked Then
     
     Startup.SetWindowPos hwnd, conHwndTopmost, _
     (Me.Left / xx1), (Me.Top / yy1), _
     (Me.Width / xx1), (Me.Height / yy1), _
      conSwpNoActivate Or conSwpShowWindow
    Else
     Startup.SetWindowPos hwnd, conHwndNoTopmost, _
     (Me.Left / xx1), (Me.Top / yy1), _
     (Me.Width / xx1), (Me.Height / yy1), _
     conSwpNoActivate Or conSwpShowWindow
     End If
'Call MDIForm_Resize
End Sub
Private Sub e__Click(Index As Integer)
Select Case (Index)
Case 2
clu.text = Clipboard.GetText
Case 1
Clipboard.Clear
Clipboard.SetText clu.text, &HBF01
Case 5
Clipboard.Clear
Clipboard.SetText clu.fontname_, 1
Case 6
Clipboard.Clear
Clipboard.SetText clu.utext
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long
Select Case LCase(Button.Key)
Case "backward"
Toolbar1.Buttons("stop").Value = tbrPressed
Timer1.enabled = False
i = ListView1(0).SelectedItem.Index
If i = 1 Then i = (ListView1(0).ListItems.Count + 1)
clu.fontname_ = ListView1(0).ListItems(i - 1).text
ListView1(0).ListItems(i - 1).Selected = True
Case "forward"
Toolbar1.Buttons("stop").Value = tbrPressed
Timer1.enabled = False
i = ListView1(0).SelectedItem.Index
If i = (ListView1(0).ListItems.Count) Then i = 0
clu.fontname_ = ListView1(0).ListItems(i + 1).text
ListView1(0).ListItems(i + 1).Selected = True
Case "refresh"
clu.refresh_
Case "new"
Call f__Click(0)
Case "close"
Unload clu
Case "copy"
Call e__Click(1)
Case "paste"
Call e__Click(2)
Case "copyfont"
Call e__Click(5)
Case "copyanscii"
Call e__Click(6)
Case "clone"
If do_copy = True Then Button.Value = tbrUnpressed: Exit Sub
If Button.Value = tbrUnpressed Then Me.MousePointer = 0: do_cloning = False
If Button.Value = tbrPressed Then Me.MousePointer = 99: Me.MouseIcon = ImageList1.ListImages("clone").Picture:  do_cloning = True
Case "stylecopy"
If do_cloning = True Then Button.Value = tbrUnpressed: Exit Sub
If Button.Value = tbrPressed Then Me.MousePointer = 99: Me.MouseIcon = ImageList1.ListImages("paint").ExtractIcon: do_copy = True
If Button.Value = tbrUnpressed Then Me.MousePointer = 0: do_copy = False
Case "reset"
Call reset(clu)
If clu.fontbold_ = True Then Chk_style(0).Value = 1 Else Chk_style(0).Value = 0
If clu.fontitalic_ = True Then Chk_style(1).Value = 1 Else Chk_style(1).Value = 0
If clu.fontunderline_ = True Then Chk_style(2).Value = 1 Else Chk_style(2).Value = 0
If clu.fontstrikethru_ = True Then Chk_style(3).Value = 1 Else Chk_style(3).Value = 0
Case "play"
mnu_fs.Checked = True
Timer1.enabled = True
Case "stop"
mnu_fs.Checked = False
Timer1.enabled = False
Case "clock"
frmsettings.Move Me.Left + Button.Left, Me.Top + Button.Top + Button.Height
frmsettings.Show
End Select
End Sub

Private Sub w__Click(Index As Integer)
Select Case Index
Case 0
Me.Arrange 1
Case 1
Me.Arrange 2
Case 2
Me.Arrange 0
Case 3
Me.Arrange 3
End Select
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As frmshow, ans As Long
'ans = MsgBox(" Exit Fontcosine ?", vbYesNo + vbQuestion)
'If ans = 7 Then Cancel = True: Exit Sub
For Each i In class_col
Unload i
Next
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
SaveSetting "FontCosine", "mainview", "Height", Str(Me.Height)
SaveSetting "FontCosine", "mainview", "Top", Str(Me.Top)
SaveSetting "FontCosine", "mainview", "Width", Str(Me.Width)
SaveSetting "FontCosine", "mainview", "Left", Str(Me.Left)
SaveSetting "FontCosine", "mainview", "State", Str(Me.WindowState)
SaveSetting "FontCosine", "mainview", "fontbox", Abs(CInt(v_box(0).Checked))
SaveSetting "FontCosine", "mainview", "colorbox", Abs(CInt(v_box(1).Checked))
SaveSetting "FontCosine", "mainview", "toolbar", Abs(CInt(v_box(2).Checked))
SaveSetting "FontCosine", "mainview", "fontboxwidth", Str(Picture1.Width)
SaveSetting "FontCosine", "mainview", "fontshow", Abs(CInt(mnu_fs.Checked))
SaveSetting "FontCosine", "mainview", "ontop", Abs(CInt(f_vaot.Checked))
'*******
'Unload mdifrmmain
Unload frmSplash
'Unload frmshow
Unload frmOptions
Unload frmAbout
'Set Startup.myappa = Nothing
'*******
Unload frmsettings
Set Startup.myappa = Nothing
End Sub
Private Sub Timer2_Timer()
reset class_col(1)
Timer2.enabled = False
End Sub




