VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF14BA5-51F1-11D3-87BC-38D00CC17206}#4.0#0"; "TREEFOLDER.OCX"
Begin VB.UserControl fontview 
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ScaleHeight     =   5355
   ScaleWidth      =   2325
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   2
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
      Begin MSComctlLib.ListView ListView1 
         Height          =   465
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Tag             =   "My Views"
         Top             =   0
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   820
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
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
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Index           =   1
      Left            =   360
      ScaleHeight     =   2355
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
      Begin VB.FileListBox File1 
         Height          =   480
         Left            =   240
         Pattern         =   "*.ttf"
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   465
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Tag             =   "Uninstalled"
         Top             =   600
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   820
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
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
      Begin MyTreeFolder.Treefolder Treefolder1 
         Height          =   30
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "font_view2.ctx":0000
         LabelEdit       =   1
         Indentation     =   566.929
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4335
      Index           =   0
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   840
      Width           =   2055
      Begin MSComctlLib.ListView ListView1 
         Height          =   105
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Tag             =   "Screen Fonts"
         Top             =   240
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   185
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
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
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Fonts"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   240
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   1125
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5355
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   9446
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Installed"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Uninstalled"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "My Views"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711680
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "font_view2.ctx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "font_view2.ctx":27D0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fontview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pre As Integer
Dim fon As String
Event Click()
Private Sub ListView1_BeforeLabelEdit(index As Integer, Cancel As Integer)
Cancel = True
End Sub
Private Sub ListView1_ItemClick(index As Integer, ByVal Item As MSComctlLib.ListItem)
fon = ListView1(index).SelectedItem.Text
RaiseEvent Click
End Sub
Private Sub TabStrip1_Click()
Static pre As Integer, sele As String
sele = TabStrip1.SelectedItem.index
sele = sele - 1
If sele = pre Then Exit Sub
Picture2(pre).Visible = False
Picture2(sele).Visible = True
'ListView1(sele).ListItems.Selected = True
Treefolder1.Visible = True
Label1.Caption = ListView1(sele).Tag
pre = sele
End Sub
Private Sub usercontrol_initialize()
Dim i As Long
For i = 0 To Screen.FontCount - 1
ListView1(0).ListItems.Add , , Screen.Fonts(i), 1, 1
Next
End Sub
Public Property Get Font() As String
Font = fon
End Property
Private Sub UserControl_Resize()
Dim i As Long
TabStrip1.Height = UserControl.ScaleHeight
TabStrip1.Width = UserControl.ScaleWidth
Picture1.Left = TabStrip1.ClientLeft
Picture1.Width = TabStrip1.ClientWidth
Picture1.Left = TabStrip1.ClientLeft
Picture1.Top = TabStrip1.ClientTop
For i = 0 To (Picture2.Count - 1)
Picture2(i).Width = TabStrip1.ClientWidth
Picture2(i).Left = TabStrip1.ClientLeft
Picture2(i).Top = Picture1.Top + Picture1.ScaleHeight + 20
Picture2(i).Height = TabStrip1.Height - Picture2(i).Top - 40
Next i
End Sub
