VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl fontview 
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ScaleHeight     =   5355
   ScaleWidth      =   2715
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   90
      ScaleHeight     =   315
      ScaleWidth      =   2535
      TabIndex        =   3
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
         TabIndex        =   4
         Top             =   30
         Width           =   1125
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4485
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Tag             =   "Both"
      Top             =   780
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   7911
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4425
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Tag             =   "Printer Fonts"
      Top             =   780
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   7805
      LabelWrap       =   -1  'True
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
         Text            =   "Font_Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4425
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "Screen Fonts"
      Top             =   780
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   7805
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5355
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   9446
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Screen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Printer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Both"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1500
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "font_view.ctx":0000
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
Dim pre As Integer
Dim fon As String
Event click()
Private Sub ListView1_BeforeLabelEdit(Index As Integer, Cancel As Integer)
Cancel = True
End Sub
Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
fon = ListView1(Index).SelectedItem.Text
RaiseEvent click
End Sub
Private Sub TabStrip1_Click()
'Static pre As Integer
sele = TabStrip1.SelectedItem.Index
sele = sele - 1
If sele = pre Then Exit Sub
ListView1(pre).Visible = False
ListView1(sele).Visible = True
ListView1(sele).ListItems(1).Selected = True
'fon=
Label1.Caption = ListView1(sele).Tag
pre = sel
End Sub
Private Sub UserControl_Initialize()
For i = 0 To Screen.FontCount - 1
ListView1(0).ListItems.Add , , Screen.Fonts(i), 1, 1
ListView1(2).ListItems.Add , , Screen.Fonts(i), 1, 1
Next
For i = 0 To Printer.FontCount - 1
ListView1(1).ListItems.Add , , Printer.Fonts(i), 1, 1
ListView1(2).ListItems.Add , , Printer.Fonts(i), 1, 1
Next
End Sub
Public Property Get font() As String
font = fon
End Property
Private Sub UserControl_Resize()
TabStrip1.Height = UserControl.Height
TabStrip1.Width = UserControl.Width
Picture1.Width = UserControl.Width - 270
ListView1(0).Width = Picture1.Width
ListView1(1).Width = Picture1.Width
ListView1(2).Width = Picture1.Width
Picture1.Left = 150
ListView1(0).Left = 150
ListView1(1).Left = 150
ListView1(2).Left = 150
'Command1.Left = Picture1.Width - Command1.Width
ListView1(0).Height = TabStrip1.Height * 0.842231559290383
ListView1(1).Height = TabStrip1.Height * 0.842231559290383
ListView1(2).Height = TabStrip1.Height * 0.842231559290383
End Sub
