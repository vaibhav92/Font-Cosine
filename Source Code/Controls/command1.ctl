VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1140
      Top             =   2430
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
            Picture         =   "command1.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "command1.ctx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1650
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   1170
      Width           =   1185
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   315
         Left            =   750
         Stretch         =   -1  'True
         Top             =   60
         Width           =   345
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   525
      Left            =   1620
      TabIndex        =   0
      Top             =   1140
      Width           =   1245
   End
   Begin VB.Menu dj 
      Caption         =   "dj"
      Begin VB.Menu s1 
         Caption         =   "1"
      End
      Begin VB.Menu s2 
         Caption         =   "2"
      End
      Begin VB.Menu s3 
         Caption         =   "3"
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim i As Menu
Private Sub Command1_Click()
If Image2.Picture = ImageList1.ListImages(1).Picture Then
Image2.Picture = ImageList1.ListImages(2).Picture
Else
Image2.Picture = ImageList1.ListImages(1).Picture
End If
End Sub
