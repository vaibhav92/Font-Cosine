VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsettings 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   465
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   50
      SmallChange     =   10
      Min             =   1
      Max             =   999
      SelStart        =   500
      TickStyle       =   3
      Value           =   500
      TextPosition    =   1
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
With Slider1
.Top = 0
.Height = Me.Height
.Left = 0
.Width = Me.Width
End With
End Sub
Private Sub Slider1_LostFocus()
Me.Hide
End Sub
Private Sub Slider1_Scroll()
mdifrmmain.Timer1.Interval = 1000 - Slider1.Value
End Sub
