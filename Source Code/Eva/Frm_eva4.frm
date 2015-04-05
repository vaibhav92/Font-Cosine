VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FontCosine Evaluation"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7245
   Icon            =   "Frm_eva4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   5160
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Evaluation Period Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enter Registration Code"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Continue Unregisterd"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4155
      Left            =   120
      Picture         =   "Frm_eva4.frx":27A2
      ScaleHeight     =   4095
      ScaleWidth      =   2610
      TabIndex        =   2
      Top             =   120
      Width           =   2670
      Begin VB.Image Image1 
         Height          =   4140
         Left            =   0
         Picture         =   "Frm_eva4.frx":46DB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2625
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   3000
      X2              =   3000
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   3000
      X2              =   4560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   3000
      X2              =   4560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "This Program is a SHAREWARE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2730
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer2_Timer()
Label1.Enabled = Not (Label1.Enabled)
End Sub
