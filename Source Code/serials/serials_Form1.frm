VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "300"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If Trim(get_serial(Text1.Text)) <> Trim(Text2.Text) Then
MsgBox "Invalid Serial"
Else
MsgBox "Valid Serial"
End If
End Sub

Private Sub Text1_Change()
Text2.Text = get_serial(Text1.Text)
Label1.Caption = Str(Len(Text2.Text))
End Sub
