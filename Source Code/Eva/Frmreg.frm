VERSION 5.00
Begin VB.Form Frmreg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FontCosine Registration"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5265
   Icon            =   "Frmreg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Registration Data"
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   405
         Index           =   1
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Unregistered"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Index           =   2
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "Evaluation"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Index           =   0
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "Unregistered"
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "&Organization:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Registeration &Code:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Username:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Please fill in the information below."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "Frmreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event validate(name As String, com As String, code As String)
Public fila As Long
Private Sub Command1_Click()
Dim stri As String, ans As Long
If Len(Text1(0).Text) < 4 Or Len(Text1(1).Text) < 4 Then MsgBox "User/Organization Names should be atleast 4 Charecters Long", vbExclamation, "Registration": Exit Sub
stri = Chr(13) + "      Name:                 " & Text1(0).Text + Chr(13)
stri = stri & "     Company/Organization: " & Text1(1).Text + Chr(13)
stri = stri & "     Registration Code:    " & Text1(2).Text + Chr(13) + Chr(13)
stri = stri & "Is this information correct?"
ans = MsgBox(stri, vbQuestion + vbYesNo, "FonCosine Registration")
If ans = 6 Then Call kl
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub kl()
Dim strin As String, sstrin As String
strin = "INVALID INFORMATION!!" & Chr(13) & "Please check you Name/Organization/Registration Code"
sstrin = "Registration Failure"
Text1(0).Text = Left(Text1(0).Text & String(30, " "), 30)
Text1(1).Text = Left(Text1(1).Text & String(30, " "), 30)
Text1(2).Text = Left(Text1(2).Text & String(30, " "), 30)
If Not (Trim(Serials.get_serial(Text1(0).Text)) = Trim(Text1(2).Text)) Then MsgBox strin, vbExclamation, sstrin
If (Trim(Serials.get_serial(Text1(0).Text)) = Trim(Text1(2).Text)) Then
Call filereset
RaiseEvent validate(Text1(0), Text1(1), Text1(2))
'Unload Me
End If
End Sub
Private Sub filereset()
Dim i As String
i = genkey()
Call save_key(i)
put_record fila, "lulsdasoO*&*)", 1 ' dummy
put_record fila, crypt("2SinCos", i).data, 2 ' dummy
put_record fila, "lul4234s&*)", 3 ' dummy
put_record fila, crypt(Str(Date), i).data, 4 ' datalastacc
put_record fila, "lulsdasdfoO*)", 5 ' dummy
put_record fila, crypt(Str(CLng(Timer)), i).data, 6 ' timelast ascce
put_record fila, "lulsdwasossd*)", 7 ' dummy
put_record fila, crypt("10", i).data, 8 ' sta
put_record fila, "luld2sdadsdoO*)", 9 ' dummy
' reg details
put_record fila, crypt(Left(Text1(0).Text, 15), i).data, 10
put_record fila, crypt(Right(Text1(0).Text, 15), i).data, 11
put_record fila, crypt(Left(Text1(1).Text, 15), i).data, 12
put_record fila, crypt(Right(Text1(1).Text, 15), i).data, 13
put_record fila, crypt(Left(Text1(2).Text, 15), i).data, 14
put_record fila, crypt(Right(Text1(2).Text, 15), i).data, 15
' date installed
put_record fila, crypt(Str(Date), i).data, 16
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
