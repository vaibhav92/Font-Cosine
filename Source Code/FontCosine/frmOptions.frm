VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4935
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6225
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picoptions 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   2
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   5775
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "&Update"
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   45
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Remove"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   41
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   40
         Top             =   3120
         Width           =   855
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   5221
         TabIndex        =   39
         Top             =   2400
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   65
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text2"
         BuddyDispid     =   196612
         OrigLeft        =   4440
         OrigTop         =   960
         OrigRight       =   4635
         OrigBottom      =   1335
         Max             =   256
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3480
         TabIndex        =   38
         Top             =   2400
         Width           =   1740
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":000C
         Height          =   1875
         Index           =   2
         Left            =   2880
         TabIndex        =   44
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charec&ter"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   43
         Top             =   2400
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Charecters"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.PictureBox picoptions 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   5775
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   46
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Remove"
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   32
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   31
         Top             =   3120
         Width           =   855
      End
      Begin VB.ListBox Lstlaunch 
         Height          =   450
         Index           =   1
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton Cmd_comm 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   2520
         Width           =   375
      End
      Begin VB.ListBox Lstlaunch 
         Height          =   3180
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "&Launch Pad"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "&Application Title"
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Application &Path"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   $"frmOptions.frx":0109
         Height          =   1335
         Left            =   3120
         TabIndex        =   8
         Top             =   120
         Width           =   1950
      End
   End
   Begin VB.PictureBox picoptions 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   480
      Width           =   5775
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   3240
         ScaleHeight     =   3195
         ScaleWidth      =   2475
         TabIndex        =   28
         Top             =   240
         Width           =   2535
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            Height          =   195
            Left            =   1200
            TabIndex        =   29
            Top             =   1560
            Width           =   105
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Font"
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   2775
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   2400
            TabIndex        =   35
            Top             =   195
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            _Version        =   393216
            Value           =   115
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtcha"
            BuddyDispid     =   196626
            OrigLeft        =   2400
            OrigTop         =   240
            OrigRight       =   2595
            OrigBottom      =   495
            Max             =   256
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtcha 
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Text            =   "S"
            Top             =   195
            Width           =   1200
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   2400
            TabIndex        =   27
            Top             =   960
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            _Version        =   393216
            Value           =   12
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Txtfontsize"
            BuddyDispid     =   196628
            OrigLeft        =   2520
            OrigTop         =   720
            OrigRight       =   2715
            OrigBottom      =   975
            Max             =   500
            Min             =   8
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Txtfontsize 
            Height          =   285
            Left            =   1200
            TabIndex        =   26
            Text            =   "10"
            Top             =   960
            Width           =   1200
         End
         Begin VB.ComboBox Cmb_fonts 
            Height          =   315
            Index           =   1
            ItemData        =   "frmOptions.frx":0195
            Left            =   1200
            List            =   "frmOptions.frx":01A5
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1320
            Width           =   1380
         End
         Begin VB.ComboBox Cmb_fonts 
            Height          =   315
            Index           =   0
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   600
            Width           =   1380
         End
         Begin VB.CheckBox Chk_fonts 
            Caption         =   "Strike"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   23
            Top             =   1680
            Width           =   855
         End
         Begin VB.CheckBox Chk_fonts 
            Caption         =   "Underline"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Lblfont 
            AutoSize        =   -1  'True
            Caption         =   "Charecter"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Lblfont 
            AutoSize        =   -1  'True
            Caption         =   "Font Style"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label Lblfont 
            AutoSize        =   -1  'True
            Caption         =   "Font Size"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Lblfont 
            AutoSize        =   -1  'True
            Caption         =   "Font Name"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Color"
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2775
         Begin VB.Label lblcolorp 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   17
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblcolorp 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Lblcolor 
            AutoSize        =   -1  'True
            Caption         =   "BackColor"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Lblcolor 
            AutoSize        =   -1  'True
            Caption         =   "ForeColor"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   675
         End
      End
   End
   Begin MSComDlg.CommonDialog comm1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4440
      WhatsThisHelpID =   2
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsoptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "Group1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Launch Pad"
            Key             =   "Group2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "My Chars"
            Key             =   "Group3"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
Private Sub Chk_fonts_Click(Index As Integer)
Label4.fontunderline = Abs(CBool(Chk_fonts(0).Value))
Label4.fontstrikethru = Abs(CBool(Chk_fonts(1).Value))
End Sub
Private Sub Cmb_fonts_Change(Index As Integer)
If Index = 0 Then
Label4.fontname = Cmb_fonts(Index).text
Else
Select Case (Cmb_fonts(Index).ListIndex)
Case 0
Label4.fontbold = False
Label4.fontitalic = False
Case 1
Label4.fontbold = False
Label4.fontitalic = True
Case 2
Label4.fontitalic = False
Label4.fontbold = True
Case 3
Label4.fontbold = True
Label4.fontitalic = True
End Select
End If
End Sub
Private Sub Cmb_fonts_Click(Index As Integer)
Call Cmb_fonts_Change(Index)
End Sub
Private Sub Cmd_comm_Click()
comm1.Filter = "Applications |*.exe; *.com ; *.bat"
comm1.ShowOpen
Text1(1).text = comm1.filename
End Sub
Private Sub cmdApply_Click()
On Error GoTo errhand
Dim i As Long
SetKeyValue "TestKey\SubKey1", "StringValue", "Hello", REG_SZ
SaveSetting App.ProductName, "Options", "forecolor", Str(lblcolorp(0).backcolor)
SaveSetting App.ProductName, "Options", "backcolor", Str(lblcolorp(1).backcolor)
SaveSetting App.ProductName, "Options", "fontname", Cmb_fonts(0).text
SaveSetting App.ProductName, "Options", "fontsize", Val(Txtfontsize.text)
SaveSetting App.ProductName, "Options", "fontstyle", Cmb_fonts(1).ListIndex
SaveSetting App.ProductName, "Options", "fontunderline", Chk_fonts(0).Value
SaveSetting App.ProductName, "Options", "fontstrike", Chk_fonts(1).Value
SaveSetting App.ProductName, "Options", "Char", Label4.Caption
DeleteSetting App.ProductName, "LaunchPad"
For i = 0 To (Lstlaunch(0).ListCount - 1)
SaveSetting App.ProductName, "Launchpad", Lstlaunch(0).List(i), Lstlaunch(1).List(i)
Next i
DeleteSetting App.ProductName, "mychars"
For i = 0 To (List1.ListCount - 1)
SaveSetting App.ProductName, "mychars", List1.List(i), " "
Next i
errhand:
If Err.Number = 5 Then Err.Clear: Resume Next
If Err.Number <> 0 Then MsgBox "Error ": Stop
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call cmdApply_Click
    Unload Me
End Sub
Private Sub Command1_Click(Index As Integer)
'On Error Resume Next
If Index = 0 Then
If Trim(Text1(0).text) = "" Or Trim(Text1(1).text) = "" Then MsgBox "Invalid Title or Path": Exit Sub
Lstlaunch(0).AddItem Text1(0).text
Lstlaunch(1).AddItem Text1(1).text
If Lstlaunch(0).ListCount > 0 Then Command1(1).enabled = True
End If
If Index = 1 Then
Dim i As Long
i = Lstlaunch(0).ListIndex
If i >= 0 Then
Lstlaunch(0).RemoveItem (i)
Lstlaunch(1).RemoveItem (i)
End If
If Lstlaunch(0).ListCount <= 0 Then Command1(1).enabled = False
End If
If Index = 2 Then
If Trim(Text1(0).text) = "" Or Trim(Text1(1).text) = "" Then MsgBox "Invalid Title or Path": Exit Sub
Lstlaunch(0).List(Lstlaunch(0).ListIndex) = Text1(0).text
Lstlaunch(1).List(Lstlaunch(0).ListIndex) = Text1(1).text
'Lstlaunch(1).text = Text1(2).text
End If
End Sub
Private Sub Command2_Click(Index As Integer)
If Index = 0 Then List1.AddItem Text2.text
If Index = 1 Then
List1.RemoveItem List1.ListIndex
If List1.ListCount <= 0 Then Command2(1).enabled = False Else Command2(1).enabled = True
Call List1_Click
End If
If Index = 2 Then
List1.List(List1.ListIndex) = Text2.text
'List1.text = Text2.text
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsoptions.SelectedItem.Index
        If i = tbsoptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsoptions.SelectedItem = tbsoptions.Tabs(1)
        Else
            'increment the tab
            Set tbsoptions.SelectedItem = tbsoptions.Tabs(i + 1)
        End If
    End If
End Sub
Private Sub Form_Load()
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
'If kl(0) = "Not Found" Then Label4.fontname = "Times New Roman" _
'Else: Label4.fontname = kl(0)
'If kl(1) = "Not Found" Then Label4.fontsize = 30 _
'Else Label4.fontsize = Val(kl(1))
'If kl(2) = "Not Found" Then Label4.forecolor = CLng("0") _
'Else Label4.forecolor = CLng(kl(2))
'If kl(3) = "Not Found" Then Picture1.backcolor = CLng(" 16777215") _
'Else Picture1.backcolor = CLng(kl(3))
'If kl(4) = "Not Found" Then Label4.fontstrikethru = CBool("0") _
'Else: Label4.fontstrikethru = CBool(kl(4))
'If kl(5) = "Not Found" Then Label4.fontunderline = CBool("0") _
'Else Label4.fontunderline = CBool(kl(5))
'If kl(6) = "Not Found" Then Cmb_fonts(1).ListIndex = Val("0") _
'Else Cmb_fonts(1).ListIndex = Val(kl(6))
'If kl(7) = "Not Found" Then Label4.Caption = "S" _
'Else Label4.Caption = kl(7)
'Dim kkl As Long
'setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Launchpad", kkl)
'Do Until setta(2) = "Not Found"
'    Lstlaunch(0).AddItem setta(1)
'    Lstlaunch(1).AddItem setta(1)
'    kkl = kkl + 1
'    setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Launchpad", kkl)
'Loop
'kkl = 0
'setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\Mychars", kkl)
'Do Until setta(2) = "Not Found"
'    List1.AddItem setta(2)
'    kkl = kkl + 1
'    setta = EllReg.ReadRegistryGetAll(EllReg.HKEY_CURRENT_USER, "Software\Intellective Solutions\FontCosine\mychars", kkl)
'Loop
'***********
Dim xx1 As Single, yy1 As Single
On Error GoTo errhand
Dim i As Integer, lp As Variant, ero As Boolean
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
For i = 1 To Screen.FontCount
Cmb_fonts(0).AddItem Screen.Fonts(i)
Next i
Cmb_fonts(0).ListIndex = Val(GetSetting(App.ProductName, "options", _
"font", "0"))
Txtfontsize.text = GetSetting(App.ProductName, "options", _
"fontsize", "10")
Cmb_fonts(1).ListIndex = Val(GetSetting(App.ProductName, "options", _
"fontstyle", "0"))
Chk_fonts(0).Value = GetSetting(App.ProductName, "options", _
"fontunderline", "0")
Chk_fonts(1).Value = GetSetting(App.ProductName, "options", _
"fontstrike", "0")
lblcolorp(1).backcolor = CLng(GetSetting(App.ProductName, "options", _
"backcolor", "16777215"))
lblcolorp(0).backcolor = CLng(GetSetting(App.ProductName, "options", _
"forecolor", "0"))
Picture1.backcolor = lblcolorp(1).backcolor
Label4.forecolor = lblcolorp(0).backcolor
Label4.Caption = GetSetting(App.ProductName, "options", _
"Char", "S")
lp = GetAllSettings(App.ProductName, "LaunchPad")
For i = 0 To UBound(lp, 1)
        If ero Then Err.Clear: Exit For
        Lstlaunch(0).AddItem lp(i, 0)
        Lstlaunch(1).AddItem lp(i, 1)
Next
ero = False
lp = GetAllSettings(App.ProductName, "mychars")
For i = 0 To UBound(lp, 1)
        If ero Then Err.Clear: Exit For
        List1.AddItem lp(i, 0)
Next
errhand:
If Err.Number = 13 Then
'Err.Clear
ero = True
Resume Next
'Else
'MsgBox "error"
'End
End If
End Sub
Private Sub lblcolorp_Click(Index As Integer)
comm1.Flags = cdlCCRGBInit
comm1.Color = lblcolorp(Index).backcolor
comm1.ShowColor
lblcolorp(Index).backcolor = comm1.Color
If Index = 1 Then Picture1.backcolor = comm1.Color Else Label4.forecolor = comm1.Color
End Sub
Private Sub List1_Click()
Text2.text = List1.text
Text2.SelStart = 0
Text2.SelLength = Len(Text2.text)
End Sub
Private Sub Lstlaunch_Click(Index As Integer)
Text1(0).text = Lstlaunch(0).text
Text1(1).text = Lstlaunch(1).List(Lstlaunch(0).ListIndex)
End Sub
Private Sub tbsOptions_Click()
        Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsoptions.Tabs.Count - 1
        If i = tbsoptions.SelectedItem.Index - 1 Then
            picoptions(i).Left = tbsoptions.ClientLeft
            picoptions(i).Top = tbsoptions.ClientTop
            picoptions(i).Width = tbsoptions.ClientWidth
            picoptions(i).Width = tbsoptions.ClientWidth
            picoptions(i).enabled = True
            
            
            picoptions(i).Visible = True
      
           
         
         Else
            picoptions(i).Left = -20000
            picoptions(i).enabled = False
            picoptions(i).Visible = False
        End If
    Next
End Sub
Private Sub txtcha_Change()
Label4.Caption = Left(txtcha.text, 1)
txtcha.SelStart = 0
txtcha.SelLength = Len(txtcha.text)
End Sub
Private Sub Txtfontsize_Change()
If Val(Txtfontsize.text) >= UpDown1.Min And Val(Txtfontsize.text) <= UpDown1.Max Then UpDown1.Value = Str(Txtfontsize.text)
End Sub
Private Sub UpDown1_Change()
Label4.fontsize = UpDown1.Value
Label4.Top = (Picture1.ScaleHeight - Label4.Height) / 2
Label4.Left = (Picture1.ScaleWidth - Label4.Width) / 2
End Sub
Private Sub UpDown2_Change()
txtcha.text = Chr(UpDown2.Value)
End Sub
Private Sub UpDown3_Change()
Text2.text = Chr(UpDown3.Value)
End Sub
