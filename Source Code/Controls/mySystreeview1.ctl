VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ScaleHeight     =   6645
   ScaleWidth      =   6945
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5745
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   10134
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2310
      TabIndex        =   3
      Top             =   2040
      Width           =   105
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   2040
      Width           =   135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2070
      TabIndex        =   1
      Top             =   2040
      Width           =   270
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2430
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview1.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview1.ctx":27B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview1.ctx":4F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview1.ctx":771C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'On Error GoTo errhand
Private Sub Usercontrol_Load()
ChDir "c:\"
TreeView1.ImageList = ImageList1
Dim i As Node
Set i = TreeView1.Nodes.Add(, , "r", "My Computer", 1)
For j = 1 To (Drive1.ListCount - 1)
Set i = TreeView1.Nodes.Add("r", 4, , Left(Drive1.List(j), 2), 2)
Next j
End Sub
Private Sub get_subdir(Path As String)
If TreeView1.SelectedItem.Tag = "1" Then Exit Sub
Dir1.Path = Path
Dim i As Node
For j = 0 To Dir1.ListCount - 1
Set i = TreeView1.Nodes.Add(TreeView1.SelectedItem.index, 4, , getdir(Dir1.List(j)), 3)
Next
Call Get_files(Path)
TreeView1.SelectedItem.Tag = "1"
End Sub
Private Sub Get_files(Path As String)
File1.Path = Path
Dim i As Node
For j = 0 To (File1.ListCount - 1)
Set i = TreeView1.Nodes.Add(TreeView1.SelectedItem.index, 4, , File1.List(j), 4)
Next
End Sub
Private Sub TreeView1_Click()
If TreeView1.SelectedItem.index > 1 Then
If TreeView1.SelectedItem.Image <> 4 Then
Call get_subdir(get_path(TreeView1.SelectedItem.index))
End If
End If
End Sub
Private Function getdir(i As String) As String
For j = Len(i) To 1 Step -1
If Mid(i, j, 1) = "\" Then Exit For
Next
j = Len(i) - j
getdir = Right$(i, j)
End Function
Private Function get_path(index As Integer) As String
Dim pa As String
i = index
While i <> 1
pa = TreeView1.Nodes(i).Text & "\" & pa
i = TreeView1.Nodes(i).Parent.index
Wend
get_path = pa
'get_path = "\" & pa
End Function

': errhand

'if err.Number




