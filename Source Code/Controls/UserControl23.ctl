VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ScaleHeight     =   4095
   ScaleWidth      =   3255
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      FillColor       =   &H80000004&
      ForeColor       =   &H8000000D&
      Height          =   135
      Left            =   690
      ScaleHeight     =   75
      ScaleWidth      =   2325
      TabIndex        =   1
      Top             =   2190
      Visible         =   0   'False
      Width           =   2385
   End
   Begin MSComctlLib.ListView list1 
      Height          =   2325
      Left            =   0
      TabIndex        =   3
      Top             =   1710
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   4101
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   630
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   630
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox Picture2 
      Height          =   105
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   1620
      Width           =   3225
   End
   Begin MSComctlLib.TreeView Treeview1 
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2778
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl23.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl23.ctx":27B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl23.ctx":4F68
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
Event treemousedown()
Event click()
Event mousedown()
Dim m_caption As String
Dim group_path As String
Dim text_  As String
Dim group_name As String
Private Sub list1_Click()
RaiseEvent click
End Sub
Private Sub list1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mousedown(Button, Shift, x, y)
End Sub
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
If Treeview1.SelectedItem.index = 1 Then
Cancel = True
End If
End Sub
Private Sub Treeview1_AfterLabelEdit(Cancel As Integer, NewString As String)
Name Treeview1.SelectedItem.Text As get_path2(Treeview1.SelectedItem.index) & NewString
End Sub
Private Sub UserControl_Initialize()
opening_path = "d:\ins"
file_pattern = "*.*"
m_caption = "Groups"
show_files = True
Dim i  As Variant
Set i = list1.ColumnHeaders.Add(, , "Name", list1.Width / 4)
Set i = list1.ColumnHeaders.Add(, , "File-Name", list1.Width / 4)
Set i = list1.ColumnHeaders.Add(, , "Location", list1.Width / 4)
Set i = list1.ColumnHeaders.Add(, , "Size", list1.Width / 4)
Set i = Treeview1.Nodes.Add(, , , m_caption, 1)
i.Tag = opening_path
list1.View = lvwReport
Dir1.Path = opening_path
File1.Path = opening_path
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent treemousedown(Button, Shift, x, y)
End Sub
Private Sub UserControl_Resize()
Treeview1.Top = 5
Treeview1.Width = UserControl.Width
list1.Width = UserControl.Width
Treeview1.Height = UserControl.Height / 2.5
list1.Height = UserControl.Height - Treeview1.Height - Picture2.Height
Picture2.Top = Treeview1.Top + Treeview1.Height
list1.Top = Picture2.Top + Picture2.Height + 10
Picture1.Width = UserControl.Width
Picture2.Width = UserControl.Width
End Sub
Private Sub get_subdir(Path As String)
If Treeview1.SelectedItem.Children > 0 Then Exit Sub
Dir1.Path = Path
Dim i As Variant
For j = 0 To Dir1.ListCount - 1
Set i = Treeview1.Nodes.Add(Treeview1.SelectedItem.index, 4, , getdir(Dir1.List(j)), 2)
Next
If show_files = True Then Call Get_files(Path)
End Sub
Private Sub Get_files(Path As String)
File1.Path = Path
list1.ListItems.Clear
Dim i As Variant
For j = 0 To (File1.ListCount - 1)
Set i = list1.ListItems.Add(, , "", 3, 3)
i.SubItems(1) = File1.List(j)
pata = get_path(Treeview1.SelectedItem.index)
i.SubItems(2) = pata
pata = pata & i.SubItems(1)
i.SubItems(3) = Str(Int(FileLen(pata) / 1024)) & " KB"
Next
End Sub
Private Sub TreeView1_Click()
If Treeview1.SelectedItem.Image <> 3 Then
Call get_subdir(get_path(Treeview1.SelectedItem.index))
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
pa = Treeview1.Nodes(i).Text & "\" & pa
i = Treeview1.Nodes(i).Parent.index
Wend
pa = Treeview1.Nodes(1).Tag & "\" & pa
get_path = pa
End Function
Private Function get_path2(index As Integer) As String
Dim pa As String
i = Treeview1.Nodes(index).Parent.index
If i = 1 Then
get_path2 = Treeview1.Nodes(1).Tag & "\"
Exit Function
End If
While i <> 1
pa = Treeview1.Nodes(i).Text & "\" & pa
i = Treeview1.Nodes(i).Parent.index
Wend
pa = Treeview1.Nodes(1).Tag & "\" & pa
get_path2 = pa
End Function

