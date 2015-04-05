VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Group_view 
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ScaleHeight     =   4530
   ScaleWidth      =   2970
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   2925
      TabIndex        =   7
      Top             =   30
      Width           =   2955
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2730
         TabIndex        =   8
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Groups"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView list1 
      Height          =   2325
      Left            =   0
      TabIndex        =   3
      Top             =   2190
      Width           =   2955
      _ExtentX        =   5212
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   2925
      TabIndex        =   4
      Top             =   1950
      Width           =   2955
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2730
         TabIndex        =   5
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fonts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      FillColor       =   &H80000004&
      ForeColor       =   &H8000000D&
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   4290
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.PictureBox Picture2 
      Height          =   75
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   15
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   1860
      Width           =   2955
   End
   Begin MSComctlLib.TreeView Treeview1 
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   270
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2778
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
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
            Picture         =   "UserControl233.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl233.ctx":27B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl233.ctx":4F68
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Group_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event treemousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event groupchanged(name As String)
Private Type font_data
name As String
file As String
filesize As String
font_file_location As String
group As String
End Type
Dim group_folder As String
Dim text_  As String
Dim font_dat As font_data
Private Sub list1_AfterLabelEdit(Cancel As Integer, NewString As String)
For i = 1 To list1.ListItems.Count - 1
If LCase(list1.ListItems(i).Text) = LCase(NewString) Then
MsgBox "Group " & NewString & " Already Exists", , "Err-2 Rename Member Error"
Cancel = True
Exit Sub
End If
Next
res = EllReg.ReadRegistry(HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & Treeview1.SelectedItem.Text & "\", list1.SelectedItem.Text)
SetKeyValue "Software\fonts\control1\groups\" & Treeview1.SelectedItem.Text & "\", NewString, res, REG_SZ
DeleteValue HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & Treeview1.SelectedItem.Text & "\", list1.SelectedItem.Text
Call get_values
End Sub

Private Sub list1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call get_values
RaiseEvent Click
End Sub
Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then Call delete_member(list1.SelectedItem.Text, Treeview1.SelectedItem.Text)
End Sub
Private Sub list1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Move Picture2.Left, Picture2.Top
Picture1.Visible = True
vbmoving = True
End Sub
Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Visible = False
vbmoving = False
End Sub
Private Sub Treeview1_AfterLabelEdit(Cancel As Integer, NewString As String)
For i = 2 To Treeview1.Nodes.Count
If LCase(Treeview1.Nodes(i).Text) = LCase(NewString) Then
MsgBox "Group " & NewString & " Already Exists", , "Err-1 Rename Group Error"
Cancel = True
Exit Sub
End If
Next
EllReg.CreateNewKey "SOFTWARE\fonts\control1\Groups\" & NewString, HKEY_LOCAL_MACHINE
Dim res As Variant, valu1 As String, valu2 As String
Dim j As Long
j = 0
res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & Treeview1.SelectedItem.Text & "\", j)
Do Until res(2) = "Not Found"
   valu1 = res(1)
   valu2 = res(2)
   SetKeyValue "SOFTWARE\fonts\control1\Groups\" & NewString, valu1, valu2, REG_SZ
   j = j + 1
   res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & Treeview1.SelectedItem.Text & "\", j)
Loop
EllReg.DeleteSubkey HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & Treeview1.SelectedItem.Text & "\"
RaiseEvent groupchanged(Treeview1.SelectedItem.Text)
Call get_values
End Sub
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
If Treeview1.SelectedItem.index = 1 Then
Cancel = True
End If
End Sub
Private Sub Treeview1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then Call delete_group(Treeview1.SelectedItem.Text)
End Sub
Function delete_member(memname As String, grname As String) As Boolean
On Error GoTo errhand
For i = 2 To Treeview1.Nodes.Count + 1
If LCase(Treeview1.Nodes(i).Text) = LCase(grname) Then Exit For
Next
For jji = 1 To list1.ListItems.Count + 1
If LCase(list1.ListItems(jji).Text) = LCase(memname) Then Exit For
Next
key_name = Treeview1.Nodes(i).Text
mem_name = list1.ListItems(jji).Text
j = MsgBox("Are You Sure To Delete Group " & key_name, vbQuestion + vbOKCancel)
If j = 1 Then
EllReg.DeleteValue HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & key_name, val_name
list1.ListItems.Remove (jji)
delete_member = True
Else
delete_member = False
End If
Call get_values
Exit Function
errhand:
If Err.Number = 35600 Then
delete_member = False
MsgBox "Invalid Group Name to Delete", , "Error-3"
Err.Clear
Exit Function
End If
End Function
Function delete_group(grname As String) As Boolean
On Error GoTo errhand
For i = 2 To Treeview1.Nodes.Count + 1
If LCase(Treeview1.Nodes(i).Text) = LCase(grname) Then Exit For
Next
key_name = Treeview1.Nodes(i).Text
j = MsgBox("Are You Sure To Delete Group " & key_name, vbQuestion + vbOKCancel)
If j = 1 Then
DeleteSubkey HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & key_name
list1.ListItems.Clear
Treeview1.Nodes.Remove (i)
delete_group = True
Else
delete_group = False
End If
Call get_values
Exit Function
errhand:
If Err.Number = 35600 Then
delete_group = False
MsgBox "Invalid Group Name to Delete", , "Error-3"
Err.Clear
Exit Function
End If
End Function
Private Sub Treeview1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent treemousedown(Button, Shift, X, Y)
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errhand
list1.ListItems.Clear
If Treeview1.SelectedItem.index = 1 Then Exit Sub
Dim i As Variant, valid As Boolean
Dim res As Variant
Dim j As Long
j = 0
valid = True
res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & Treeview1.SelectedItem.Text & "\", j)
Do Until res(2) = "Not Found"
   file_size1 = FileLen(group_folder & "\" & res(2))
   If valid = True Then
   Set i = list1.ListItems.Add(, , res(1), 3, 3)
   i.SubItems(1) = res(2)
   i.SubItems(2) = group_folder
   i.SubItems(3) = file_size1 & " KB"
      End If
   valid = True
   j = j + 1
   res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & Treeview1.SelectedItem.Text & "\", j)
Loop
If list1.ListItems.Count >= 1 Then list1.ListItems(1).Selected = True
Call get_values
errhand:
 Select Case Err.Number
 Case 53
    valid = False
    Resume Next
 Case 76
    valid = False
    Resume Next
 End Select
End Sub
Private Sub get_values()
If Treeview1.SelectedItem.index > 1 Then
font_dat.group = Treeview1.SelectedItem.Text
Else
font_dat.group = ""
End If
If list1.ListItems.Count > 0 Then
font_dat.name = list1.SelectedItem.Text
font_dat.file = list1.SelectedItem.SubItems(1)
font_dat.font_file_location = list1.SelectedItem.SubItems(2)
font_dat.filesize = list1.SelectedItem.SubItems(3)
Else
font_dat.name = ""
font_dat.file = ""
font_dat.font_file_location = ""
font_dat.filesize = ""
End If
End Sub
Private Sub UserControl_Initialize()
Call refresh
End Sub
Sub refresh()
list1.ListItems.Clear
Treeview1.Nodes.Clear
Dim i  As Variant
Set i = list1.ColumnHeaders.Add(, , "Name", list1.width / 4)
Set i = list1.ColumnHeaders.Add(, , "File-Name", list1.width / 4)
Set i = list1.ColumnHeaders.Add(, , "Location", list1.width / 4)
Set i = list1.ColumnHeaders.Add(, , "Size", list1.width / 4)
list1.View = lvwReport
group_folder = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\", "Group_folder")
Set i = Treeview1.Nodes.Add(, , , "Groups", 1, 1)
Dim j As Long
j = 0
res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\", j)
Do Until res = "Not Found"
   Set i = Treeview1.Nodes.Add(1, 4, , res, 2, 2)
   j = j + 1
   res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\", j)
Loop
Treeview1.Nodes(1).Selected = True
Call get_values
RaiseEvent Click
End Sub
Function rename_group(old_name As String, new_name As String) As Boolean
On Error GoTo errhand
For i1 = 2 To Treeview1.Nodes.Count + 1 ' for validity of the old_name
If LCase(Treeview1.Nodes(i1).Text) = LCase(old_name) Then Exit For
Next
For i2 = 2 To Treeview1.Nodes.Count 'for validity of new_name
If LCase(Treeview1.Nodes(i2).Text) = LCase(new_name) Then
MsgBox "Group " & NewString & " Already Exists", , "Err-1 Rename Group Error"
rename_group = False
Exit Function
End If
Next
EllReg.CreateNewKey "SOFTWARE\fonts\control1\Groups\" & new_name, HKEY_LOCAL_MACHINE
Dim res As Variant, valu1 As String, valu2 As String
Dim j As Long
j = 0
res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & old_name & "\", j)
Do Until res(2) = "Not Found"
   valu1 = res(1)
   valu2 = res(2)
   SetKeyValue "SOFTWARE\fonts\control1\Groups\" & new_name, valu1, valu2, REG_SZ
   j = j + 1
   res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & old_name & "\", j)
Loop
EllReg.DeleteSubkey HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & old_name & "\"
Treeview1.Nodes(i1).Text = new_name
Call get_values
rename_group = True
errhand:
Select Case Err.Number
Case 35600
rename_group = False
MsgBox "invalid  group to rename"
Exit Function
Err.Clear
End Select
End Function
Function rename_member(group_name As String, member_name As String, new_member_name As String) As Boolean
On Error GoTo errhand
For i1 = 2 To Treeview1.Nodes.Count + 1 ' for validity of the old_name
If LCase(Treeview1.Nodes(i1).Text) = LCase(group_name) Then Exit For
Next
For i2 = 1 To list1.ListItems.Count + 1 ' for validity of member name
If LCase(list1.ListItems(i2).Text) = LCase(member_name) Then Exit For
Next i2
For i3 = 1 To list1.ListItems.Count
If LCase(list1.ListItems(i3).Text) = LCase(new_member_name) Then
rename_member = False
MsgBox "Member name already exists"
Exit Function
End If
Next i3
res = EllReg.ReadRegistry(HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & group_name & "\", member_name)
SetKeyValue "Software\fonts\control1\groups\" & group_name & "\", new_member_name, res, REG_SZ
DeleteValue HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & group_name & "\", member_name
list1.ListItems(i2).Text = new_member_name
Call get_values
rename_member = True
RaiseEvent Click
errhand:
Select Case Err.Number
Case 35600
rename_member = False
MsgBox "invalid  member to rename"
Err.Clear
Exit Function
End Select
End Function
Function add_group(new_group_name As String) As Boolean
For i = 2 To Treeview1.Nodes.Count
If LCase(Treeview1.Nodes(i).Text) = LCase(new_group_name) Then
add_group = False
MsgBox "Invalid Group To add"
Exit Function
End If
Next i
Treeview1.Nodes.Add 1, 4, , new_group_name, 2
EllReg.CreateNewKey "SOFTWARE\fonts\control1\Groups\" & new_group_name, HKEY_LOCAL_MACHINE
Call get_values
End Function
Function add_member(group_name As String, member_name As String, file_name As String, file_path As String) As Boolean
On Error GoTo errhand
For i1 = 2 To Treeview1.Nodes.Count + 1
If LCase(Treeview1.Nodes(i1).Text) = LCase(group_name) Then Exit For
Next
Dim res As Variant, j As Long
res = EllReg.ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "Software\fonts\control1\groups\" & group_name & "\", j)
Do Until res(2) = "Not Found"
   If LCase(res(2)) = LCase(member_name) Then
   add_member = False
   MsgBox "Invalid Member To add:"
   Exit Function
   End If
   j = j + 1
   res = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, "SOFTWARE\fonts\control1\Groups\" & old_name & "\", j)
Loop
file_to_copy = file_path & file_name
folder_to_copy = group_folder & "\" & file_name
FileCopy file_to_copy, folder_to_copy
SetKeyValue "Software\fonts\control1\groups\" & group_name & "\", member_name, file_name, REG_SZ
errhand:
Select Case Err.Number
Case 35600
add_member = False
MsgBox "invalid group"
Err.Clear
    Exit Function
Case 58
    Resume Next
Case 53, 52, 76, 75
    add_member = False
    MsgBox "error accsing  File "
    Err.Clear
    Exit Function
End Select
End Function
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", "True")
End Sub
Private Sub UserControl_Resize()
Call resize_height(UserControl.height)
Call resize_width(UserControl.width)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, "True")
End Sub
Public Property Let Enabled(stat As Boolean)
Treeview1.Enabled = stat
list1.Enabled = stat
End Property
Public Property Get Enabled() As Boolean
Enabled = Treeview1.Enabled
End Property
Public Property Get f_name() As String
f_name = font_dat.name
End Property
Public Property Get file_size() As String
file_size = font_dat.filesize
End Property
Public Property Get file_location() As String
file_location = font_dat.font_file_location
End Property
Public Property Get group() As String
group = font_dat.group
End Property
Public Property Get file() As String
file = font_dat.file
End Property
Sub resize_width(width As Integer)
Dim wi As Integer
UserControl.width = width
wi = width - 10
Treeview1.width = wi
list1.width = wi
Picture3(0).width = wi
Picture3(1).width = wi
Picture1.width = wi
Picture2.width = wi
Command1(1).Left = Picture3(1).width - Command1(1).width
Command1(0).Left = Picture3(0).width - Command1(0).width
End Sub
Sub resize_height(height As Integer)
UserControl.height = height
list1.height = height / 2.45
Treeview1.height = height / 2.34
Treeview1.Top = Picture3(1).Top + Picture3(1).height
Picture2.Top = Treeview1.Top + Treeview1.height
Picture3(0).Top = Picture2.Top + 75
list1.Top = Picture3(0).Top + Picture3(0).height
End Sub

