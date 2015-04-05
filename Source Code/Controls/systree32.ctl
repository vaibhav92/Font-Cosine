VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl systree 
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   ScaleHeight     =   5805
   ScaleWidth      =   3930
   ToolboxBitmap   =   "systree32.ctx":0000
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5745
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   10134
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2340
      Pattern         =   "*.txt"
      TabIndex        =   3
      Top             =   2040
      Width           =   105
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2130
      TabIndex        =   2
      Top             =   2040
      Width           =   135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   2040
      Width           =   270
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2460
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
            Picture         =   "systree32.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "systree32.ctx":2AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "systree32.ctx":527A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "systree32.ctx":7A2E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "systree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim caption As String
'Default Property Values:
Const m_def_caption = "My Computer"
Const m_def_show_files = 0
'Property Variables:
Dim m_caption As String
Dim m_show_files As Boolean
'Event Declarations:
Event Click() 'MappingInfo=TreeView1,TreeView1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=TreeView1,TreeView1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TreeView1,TreeView1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TreeView1,TreeView1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TreeView1,TreeView1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TreeView1,TreeView1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TreeView1,TreeView1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TreeView1,TreeView1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Private Sub get_subdir(Path As String)
'On Error GoTo errhand
If Treeview1.SelectedItem.Children > 0 Then Exit Sub
Dir1.Path = Path
Dim i As Node
For j = 0 To Dir1.ListCount - 1
Set i = Treeview1.Nodes.Add(Treeview1.SelectedItem.index, 4, , getdir(Dir1.List(j)), 3)
Next
If show_files = True Then Call Get_files(Path)
errhand:
Select Case Err().Number
    Case 68
    Exit Sub
End Select
End Sub
Private Sub Get_files(Path As String)
File1.Path = Path
Dim i As Node
For j = 0 To (File1.ListCount - 1)
Set i = Treeview1.Nodes.Add(Treeview1.SelectedItem.index, 4, , File1.List(j), 4)
Next
End Sub
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub
Private Sub TreeView1_Click()
    RaiseEvent Click
If Treeview1.SelectedItem.Image <> 4 Then
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
pa = Treeview1.Nodes(1).Tag & pa
get_path = pa
End Function
Private Sub usercontrol_initialize()
Treeview1.ImageList = ImageList1
Dim i As Variant
Set i = Treeview1.Nodes.Add(, , "r", m_caption, 1)
Treeview1.Nodes(1).Tag = File1.Path
End Sub
Private Sub UserControl_Resize()
Treeview1.Height = UserControl.Height
Treeview1.Width = UserControl.Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Treeview1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Treeview1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Treeview1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Treeview1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Treeview1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    Treeview1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    Treeview1.Refresh
End Sub

Private Sub TreeView1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 '   RaiseEvent MouseMove(Button, Shift, x, y)
'End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,Pattern
Public Property Get Pattern() As String
Attribute Pattern.VB_Description = "Returns/sets a value indicating the filenames displayed in a control at run time."
    Pattern = File1.Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    File1.Pattern() = New_Pattern
    PropertyChanged "Pattern"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,Path
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns/sets the current path."
    Path = File1.Path
End Property

Public Property Let Path(ByVal New_Path As String)
    File1.Path() = New_Path
    Dir1.Path = New_Path
    PropertyChanged "Path"
End Property
Public Property Get caption() As String
      caption = m_caption
End Property
Public Property Let caption(ByVal New_caption As String)
    m_caption = New_caption
    Treeview1.Nodes(1).Text = m_caption
    Treeview1.Refresh
    PropertyChanged "caption"
End Property

Public Property Get show_files() As Boolean
    show_files = m_show_files
End Property

Public Property Let show_files(ByVal New_show_files As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_show_files = New_show_files
    PropertyChanged "show_files"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_caption = m_def_caption
    m_show_files = m_def_show_files
Treeview1.Nodes(1).Text = m_caption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Treeview1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Treeview1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    File1.Pattern = PropBag.ReadProperty("Pattern", "*.txt")
    File1.Path = PropBag.ReadProperty("Path", "")
    m_caption = PropBag.ReadProperty("caption", m_def_caption)
    m_show_files = PropBag.ReadProperty("show_files", m_def_show_files)
    Treeview1.LineStyle = PropBag.ReadProperty("LineStyle", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Treeview1.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", Treeview1.BorderStyle, 0)
    Call PropBag.WriteProperty("Pattern", File1.Pattern, "*.txt")
    Call PropBag.WriteProperty("Path", File1.Path, "")
    Call PropBag.WriteProperty("caption", m_caption, m_def_caption)
    Call PropBag.WriteProperty("show_files", m_show_files, m_def_show_files)
    Call PropBag.WriteProperty("LineStyle", Treeview1.LineStyle, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,LineStyle
Public Property Get LineStyle() As TreeLineStyleConstants
Attribute LineStyle.VB_Description = "Returns/sets the style of lines displayed between Node objects."
    LineStyle = Treeview1.LineStyle
End Property

Public Property Let LineStyle(ByVal New_LineStyle As TreeLineStyleConstants)
    Treeview1.LineStyle() = New_LineStyle
    PropertyChanged "LineStyle"
End Property

