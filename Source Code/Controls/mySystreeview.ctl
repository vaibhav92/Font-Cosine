VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SysTreeView 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ScaleHeight     =   5850
   ScaleWidth      =   3315
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   10345
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2070
      TabIndex        =   1
      Top             =   2010
      Width           =   270
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   2010
      Width           =   105
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2070
      TabIndex        =   2
      Top             =   2010
      Width           =   135
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "mySystreeview.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview.ctx":27B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview.ctx":4F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mySystreeview.ctx":771C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SysTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_BackColor = &H8000000F
Const m_def_Selected = 0
Const m_def_Show_Files = 0
'Property Variables:
Dim m_BackColor As OLE_COLOR
Dim m_Selected As Boolean
Dim m_Show_Files As Boolean
'Event Declarations:
Event Click() 'MappingInfo=TreeView1,TreeView1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=TreeView1,TreeView1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
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
Event PathChange() 'MappingInfo=File1,File1,-1,PathChange
Attribute PathChange.VB_Description = "Occurs when the path is changed by setting the FileName or Path property in code."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Private Sub usercontrol_initialize()
ChDir "c:\"
TreeView1.ImageList = ImageList1
Dim i As Node
Set i = UserControl.TreeView1.Nodes.add(, , "r", "My Computer", 1)
For j = 1 To (UserControl.Drive1.ListCount - 1)
Set i = UserControl.TreeView1.Nodes.add("r", 4, , Drive1.list(j), 2)
Next j
End Sub
Private Sub get_subdir(Path As String)
If UserControl.TreeView1.SelectedItem.Tag = "1" Then Exit Sub
UserControl.Dir1.Path = Path
Dim i As Node
For j = 0 To UserControl.Dir1.ListCount - 1
Set i = UserControl.TreeView1.Nodes.add(TreeView1.SelectedItem.Index, 4, , Dir1.list(j), 3)
Next
If m_Show_Files = True Then Call Get_files(Path)
UserControl.TreeView1.SelectedItem.Tag = "1"
End Sub
Private Sub Get_files(Path As String)
UserControl.File1.Path = Path
Dim i As Node
For j = 0 To (UserControl.File1.ListCount - 1)
Set i = UserControl.TreeView1.Nodes.add(TreeView1.SelectedItem.Index, 4, , File1.list(j), 4)
Next
End Sub
Private Sub TreeView1_Click()
    RaiseEvent Click
If UserControl.TreeView1.SelectedItem.Index > 1 Then
If UserControl.TreeView1.SelectedItem.Image <> 4 Then
Call get_subdir(UserControl.TreeView1.SelectedItem.Text)
End If
End If
End Sub
Private Sub UserControl_Resize()
    RaiseEvent Resize
UserControl.TreeView1.Height = UserControl.Height
UserControl.TreeView1.Width = UserControl.Width
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
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
    Enabled = UserControl.TreeView1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.TreeView1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.TreeView1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.TreeView1.Font = New_Font
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
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    UserControl.TreeView1.Refresh
End Sub

Private Sub TreeView1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,Path
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns/sets the current path."
    Path = UserControl.File1.Path
End Property

Public Property Let Path(ByVal New_Path As String)
    UserControl.File1.Path() = New_Path
    PropertyChanged "Path"
End Property

Private Sub File1_PathChange()
    RaiseEvent PathChange
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,PathSeparator
Public Property Get PathSeparator() As String
Attribute PathSeparator.VB_Description = "Returns/sets the delimiter string used for the path returned by the FullPath property."
    PathSeparator = UserControl.TreeView1.PathSeparator
End Property

Public Property Let PathSeparator(ByVal New_PathSeparator As String)
    UserControl.TreeView1.PathSeparator() = New_PathSeparator
    PropertyChanged "PathSeparator"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,Pattern
Public Property Get Pattern() As String
Attribute Pattern.VB_Description = "Returns/sets a value indicating the filenames displayed in a control at run time."
    Pattern = UserControl.File1.Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    UserControl.File1.Pattern() = New_Pattern
    PropertyChanged "Pattern"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Scroll
Public Property Get Scroll() As Boolean
Attribute Scroll.VB_Description = "Returns/sets a value which determines if the TreeView displays scrollbars and allows scrolling (vertical and horizontal)."
    Scroll = UserControl.TreeView1.Scroll
End Property

Public Property Let Scroll(ByVal New_Scroll As Boolean)
    UserControl.TreeView1.Scroll() = New_Scroll
    PropertyChanged "Scroll"
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets the selection status of an item in a control."
    Selected = m_Selected
End Property

Public Property Let Selected(ByVal New_Selected As Boolean)
    m_Selected = New_Selected
    PropertyChanged "Selected"
End Property

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = UserControl.TreeView1.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    UserControl.TreeView1.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,System
Public Property Get System() As Boolean
Attribute System.VB_Description = "Determines whether a FileListBox control displays files with System attributes."
    System = UserControl.File1.System
End Property

Public Property Let System(ByVal New_System As Boolean)
    UserControl.File1.System() = New_System
    PropertyChanged "System"
End Property

Public Property Get Show_Files() As Boolean
    Show_Files = m_Show_Files
End Property

Public Property Let Show_Files(ByVal New_Show_Files As Boolean)
    m_Show_Files = New_Show_Files
    PropertyChanged "Show_Files"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_Selected = m_def_Selected
    m_Show_Files = m_def_Show_Files
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.TreeView1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.File1.Path = PropBag.ReadProperty("Path", "")
    UserControl.TreeView1.PathSeparator = PropBag.ReadProperty("PathSeparator", "\")
    UserControl.File1.Pattern = PropBag.ReadProperty("Pattern", "*.*")
    UserControl.TreeView1.Scroll = PropBag.ReadProperty("Scroll", True)
    m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
    UserControl.TreeView1.Sorted = PropBag.ReadProperty("Sorted", False)
    UserControl.File1.System = PropBag.ReadProperty("System", False)
    m_Show_Files = PropBag.ReadProperty("Show_Files", m_def_Show_Files)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", TreeView1.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Path", File1.Path, "")
    Call PropBag.WriteProperty("PathSeparator", TreeView1.PathSeparator, "\")
    Call PropBag.WriteProperty("Pattern", File1.Pattern, "*.*")
    Call PropBag.WriteProperty("Scroll", TreeView1.Scroll, True)
    Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
    Call PropBag.WriteProperty("Sorted", TreeView1.Sorted, False)
    Call PropBag.WriteProperty("System", File1.System, False)
    Call PropBag.WriteProperty("Show_Files", m_Show_Files, m_def_Show_Files)
End Sub

