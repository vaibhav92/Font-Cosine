VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   PropertyPages   =   "spinner2.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   1350
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1060
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   661
      _Version        =   327681
      Value           =   1
      Alignment       =   0
      Max             =   2
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "The down button on the UpDown control has been clicked"
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Private Sub UpDown1_Change()
Text1.Text = Str(UpDown1.Value)
UpDown1.Max = UpDown1.Max + 1
End Sub
Private Sub UserControl_Initialize()
UpDown1.Value = Val(Ini_value)
Text1.Text = Str(UpDown1.Value)
End Sub
Private Sub UserControl_Resize()
Text1.Height = UserControl.Height
Text1.Width = UserControl.Width * 0.785185
Text1.Left = 0
UpDown1.Height = Text1.Height
UpDown1.Top = Text1.Top
UpDown1.Left = Text1.Left + Text1.Width + 20
End Sub
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UpDown1,UpDown1,-1,Max
Public Property Get Max_Value() As Long
Attribute Max_Value.VB_Description = "Get/Set the upper bound of the scroll range"
Attribute Max_Value.VB_ProcData.VB_Invoke_Property = "General"
    Max_Value = UpDown1.Max
End Property

Public Property Let Max_Value(ByVal New_Max_Value As Long)
    UpDown1.Max() = New_Max_Value
    PropertyChanged "Max_Value"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,Text
'Public Property Get Ini_value() As String
'    Ini_value = Text1.Text
'End Property
'
'Public Property Let Ini_value(ByVal New_Ini_value As String)
'    Text1.Text() = New_Ini_value
'    PropertyChanged "Ini_value"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UpDown1,UpDown1,-1,Min
Public Property Get Min_Value() As Long
Attribute Min_Value.VB_Description = "Get/Set the lower bound of the scroll range"
Attribute Min_Value.VB_ProcData.VB_Invoke_Property = "General"
    Min_Value = UpDown1.Min
End Property

Public Property Let Min_Value(ByVal New_Min_Value As Long)
    UpDown1.Min() = New_Min_Value
    PropertyChanged "Min_Value"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    UpDown1.Max = PropBag.ReadProperty("Max_Value", 2)
    Text1.Text = PropBag.ReadProperty("Ini_value", "")
    UpDown1.Min = PropBag.ReadProperty("Min_Value", 0)
    UpDown1.Value = PropBag.ReadProperty("Ini_value", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Max_Value", UpDown1.Max, 2)
    Call PropBag.WriteProperty("Ini_value", Text1.Text, "")
    Call PropBag.WriteProperty("Min_Value", UpDown1.Min, 0)
    Call PropBag.WriteProperty("Ini_value", UpDown1.Value, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UpDown1,UpDown1,-1,Value
Public Property Get Ini_value() As Long
Attribute Ini_value.VB_Description = "Get/Set the current position in the scroll range"
Attribute Ini_value.VB_ProcData.VB_Invoke_Property = "General"
    Ini_value = UpDown1.Value
End Property

Public Property Let Ini_value(ByVal New_Ini_value As Long)
    UpDown1.Value() = New_Ini_value
    PropertyChanged "Ini_value"
End Property

