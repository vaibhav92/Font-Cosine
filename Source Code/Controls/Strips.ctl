VERSION 5.00
Begin VB.UserControl Strips 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   ScaleHeight     =   540
   ScaleWidth      =   2430
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   1215
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Strips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event click()
Event mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyAscii As Integer)
'Default Property Values:
Const m_def_max = 1
Const m_def_value = 0
Const m_def_min = 0
'Property Variables:
Dim m_max As Long
Dim m_value As Long
Dim m_min As Long
Private Sub Label1_Click()
'Label1.FontBold = True
RaiseEvent click
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent mousedown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_GotFocus()
Label1.BorderStyle = 1
End Sub
Property Get value() As Long
value = m_value
End Property
Property Let value(valu As Long)
If valu > m_max Or valu < m_min Then Err.Raise 1, , "Invalid Value": Err.Clear: Exit Property
m_value = valu
Dim j As Long
j = UserControl.Width
Shape1(0).Width = m_value / m_max * j
Shape1(1).Width = j - Shape1(0).Width
Shape1(0).Left = 0
Shape1(1).Left = Shape1(0).Width
PropertyChanged "value"
End Property
Private Sub UserControl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Label1_Click
RaiseEvent KeyDown(KeyAscii)
End Sub
Private Sub UserControl_LostFocus()
'Label1.FontBold = False
Label1.BorderStyle = 0
End Sub
Private Sub UserControl_Resize()
Dim i As Long, j As Long
i = UserControl.Height
j = UserControl.Width
Label1.Height = i
Label1.Width = j
Shape1(0).Height = i
Shape1(1).Height = i
Shape1(0).Width = (m_value / m_max) * j
Shape1(1).Width = j - Shape1(0).Width
Shape1(0).Left = 0
Shape1(1).Left = Shape1(0).Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Label1.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Integer)
    Label1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Label1.Caption = PropBag.ReadProperty("Caption", "Strip1")
    m_min = PropBag.ReadProperty("min", m_def_min)
    m_max = PropBag.ReadProperty("max", m_def_max)
    Label1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Strip1")
    Call PropBag.WriteProperty("min", m_min, m_def_min)
    Call PropBag.WriteProperty("max", m_max, m_def_max)
    Call PropBag.WriteProperty("BorderStyle", Label1.BorderStyle, 0)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
Public Property Get min() As Long
    min = m_min
End Property

Public Property Let min(ByVal New_min As Long)
    m_min = New_min
    PropertyChanged "min"
End Property

Public Function changecolor(col As OLE_COLOR, flag As Long) As OLE_COLOR
Select Case flag
Case 0
Label1.ForeColor = col
Case 1
Shape1(0).BackColor = col
Case 2
Shape1(1).BackColor = col
End Select
End Function
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_value = m_def_value
    m_min = m_def_min
    m_max = m_def_max
End Sub
Public Property Get max() As Long
    max = m_max
End Property
Public Property Let max(ByVal New_max As Long)
    m_max = New_max
    PropertyChanged "max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Label1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Label1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

