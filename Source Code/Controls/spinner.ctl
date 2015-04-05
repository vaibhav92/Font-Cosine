VERSION 5.00
Begin VB.UserControl Spinner 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   ScaleHeight     =   360
   ScaleWidth      =   1245
   ToolboxBitmap   =   "spinner.ctx":0000
   Begin VB.VScrollBar VScroll1 
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   30
      Width           =   225
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   1245
   End
End
Attribute VB_Name = "Spinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event click()
Private Sub Text1_Change()
On Error GoTo errhand
VScroll1.value = Val(Text1.Text)
RaiseEvent click
errhand:
If Err.Number > 0 Then
Text1.Text = VScroll1.value
Err.Clear
Exit Sub
End If
End Sub
Private Sub Text1_GotFocus()
Text1.SelStart = 1
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub UserControl_Initialize()
VScroll1.Top = 30
End Sub
Private Sub UserControl_Resize()
Text1.Height = UserControl.Height
Text1.Width = UserControl.Width
VScroll1.Height = Text1.Height - 60
VScroll1.Left = Text1.Width - VScroll1.Width - 10
End Sub
Private Sub VScroll1_Change()
Text1.Text = Str(VScroll1.value)
RaiseEvent click
End Sub
Public Property Let max(i As Integer)
VScroll1.max = i
PropertyChanged "max"
End Property
Public Property Let min(i As Integer)
VScroll1.min = i
PropertyChanged "min"
End Property
Public Property Let value(i As Integer)
VScroll1.value = i
PropertyChanged "value"
End Property
Public Property Let smallchange(i As Integer)
VScroll1.smallchange = i
PropertyChanged "smallchange"
End Property
Public Property Let largechange(i As Integer)
VScroll1.largechange = i
PropertyChanged "largechange"
End Property
Public Property Get max() As Integer
max = VScroll1.max
End Property
Public Property Get min() As Integer
min = VScroll1.min
End Property
Public Property Get value() As Integer
value = VScroll1.value
End Property
Public Property Get smallchange() As Integer
smallchange = VScroll1.smallchange
End Property
Public Property Get largechange() As Integer
largechange = VScroll1.largechange
End Property
Public Property Let fontsize(i As Integer)
Text1.fontsize = i
PropertyChanged "smallchange"
End Property
Private Sub VScroll1_Scroll()
Call VScroll1_Change
End Sub
