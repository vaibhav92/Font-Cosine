VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public locked As Boolean
Private WithEvents myform As frmshow, _
my_mdi As mdifrmmain
Attribute myform.VB_VarHelpID = -1
Public my_id As String
Event sizecha(size As Integer)
Event cluchanged()
Public Sub refresh()
Call myform.con(myform.cur_con).refresh
If myform.cur_con = 1 Then _
myform.Picture1.refresh
End Sub
Private Sub relocate()
With myform.Label1
.Visible = False
.Top = (myform.Picture1.Height - .Height) / 2
.Left = (myform.Picture1.Width - .Width) / 2
.Visible = True
myform.Picture1.refresh
End With
End Sub
Public Property Let fontsize(ByVal new_size As Long)
myform.con(myform.cur_con).fontsize = new_size
myform.Slider1.Value = new_size
If myform.cur_con = 1 Then Call relocate
myform.font_size = new_size
myform.Check1.caption = myform.font_name & _
"-" & Str(myform.font_size)
'Call relocate
End Property
Public Property Let fontbold(ByVal new_val As Boolean)
myform.con(myform.cur_con).fontbold = new_val
myform.font_bold = new_val
Call relocate
End Property
Public Property Let fontitalic(ByVal new_val As Boolean)
myform.con(myform.cur_con).fontitalic = new_val
myform.font_italics = new_val
Call relocate
End Property
Public Property Let fontname(ByVal new_font As String)
myform.con(myform.cur_con).fontname = new_font
If myform.cur_con = 1 Then Call relocate
myform.font_name = new_font
myform.Check1.caption = myform.font_name & _
"-" & Str(myform.font_size)
Call relocate
End Property
Public Property Let fontunderline(ByVal new_val As Boolean)
myform.con(myform.cur_con).fontunderline = new_val
myform.font_under = new_val
Call relocate
End Property
Public Property Let fontstrikethru(ByVal new_font As Boolean)
myform.con(myform.cur_con).fontstrikethru = new_font
myform.font_strikethru = new_font
Call relocate
End Property
'************
Public Property Get fontsize() As Long
fontsize = myform.con(myform.cur_con).fontsize
End Property
Public Property Get fontbold() As Boolean
fontbold = myform.con(myform.cur_con).fontbold
End Property
Public Property Get fontitalic() As Boolean
fontitalic = myform.con(myform.cur_con).fontitalic
End Property
Public Property Get fontname() As String
fontname = myform.con(myform.cur_con).fontname
End Property
Public Property Get fontunderline() As Boolean
fontunderline = myform.con(myform.cur_con).fontunderline
End Property
Public Property Get fontstrikethru() As Boolean
fontstrikethru = myform.con(myform.cur_con).fontstrikethru
End Property
Property Set mdi(md As mdifrmmain)
Set my_mdi = md
Let my_id = my_mdi.get_my_id
my_mdi.class_col.Add Me, my_id
End Property
Property Get mdi() As mdifrmmain
Set mdi = my_mdi
End Property
Property Set my_form(orm As frmshow)
If Not (my_form Is Nothing) Then Unload myform: Set myform = Nothing
Set myform = orm
myform.Show
End Property
Property Get my_form() As frmshow
Set my_form = myform
End Property
Private Sub Class_Initialize()
Dim i As New frmshow
'i.Form_Resize
Set myform = i
'Set myform = New frmshow
myform.Show
End Sub
Private Sub myform_igotfocus()
Set my_mdi.clu = Me
RaiseEvent cluchanged
'RaiseEvent sizecha(myform.con(myform.cur_con).FontSize)
'sizecha (myform.con(myform.curcon).FontSize)
End Sub
Public Property Let caption(ByVal new_caption As String)
    myform.caption = new_caption
End Property
Public Property Get caption() As String
    caption = myform.caption
End Property
Public Property Let Text(ByVal new_text As String)
    myform.Text = new_text
End Property
Public Property Get Text() As String
    caption = myform.Text
End Property
Public Property Let forecolor(ByVal new_col As OLE_COLOR)
myform.con(myform.cur_con).forecolor = new_col
myform.my_fore_col = new_col
End Property
Public Property Get forecolor() As OLE_COLOR
forecolor = myform.my_fore_col
End Property
Public Property Let backcolor(ByVal new_col As OLE_COLOR)
    myform.con(myform.cur_con).backcolor = new_col
     If myform.cur_con = 1 Then myform.Picture1.backcolor = new_col
     myform.my_back_col = new_col
End Property
Public Property Get backcolor() As OLE_COLOR
backcolor = myform.my_back_col
End Property
Private Sub myform_locker(va As Boolean)
locked = va
If va = True Then
my_mdi.locked_clu.Add Me, my_id
Else
my_mdi.locked_clu.Remove my_id
End If
End Sub
'Private Sub myform_sizecha()
'RaiseEvent sizecha(myform.font_size)
'End Sub
