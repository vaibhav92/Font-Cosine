VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Myapp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents ev As Eva
Attribute ev.VB_VarHelpID = -1
Dim mdi As mdifrmmain
Private Sub Class_Initialize()
frmSplash.Show
Set ev = New Eva
End Sub
Private Sub ev_validate(name As String, company As String, serial As String)
frmAbout.lblcompany.Caption = company
frmAbout.lblserial.Caption = serial
frmAbout.lblusername.Caption = name
Set ev = Nothing
'Set Startup.myappa = Nothing
Set mdi = mdifrmmain
Set frmshow.mdi = mdi
frmshow.my_id = mdi.get_my_id
mdi.reset frmshow
frmshow.Caption = frmshow.my_id
mdi.class_col.Add frmshow, frmshow.my_id
mdi.Show
'*********
'*******
frmshow.Show
frmSplash.Hide
Unload frmSplash
Set ev = Nothing
mdi.Timer2.enabled = True
End Sub
Private Sub ev_DD()
Set ev = Nothing
Unload mdifrmmain
Unload frmSplash
Unload frmshow
Unload frmOptions
Unload frmAbout
Set Startup.myappa = Nothing
End Sub
'Private Sub Class_Terminate()
'Unload mdifrmmain
'Unload frmSplash
'Unload frmshow
'Unload frmOptions
'Unload frmAbout
'End Sub
