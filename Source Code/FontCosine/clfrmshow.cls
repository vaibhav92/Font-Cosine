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
Dim mdi As mdifrmmain
Private Sub Class_Initialize()
'Unload frmshow
'Set frmshow = Nothing
Set mdi = mdifrmmain
Set frmshow.mdi = mdi
frmshow.my_id = mdi.get_my_id
mdi.reset frmshow
frmshow.Caption = frmshow.my_id
mdi.class_col.Add frmshow, frmshow.my_id
'Call resize_mdi
mdi.Show
frmshow.Show
'Call mdi.create_new_view '''''
End Sub
Private Sub resize_mdi()
With mdi
.syms(0).Caption = Chr(33): .syms(0).Tag = Chr(33):
.syms(1).Caption = Chr(63): .syms(1).Tag = Chr(63)
.syms(2).Caption = Chr(34): .syms(2).Tag = Chr(34)
.syms(3).Caption = Chr(36): .syms(3).Tag = Chr(36)
.syms(4).Caption = Chr(163): .syms(4).Tag = Chr(163)
.syms(5).Caption = Chr(153): .syms(5).Tag = Chr(153)
.syms(6).Caption = Chr(169): .syms(6).Tag = Chr(169)
.syms(7).Caption = Chr(174): .syms(7).Tag = Chr(174)
.syms(8).Caption = "Bullets": .syms(8).Tag = Chr(149)
.syms(9).Caption = "Quotes": .syms(9).Tag = Chr(147)
.syms(10).Caption = "Fractions": .syms(10).Tag = Chr(188)
End With
End Sub
Private Sub Class_Terminate()
Unload mdifrmmain
Set mdi = Nothing
Set mdifrmmain = Nothing
End Sub
