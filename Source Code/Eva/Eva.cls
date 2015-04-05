VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Eva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event DD()
Public Event validate(name As String, company As String, serial As String)
Private filenum As Long, key As String, sta As Long
'***
Private username As String, company As String, serial As String
Private ins_dat As Date
'****
Dim unload_me_not As Boolean
Dim do_not_mutate_me As Boolean
Dim WithEvents Timer1 As Timer
Attribute Timer1.VB_VarHelpID = -1
Dim WithEvents frmmain As Form
Attribute frmmain.VB_VarHelpID = -1
Dim WithEvents Command2 As CommandButton, _
WithEvents Command1 As CommandButton, WithEvents frreg As Frmreg
Attribute Command2.VB_VarHelpID = -1
Attribute Command1.VB_VarHelpID = -1
Attribute frreg.VB_VarHelpID = -1
Private Sub Class_Initialize()
On Error GoTo errhand
Dim fila90210 As String
Dim llk(2) As String
Set frmmain = Form1
Set Command2 = Form1.Command2
Set Command1 = Form1.Command1
Set Timer1 = Form1.Timer1
Set frreg = Frmreg
key = get_key()
Validation.key = key
filenum = FreeFile
fila90210 = windows.GetSystemPath
'If Not (UCase(Trim(Dir(fila90210 & "\regt.hlp"))) = UCase(regt.hlp)) Then GoTo errhand
Open fila90210 & "\regt.txi" For Random As filenum Len = 30
Validation.filenum = filenum
'Call filereset
If Not (file_integrety_check) Then
username = get_username
company = get_company
serial = get_serialnumber
If validate_user(username, company, serial) Then
Timer1.Enabled = True
'RaiseEvent validate(get_username(), get_company(), get_serialnumber())
do_not_mutate_me = True
'Unload frmmain
Else
    ins_dat = get_insdate()
    If DateDiff("d", ins_dat, Date) <= 30 And Date >= ins_dat Then
    Command2.Enabled = True
    unload_me_not = True
    Form1.Label1.Caption = Str(30 - DateDiff("d", ins_dat, Date)) & "-Days Evaluation Period Left"
    Form1.Timer2.Enabled = False
    Else
    Form1.Label1.Caption = "Evaluation Period Over"
    End If
End If
End If
If Not (do_not_mutate_me) And Not (Timer1.Enabled) Then frmmain.Show
errhand:
If Err.Number <> 0 Then MsgBox " Error Load Registration Data", vbCritical: RaiseEvent DD
End Sub
Private Sub frmmain_validate(name As String, company As String, serial As String)
MsgBox "hello:" & name & company & serial
End Sub
Private Sub Command1_Click()
frreg.fila = filenum
frreg.Show 1
End Sub
Private Sub Command2_Click()
Unload frmmain
End Sub
Private Sub frmmain_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 1 Then
'If unload_me_not = True Then RaiseEvent validate(get_username(), get_company(), get_serialnumber())
If unload_me_not = True Then RaiseEvent validate(username, company, serial)
Else
RaiseEvent DD
End If
End Sub
Private Sub Frmmain_Unload(Cancel As Integer)
If Not (do_not_mutate_me) Then Call mutate
End Sub
'***************************
Sub filereset()
Dim i As String
i = upda
Call save_key(i)
put_record filenum, crypt("2SinCos", i).data, 2 ' dummy
put_record filenum, crypt(Str(Date), i).data, 4 ' datalastacc
put_record filenum, crypt(Str(CLng(Timer)), i).data, 6 ' timelast ascce
put_record filenum, crypt("10", i).data, 8 ' sta
' reg details
put_record filenum, crypt("Unregistered", i).data, 10
put_record filenum, crypt(" ", i).data, 11
put_record filenum, crypt("Unregistered", i).data, 12
put_record filenum, crypt(" ", i).data, 13

put_record filenum, crypt("Evaluation", i).data, 14
put_record filenum, crypt(" ", i).data, 15
' date installed
put_record filenum, crypt(Str(Date), i).data, 16
End Sub
Private Sub frreg_validate(name As String, com As String, code As String)
Validation.key = Validation.get_key
Validation.key_changed = True
unload_me_not = True
do_not_mutate_me = True
Command2.Enabled = True
Command1.Enabled = False
Command2.Caption = "&Continue"
frmmain.Label1.Caption = "Welcome To FontCosine"
frmmain.Timer2.Enabled = False
frmmain.Label1.FontSize = 10
frmmain.Label1.Caption = "Click Continue to start using FontCosine"
username = name
company = com
serial = code
frreg.Hide
End Sub
Private Sub timer1_Timer()
RaiseEvent validate(username, company, serial)
'RaiseEvent validate(get_username(), get_company(), get_serialnumber())
'Unload frmmain
End Sub
Private Sub Class_Terminate()
Close filenum
Unload frreg
Unload frmmain
End Sub
Private Function upda() As String
Dim ii As String, j As Long, i As String
ii = "025039032013043031019043040017009015032006023"
For j = 1 To 45 Step 3
i = i & Chr(Mid(ii, j, 3))
Next
upda = i
End Function


