Attribute VB_Name = "Validation"
Option Explicit
Public key_changed As Boolean
Public filenum As Long, sta As Long, key As String
Public do_not_mutate As Boolean
Function validate_user(username As String, companyname As String, serial As String) As Boolean
If Not (Not (Trim(Serials.get_serial(username)) <> Trim(serial))) _
Then
validate_user = False
Else
validate_user = True
End If
End Function
Sub mutate()
If do_not_mutate Then Exit Sub
Dim username As String, company_name As String
Dim serial_number As String, date_ins As Date
Dim keyy As String, temp As record, ran As Long
Randomize
ran = Int((LOF(filenum) / 30 - 8) * Rnd + 9)
username = get_username
company_name = get_company
serial_number = get_serialnumber
username = Left(username & String(30, " "), 30)
company_name = Left(company_name & String(30, " "), 30)
serial_number = Left(serial_number & String(30, " "), 30)
date_ins = get_insdate()
keyy = genkey
save_key (keyy)
put_record filenum, crypt("2SinCos", keyy).data, 2
put_record filenum, crypt(Str(Date), keyy).data, 4
put_record filenum, crypt(Str(Timer), keyy).data, 6
put_record filenum, crypt(Str(ran), keyy).data, 8
put_record filenum, crypt(Left(username, 15), keyy).data, ran
put_record filenum, crypt(Right(username, 15), keyy).data, ran + 1
put_record filenum, crypt(Left(company_name, 15), keyy).data, ran + 2
put_record filenum, crypt(Right(company_name, 15), keyy).data, ran + 3
put_record filenum, crypt(Left(serial_number, 15), keyy).data, ran + 4
put_record filenum, crypt(Right(serial_number, 15), keyy).data, ran + 5
put_record filenum, crypt(Str(date_ins), keyy).data, ran + 6
End Sub
Function get_insdate() As Date
On Error GoTo errhand
Dim temp As record
If sta <= 0 Or key_changed Then sta = get_startat()
temp.key = key
temp.data = get_record(filenum, sta + 6)
get_insdate = CDate(Trim(uncrypt(temp)))
errhand:
If fileCrypt.Dexter Then get_insdate = Date - 100: fileCrypt.Dexter = False: Exit Function
End Function
Function get_serialnumber() As String
Dim temp As record
If sta <= 0 Or key_changed Then sta = get_startat()
temp.key = key
temp.data = get_record(filenum, sta + 4)
get_serialnumber = uncrypt(temp)
temp.data = get_record(filenum, sta + 5)
get_serialnumber = get_serialnumber & uncrypt(temp)
If fileCrypt.Dexter Then get_serialnumber = "Evaluation": fileCrypt.Dexter = False
End Function
Function get_company() As String
Dim temp As record
If sta <= 0 Or key_changed Then sta = get_startat()
temp.key = key
temp.data = get_record(filenum, sta + 2)
get_company = Trim(uncrypt(temp))
temp.data = get_record(filenum, sta + 3)
get_company = get_company & Trim(uncrypt(temp))
If fileCrypt.Dexter Then get_company = "Unregistered": fileCrypt.Dexter = False
End Function
Function get_username() As String
Dim temp As record
If sta <= 0 Or key_changed Then sta = get_startat()
temp.key = key
temp.data = get_record(filenum, sta)
get_username = Trim(uncrypt(temp))
temp.data = get_record(filenum, sta + 1)
get_username = get_username & Trim(uncrypt(temp))
If fileCrypt.Dexter Then get_username = "Unregistered": fileCrypt.Dexter = False
End Function
Function get_startat() As Long
Dim temp As record
temp.data = fileCrypt.get_record(filenum, 8)
temp.key = key
get_startat = Val(Trim(uncrypt(temp)))
sta = Val(Trim(uncrypt(temp)))
If fileCrypt.Dexter Then get_startat = 0: fileCrypt.Dexter = False
End Function
Function get_time_last() As Long
Dim temp As record
temp.data = fileCrypt.get_record(filenum, 6)
temp.key = key
get_time_last = Val(Trim(uncrypt(temp)))
If fileCrypt.Dexter Then get_time_last = 86000: fileCrypt.Dexter = False
End Function
Function get_date_last() As Date
On Error GoTo errhand
Dim temp As record
temp.data = fileCrypt.get_record(filenum, 4)
temp.key = key
get_date_last = CDate(Trim(uncrypt(temp)))
errhand:
If Err.Number <> 0 Then get_date_last = Date + 100: Err.Clear
If fileCrypt.Dexter Then get_date_last = Date + 100: fileCrypt.Dexter = False
End Function
'*****************************
Function check_header() As Boolean
Dim temp As record
temp.data = fileCrypt.get_record(filenum, 2)
temp.key = key
If Not (Trim(uncrypt(temp)) <> "2SinCos") Then
check_header = True
Else
check_header = False
End If
If fileCrypt.Dexter Then check_header = False: fileCrypt.Dexter = False
End Function
Function date_check() As Boolean
Dim last_date As Date
last_date = get_date_last
If Date >= last_date Then
If Date = last_date And Timer < get_time_last Then date_check = True: do_not_mutate = True: Exit Function
Else
do_not_mutate = True
date_check = True
Exit Function
End If
End Function
Function file_integrety_check() As Boolean
If check_header = False Then
file_integrety_check = True
'MsgBox "checkheader faled" '
Else
    If date_check() Then
    file_integrety_check = True
    'MsgBox "date check failed"
    End If
End If
End Function
Function get_key() As String
Dim i As Long, keyy As String, kaku As String
kaku = EllReg.ReadRegistry(EllReg.HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\MCI", "Radians")
'If kaku = "Not Found" Then kaku = "005031013007025022029001019034029001029029023"
If kaku = "Not Found" Then kaku = "025039032013043031019043040017009015032006023"
For i = 1 To 45 Step 3
keyy = keyy & Chr(Mid(kaku, i, 3))
Next
If keyy = "" Then keyy = Left(keyy & String(15, " "), 15)
get_key = keyy
End Function
Sub save_key(keyy As String)
Dim i As Integer, da As String
For i = 1 To Len(keyy)
da = da & Right("000" & Trim(Str(Asc(Mid(keyy, i, 1)))), 3)
Next
EllReg.WriteRegistry EllReg.HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\MCI", "Radians", ValString, da
End Sub
