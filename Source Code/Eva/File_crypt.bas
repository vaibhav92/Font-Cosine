Attribute VB_Name = "fileCrypt"
Public Dexter As Boolean
Option Explicit
Type record
key As String * 15
data As String * 15
End Type
Function get_record(filenumber As Long, rec_number As Long, Optional key As String) As String
On Error GoTo errhand
Dim i As record, dat As String
Get filenumber, rec_number, i
If Not (key = "") Then
i.key = key
dat = uncrypt(i)
Else
dat = uncrypt(i)
End If
get_record = dat
errhand:
If Err.Number <> 0 Then Dexter = True: Err.Clear
End Function
Sub put_record(filenumber As Long, data As String, _
Optional rec_no As Long, Optional key As String)
Dim rec As record, temp As Long
temp = (LOF(filenumber) / 30 + 1)
'If rec_no = 0 Or rec_no > (LOF(filenumber) / 30 + 1) Then rec_no = LOF(filenumber) / 30 + 1
'***
If rec_no = 0 Then rec_no = LOF(filenumber) / 30 + 1
Do While rec_no > ((LOF(filenumber) / 30) + 1)
rec.data = "sdaqaUIJHGHHGJH"
rec.key = "111111111111111"
Put #filenumber, temp, rec
temp = temp + 1
Loop
'****
If key = "" Then rec = crypt(data) Else _
rec = crypt(data, key)
Put #filenumber, rec_no, rec
End Sub
Function genkey() As String
Dim key As String, i As Long
Randomize
For i = 1 To 15
key = key & Chr(Int(45 * Rnd))
Next
genkey = key
End Function
Function crypt(data As String, Optional keyy As String) As record
Dim dat As String, i As Long, k As String, ii As Long, _
key As String, dataa As String
data = Left(data & String(15, " "), 15)
Randomize
For i = 1 To 15
k = Mid(data, i, 1)
If Not (keyy = "") Then
ii = Asc(Mid(keyy, i, 1))
Else
ii = Int(45 * Rnd)
End If
k = Chr(Asc(k) + ii)
key = key & Chr(Val(ii))
dataa = dataa & k
Next
crypt.data = dataa
crypt.key = key
End Function
Function uncrypt(data As record) As String
On Error GoTo errhand:
Dim dat As String, i As Long, k As Long
For i = 1 To 15
k = Asc(Mid(data.key, i, 1))
dat = dat & Chr(Asc(Mid(data.data, i, 1)) - k)
Next
uncrypt = dat
errhand:
If Err.Number <> 0 Then uncrypt = "Not-Known": Err.Clear: Dexter = True
End Function


