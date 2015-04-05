Attribute VB_Name = "Serials"
Option Explicit
Function get_serial(name As String) As String
Dim serial As String
Dim mat() As Currency
On Error GoTo errhand
If Len(name) > 30 Then name = Left(name, 30)
If Len(name) < 30 Then name = Left(name & String(30, " "), 30)
Dim i As Long, j As Long
ReDim mat(Len(name) - 1)
For i = 1 To Len(name)
mat(i - 1) = Asc(Mid(name, i, 1))
Next '''''''''step1

For j = 1 To Val(Right(Str(Asc(Right(name, 1))), 1))
For i = 0 To UBound(mat, 1)
Do While Len(Str(mat(i))) <= 3
'Int(90 / (Len(name) * 3))
mat(i) = (mat(i) * 2) + 1
Loop
Next 'step 2
For i = 1 To (UBound(mat, 1))
mat(i) = mat(i) + mat(i - 1)
Next i
For i = 0 To UBound(mat, 1)
mat(i) = clocker(mat(i))
Next 'step 2
Next j

For i = 0 To UBound(mat, 1)
j = (mat(i))
If (j >= 97 And j <= 122) Or (j >= 65 And j <= 90) Then
serial = serial & Chr(j)
Else
serial = serial & Trim(Str(j))
End If
Next i
If Len(serial) > 30 Then serial = Left(serial, 30)
errhand:
If LCase(Trim(Err.Description)) = "overflow" Then
mat(i) = Int(mat(i) / 10)
Err.Clear
Resume
Else
If Err.Number <> 0 Then Err.Raise "90", , "Serial Number Generator Failure"
End If
get_serial = serial
End Function
Private Function clocker(valu As Currency) As Long
Do While valu > 257
valu = Int(valu / 10)
Loop
clocker = valu
End Function
