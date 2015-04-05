Attribute VB_Name = "windows"
Option Explicit
Public Const MAX_PATH = 260
Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long
Public Function GetWinPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetWinPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetWinPath = ""
End If
End Function
Public Function GetSystemPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetSystemPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetSystemPath = ""
End If
End Function



