Attribute VB_Name = "Startup"
Option Explicit
Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal _
hWndInsertAfter As Long, ByVal X As Long, ByVal Y As _
Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags _
As Long) As Long
Public myappa As Myapp
'------------------------------------------------------
'Author:Don (Donzo) Zieler
'Posted:5/26/98
'webmaster@dad.win.net      webmaster@cd-mall.com
'http://www.win.net/dad     http://www.cd-mall.com
'
'Here's some code to read a text box and start the
'default browser to go to the site. It works really well
'in a phone book program I wrote! I wrote this in VB 5.0
'Let me know what you think.
'------------------------------------------------------
'OBJECTS to create for this code:
'text box named txtWeb  with the text set to http://www.
'button named cmdWeb
 
' Place this in a BAS module
Dim success As Integer
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Here's the function code
Function ShellToBrowser%(Frm As Form, ByVal URL$, ByVal WindowStyle%)
    
    Dim api%
        api% = ShellExecute(Frm.hwnd, "open", URL$, "", App.Path, WindowStyle%)
 
    'Check return value
    If api% < 31 Then
        'error code - see api help for more info
        MsgBox App.Title & " had a problem running your web browser." & _
          "You should check that your browser is correctly installed." & _
          ("Error" & Format$(api%)), 48, "Browser Unavailable"
        ShellToBrowser% = False
    ElseIf api% = 32 Then
        'no file association
        MsgBox App.Title & " could not find a file association for " & _
          URL$ & " on your system. You should check that your browser" & _
          "is correctly installed and associated with this type of file.", 48, "Browser Unavailable"
        ShellToBrowser% = False
    Else
        'It worked!
        ShellToBrowser% = True
 
    End If
    
End Function
 
Sub Main()
Set myappa = New Myapp
End Sub
