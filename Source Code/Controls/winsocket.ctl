VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl winsocket 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   ForwardFocus    =   -1  'True
   InvisibleAtRuntime=   -1  'True
   MaskPicture     =   "winsocket.ctx":0000
   ScaleHeight     =   465
   ScaleWidth      =   480
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1080
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   615
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "winsocket.ctx":27A2
      Top             =   -15
      Width           =   480
   End
End
Attribute VB_Name = "winsocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ostack() As String
Event ConnectionReqest(requestID As Long) 'MappingInfo=Winsock1,Winsock1,-1,ConnectionRequest
Event Error(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean) 'MappingInfo=Winsock1,Winsock1,-1,Error
Attribute Error.VB_Description = "Error occurred"
Event DataArrival(Command As String, data As String)
Private Sub Timer1_Timer()
On Error GoTo errhand
If UBound(ostack, 1) >= 0 Then
packet = ostack(0)
Winsock1.senddata packet
If UBound(ostack, 1) > 0 Then
For ij = 1 To UBound(ostack, 1)
ostack(ij - 1) = ostack(ij)
Next ij
ReDim Preserve ostack(UBound(ostack, 1) - 1)
Else
ostack(0) = ""
Timer1.Enabled = False
End If
End If
errhand:
If Err.Number = 40006 Then Err.Clear: Resume
End Sub
Sub senddata(ByVal Command As String, Optional ByVal data As String)
Dim lendata As String, lencom As String
On Error GoTo errhand
lendata = Trim(Str(Len(data)))
lencom = Trim(Str(Len(Trim(Command))))
Do While Len(lendata) < 3
lendata = "0" & Trim(lendata)
Loop
Do While Len(Trim(lencom)) < 2
lencom = "0" & Trim(lencom)
Loop
packet = lendata & lencom & Trim(UCase(Command)) & data
If Timer1.Enabled = True Then
ubo = UBound(ostack, 1) + 1
ReDim Preserve ostack(ubo)
Else
ubo = 0
End If
ostack(ubo) = packet
errhand:
If Err.Number = 9 Then Err.Clear: ubo = 0: Resume Next
Timer1.Enabled = True
End Sub
Private Function getcommand(data As String) As String
commandlen = Val(Mid(data, 4, 2))
getcommand = Mid(data, 6, commandlen)
End Function
Private Function getdata(data As String) As String
commandlen = Val(Mid(data, 4, 2))
datalen = Val(Left(data, 3))
getdata = Mid(data, 5 + commandlen + 1, datalen)
End Function
Private Sub UserControl_Initialize()
ReDim ostack(0)
End Sub
Private Sub UserControl_Resize()
UserControl.Height = 495
UserControl.Width = 510
End Sub
Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Static add As Boolean, predata As String
Winsock1.getdata data
If add = True Then data = predata & data: predata = "": add = False
commandlen = Val(Mid(data, 4, 2))
datalen = Val(Left(data, 3))
If Len(data) < (commandlen + datalen + 5) Then add = True: predata = data: Exit Sub
If Len(data) > (commandlen + datalen + 5) Then predata = _
Right(data, Len(data) - commandlen - datalen - 5): add = _
True: data = Left(data, commandlen + datalen + 5)
RaiseEvent DataArrival(getcommand(data), getdata(data))
End Sub
'Property Let send_ena(data As Boolean)
'Timer1.Enabled = data
'PropertyChanged "send_ena"
'End Property
Property Let throttle(throttl As Integer)
new_interval = Timer1.Interval - throttl
If new_interval > 100 Then Timer1.Interval = new_interval
PropertyChanged "throttle"
End Property
Property Get throttle() As Integer
thtottle = 400 - Timer1.Interval
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,Accept
Public Sub Accept(requestID As Long)
Attribute Accept.VB_Description = "Accept an incoming connection request"
   If Winsock1.State <> sckClosed Then _
    Winsock1.Close
   Winsock1.Accept requestID
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,Close
Public Sub Cloze()
    Winsock1.Close
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,Connect
Public Sub Connect(Optional RemoteHost As Variant, Optional RemotePort As Variant)
Attribute Connect.VB_Description = "Connect to the remote computer"
    If IsEmpty(RemoteHost) And IsEmpty(RemotePort) Then
    Winsock1.Connect
    Else
    Winsock1.Connect RemoteHost, RemotePort
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,ConnectionRequest
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    RaiseEvent ConnectionReqest(requestID)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,error
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Private Sub Winsock1_Error(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,Listen
Public Sub Listen()
Attribute Listen.VB_Description = "Listen for incoming connection requests"
    Winsock1.Listen
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,LocalHostName
Public Property Get LocalHostName() As String
Attribute LocalHostName.VB_Description = "Returns the local machine name"
    LocalHostName = Winsock1.LocalHostName
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,LocalIP
Public Property Get LocalIP() As String
Attribute LocalIP.VB_Description = "Returns the local machine IP address"
    LocalIP = Winsock1.LocalIP
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,LocalPort
Public Property Get LocalPort() As Long
Attribute LocalPort.VB_Description = "Returns/Sets the port used on the local computer"
    LocalPort = Winsock1.LocalPort
End Property

Public Property Let LocalPort(ByVal New_LocalPort As Long)
    Winsock1.LocalPort() = New_LocalPort
    PropertyChanged "LocalPort"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,RemoteHost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
    RemoteHost = Winsock1.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    Winsock1.RemoteHost() = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,RemoteHostIP
Public Property Get RemoteHostIP() As String
Attribute RemoteHostIP.VB_Description = "Returns the remote host IP address"
    RemoteHostIP = Winsock1.RemoteHostIP
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,RemotePort
Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "Returns/Sets the port to be connected to on the remote computer"
    RemotePort = Winsock1.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    Winsock1.RemotePort() = New_RemotePort
    PropertyChanged "RemotePort"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,State
Public Property Get State() As Integer
Attribute State.VB_Description = "Returns the state of the socket connection"
    State = Winsock1.State
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Winsock1.LocalPort = PropBag.ReadProperty("LocalPort", 0)
    Winsock1.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    Winsock1.RemotePort = PropBag.ReadProperty("RemotePort", 0)
    throttle = PropBag.ReadProperty("throttle", 0)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("LocalPort", Winsock1.LocalPort, 0)
    Call PropBag.WriteProperty("RemoteHost", Winsock1.RemoteHost, "")
    Call PropBag.WriteProperty("RemotePort", Winsock1.RemotePort, 0)
    Call PropBag.WriteProperty("throttle", m_throttle, 0)
End Sub


