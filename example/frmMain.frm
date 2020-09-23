VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBENet Chat"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPoll 
      Interval        =   1
      Left            =   3180
      Top             =   2940
   End
   Begin VB.ListBox lstUsers 
      Height          =   5955
      IntegralHeight  =   0   'False
      Left            =   6030
      TabIndex        =   2
      Top             =   30
      Width           =   1515
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   1
      Top             =   6015
      Width           =   7515
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5955
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   5970
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnectToServer 
         Caption         =   "&Connect to Server..."
      End
      Begin VB.Menu mnuStartServer 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu sepFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ServerPort As Long = 1234

Public WithEvents Client As VBENetHost
Attribute Client.VB_VarHelpID = -1
Public WithEvents Server As VBENetHost
Attribute Server.VB_VarHelpID = -1
Public ServerConnection As VBENetPeer

Private Sub DoCommand(Command As String)
On Error Resume Next
Dim l_varParts
    If Left(Command, 1) = "/" Then
        l_varParts = Split(Mid(Command, 2), " ")
        Select Case LCase(Trim(l_varParts(0)))
        Case "nick"
            ' Change nick
            ServerConnection.SendText "nick:" & l_varParts(1)
        Case "kick"
            ' Attempt to kick a user
            If Server Is Nothing Then
                LogPrint "*** You must be running a server to kick users"
            Else
                KickUser CStr(l_varParts(1))
            End If
        End Select
    Else
        ' Send message
        ServerConnection.SendText "msg:" & Command
    End If
End Sub

Private Sub KickUser(Nick As String)
On Error Resume Next
Dim l_perPeer As VBENetPeer
    For Each l_perPeer In Server.Peers
        If Trim(LCase(Nick)) = LCase(Trim(l_perPeer.Tag)) Then
            ' Send a disconnect message
            l_perPeer.SendText "msg:*** You were kicked"
            ' If we don't manually poll the server before disconnecting the user, they will never recieve the message we just sent
            Server.Poll
            ' Terminate their connection
            l_perPeer.Disconnect
            Exit For
        End If
    Next l_perPeer
End Sub

Private Function AutoSelectNick() As String
On Error Resume Next
Dim l_strNick As String, l_lngIndex As Long
Dim l_perPeer As VBENetPeer, l_booInUse As Boolean
    l_lngIndex = 1
    l_strNick = "ChatUser"
    l_booInUse = True
    Do While l_booInUse
        l_booInUse = False
        For Each l_perPeer In Server.Peers
            If Trim(LCase(l_strNick)) = LCase(Trim(l_perPeer.Tag)) Then
                l_booInUse = True
                l_lngIndex = l_lngIndex + 1
                l_strNick = "ChatUser" & l_lngIndex
                Exit For
            End If
        Next l_perPeer
    Loop
    AutoSelectNick = l_strNick
End Function

Public Sub LogPrint(ByRef Text As String)
On Error Resume Next
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelText = Text & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub Client_Connect(ByVal Peer As VBENetActiveX.VBENetPeer)
On Error Resume Next
    LogPrint "*** Connected to " & Peer.Address.Text
    Peer.SendText "join:" & InputBox("Choose a nick", "Connecting to server", "ChatUser")
End Sub

Private Sub Client_Disconnect(ByVal Peer As VBENetActiveX.VBENetPeer)
On Error Resume Next
    LogPrint "*** Disconnected"
    ' Wipe the user list
    lstUsers.Clear
End Sub

Private Sub Client_Recieve(ByVal Peer As VBENetActiveX.VBENetPeer, ByVal Channel As Long, ByVal Packet As VBENetActiveX.VBENetPacket)
On Error Resume Next
Dim l_strPacket As String
Dim l_strType As String, l_strData As String
Dim l_varParts As Variant
Dim l_lngIndex As Long
    l_strPacket = Packet.Text
    l_strType = Left(l_strPacket, InStr(l_strPacket, ":") - 1)
    l_strData = Mid(l_strPacket, InStr(l_strPacket, ":") + 1)
    ' Determine packet type
    Select Case LCase(Trim(l_strType))
    Case "join"
        ' Join packet recieved; add user to user list
        lstUsers.AddItem l_strData
        LogPrint "*** " & l_strData & " has joined"
    Case "quit"
        ' Quit packet recieved; if user is in user list, then remove them
        For l_lngIndex = 0 To lstUsers.ListCount - 1
            If Trim(LCase(lstUsers.List(l_lngIndex))) = Trim(LCase(l_strData)) Then
                lstUsers.RemoveItem l_lngIndex
                Exit For
            End If
        Next l_lngIndex
        LogPrint "*** " & l_strData & " has quit"
    Case "nick"
        ' Nick change packet recieved; if user is in user list, change their entry to their new name
        l_varParts = Split(l_strData, ",")
        For l_lngIndex = 0 To lstUsers.ListCount - 1
            If Trim(LCase(lstUsers.List(l_lngIndex))) = Trim(LCase(l_varParts(0))) Then
                lstUsers.List(l_lngIndex) = l_varParts(1)
                Exit For
            End If
        Next l_lngIndex
        LogPrint "*** " & l_varParts(0) & " changed nick to " & l_varParts(1)
    Case "msg"
        ' Message packet recieved; print to log textbox
        LogPrint l_strData
    End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
    LogPrint "*** Welcome to Chat"
End Sub

Private Sub mnuConnectToServer_Click()
On Error Resume Next
    Set Client = CreateClientHost()
    Set ServerConnection = Client.Connect(CreateAddressFromString(InputBox("Enter server address", "Connect to server", "localhost:1234")))
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    End
End Sub

Private Sub mnuStartServer_Click()
On Error Resume Next
    ' Create a listen server
    Set Server = CreateServerHost(CreateAddress("", ServerPort))
    If Server Is Nothing Then
        LogPrint "*** Failed to start server"
    Else
        LogPrint "*** Server started on port " & ServerPort
    End If
End Sub

Private Sub Server_Connect(ByVal Peer As VBENetActiveX.VBENetPeer)
On Error Resume Next
    LogPrint "*** " & Peer.Address.Text & " connected to server"
    ' Set the peer's username to blank
    Peer.Tag = ""
End Sub

Private Sub Server_Disconnect(ByVal Peer As VBENetActiveX.VBENetPeer)
On Error Resume Next
    If Peer.Tag = "" Then
        ' This client never sent a join packet
        LogPrint "*** " & Peer.Address.Text & " disconnected from server"
    Else
        Server.BroadcastText "quit:" & Peer.Tag
        LogPrint "*** " & Peer.Tag & "(" & Peer.Address.Text & ") disconnected from server"
    End If
End Sub

Private Sub Server_Recieve(ByVal Peer As VBENetActiveX.VBENetPeer, ByVal Channel As Long, ByVal Packet As VBENetActiveX.VBENetPacket)
On Error Resume Next
Dim l_strPacket As String
Dim l_strType As String, l_strData As String
Dim l_varParts As Variant
Dim l_perPeer As VBENetPeer
    l_strPacket = Packet.Text
    l_strType = Left(l_strPacket, InStr(l_strPacket, ":") - 1)
    l_strData = Mid(l_strPacket, InStr(l_strPacket, ":") + 1)
    ' Determine packet type
    Select Case LCase(Trim(l_strType))
    Case "join"
        ' Join packet recieved; check for nick collision and then broadcast the packet
        For Each l_perPeer In Server.Peers
            If Trim(LCase(l_strData)) = LCase(Trim(l_perPeer.Tag)) Then
                Peer.SendText "msg:*** Nick " & l_strData & " already in use. Auto-assigning nick."
                l_strData = AutoSelectNick()
                Exit For
            End If
        Next l_perPeer
        Peer.Tag = l_strData
        Server.BroadcastText "join:" & l_strData
        For Each l_perPeer In Server.Peers
            If l_perPeer Is Peer Then
            Else
                Peer.SendText "join:" & l_perPeer.Tag
            End If
        Next l_perPeer
    Case "nick"
        ' Nick change packet recieved; check for nick collision and then broadcast the packet
        For Each l_perPeer In Server.Peers
            If Trim(LCase(l_strData)) = LCase(Trim(l_perPeer.Tag)) Then
                Peer.SendText "msg:*** Nick " & l_strData & " already in use"
                Exit Sub
            End If
        Next l_perPeer
        Server.BroadcastText "nick:" & Peer.Tag & "," & l_strData
        Peer.Tag = l_strData
    Case "msg"
        ' Message packet recieved; append username and broadcast
        Server.BroadcastText "msg:<" & Peer.Tag & "> " & l_strData
    End Select
End Sub

Private Sub tmrPoll_Timer()
On Error Resume Next
    ' ENet must be 'polled' for it to process network events, so it is usually best to do this in a very rapidly firing timer.
    Server.Poll
    Client.Poll
    Err.Clear
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        DoCommand txtInput.Text
        txtInput.Text = ""
        KeyAscii = 0
    End If
End Sub
