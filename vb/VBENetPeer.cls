VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBENetPeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Tag As Variant
Private m_hosHost As VBENetHost
Private m_lngPeer As Long
Private m_perPeer As ENetPeer
Private m_adrAddress As VBENetAddress

Friend Sub Free()
On Error Resume Next
    Set m_hosHost = Nothing
    m_lngPeer = 0
End Sub

Public Sub Refresh()
On Error Resume Next
    If m_lngPeer = 0 Then Exit Sub
    CopyMemory m_perPeer, ByVal m_lngPeer, Len(m_perPeer)
End Sub

Friend Sub Create(ByVal Host As VBENetHost, ByVal Peer As Long)
On Error Resume Next
    If Peer = 0 Then Exit Sub
    Set m_hosHost = Host
    m_lngPeer = Peer
    Refresh
End Sub

Public Function Send(ByVal Packet As VBENetPacket, Optional ByVal Channel As Long = 0) As Boolean
On Error Resume Next
    If m_lngPeer = 0 Then Exit Function
    Debug.Print "Sending packet..."
    If enet_peer_send(m_lngPeer, Channel, Packet.Handle) = 0 Then
        Packet.Owner = False
        Send = True
    End If
End Function

Public Function SendText(ByRef Data As String, Optional ByVal Channel As Long = 0, Optional Reliable As Boolean = True) As Boolean
On Error Resume Next
Dim l_bytData() As Byte
    l_bytData = Data
    SendText = SendBytes(l_bytData, Channel, Reliable)
End Function

Public Function SendBytes(ByRef Data() As Byte, Optional ByVal Channel As Long = 0, Optional Reliable As Boolean = True) As Boolean
On Error Resume Next
Dim l_lngPacket As Long, l_lngLength As Long
    If m_lngPeer = 0 Then Exit Function
    l_lngLength = UBound(Data) - LBound(Data) + 1
    If l_lngLength < 1 Then Exit Function
    Debug.Print "Creating temporary packet..."
    l_lngPacket = enet_packet_create(Data(LBound(Data)), l_lngLength, IIf(Reliable, 1, 0))
    Debug.Print "Sending temporary packet..."
    If enet_peer_send(m_lngPeer, Channel, l_lngPacket) = 0 Then
        SendBytes = True
    Else
        enet_packet_destroy l_lngPacket
    End If
End Function

Public Sub Disconnect()
On Error Resume Next
    If m_lngPeer = 0 Then Exit Sub
    enet_peer_disconnect m_lngPeer
End Sub

Public Sub ForceDisconnect()
On Error Resume Next
    If m_lngPeer = 0 Then Exit Sub
    enet_peer_reset m_lngPeer
End Sub

Public Sub Ping()
On Error Resume Next
    If m_lngPeer = 0 Then Exit Sub
    enet_peer_ping m_lngPeer
End Sub

Friend Property Get PrivateData() As ENetPeer
On Error Resume Next
    PrivateData = m_perPeer
End Property

Public Property Get Host() As VBENetHost
On Error Resume Next
    Set Host = m_hosHost
End Property

Public Property Get Handle() As Long
On Error Resume Next
    Handle = m_lngPeer
End Property

Public Property Get Address() As VBENetAddress
On Error Resume Next
    If m_adrAddress Is Nothing Then
        Set m_adrAddress = New VBENetAddress
        m_adrAddress.Address = m_perPeer.Address
    End If
    Set Address = m_adrAddress
End Property

Public Property Get ChannelCount() As Long
On Error Resume Next
    ChannelCount = m_perPeer.ChannelCount
End Property

Public Property Get IncomingBandwidth() As Long
On Error Resume Next
    IncomingBandwidth = m_perPeer.IncomingBandwidth
End Property

Public Property Get OutgoingBandwidth() As Long
On Error Resume Next
    OutgoingBandwidth = m_perPeer.OutgoingBandwidth
End Property

Public Property Get IncomingDataTotal() As Long
On Error Resume Next
    IncomingDataTotal = m_perPeer.IncomingDataTotal
End Property

Public Property Get OutgoingDataTotal() As Long
On Error Resume Next
    OutgoingDataTotal = m_perPeer.OutgoingDataTotal
End Property

Public Property Get PacketLoss() As Single
On Error Resume Next
    PacketLoss = CSng(m_perPeer.PacketLoss) / CSng(ENET_PEER_PACKET_LOSS_SCALE)
End Property

Public Property Get Latency() As Long
On Error Resume Next
    Latency = m_perPeer.RoundTripTime
End Property
