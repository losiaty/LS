VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   Caption         =   "Lów S³ow Serwer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock m_sckListener 
      Left            =   480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView m_KlientsGrid 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP Address"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Port"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "State"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSWinsockLib.Winsock m_Gniazdo 
      Index           =   0
      Left            =   4080
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim i As Long
Me.Caption = "£ów S³ów Server"

m_sckListener.LocalPort = 2595
m_sckListener.Listen

m_Gniazdo(0).LocalPort = 2596

For i = 1 To 3
    Load m_Gniazdo(i)
    m_Gniazdo(i).LocalPort = 2596 + i
Next i

FillGrid

End Sub

Private Sub m_btnPlayer_Click(Index As Integer)

End Sub

Public Sub FillGrid()
Dim i As Long
Dim lItem As ListItem

m_KlientsGrid.ListItems.Clear

For i = 0 To 3
    Set lItem = m_KlientsGrid.ListItems.Add(, , i + 1)
    
    'If m_Gniazdo(i).State = 7 Then

        lItem.ListSubItems.Add , , m_Gniazdo(i).State
        lItem.ListSubItems.Add , , m_Gniazdo(i).RemoteHostIP
        lItem.ListSubItems.Add , , m_Gniazdo(i).RemotePort
        lItem.ListSubItems.Add , , GetStatusText(m_Gniazdo(i))
    
Next i

End Sub

Private Sub m_Gniazdo_Close(Index As Integer)
FillGrid
End Sub

Private Sub m_Gniazdo_Connect(Index As Integer)
FillGrid
End Sub

Private Sub m_Gniazdo_ConnectionRequest(Index As Integer, ByVal requestID As Long)
FillGrid
End Sub

Private Sub m_Gniazdo_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim vData As Variant
Dim pIP() As String
Dim strData As String
Dim k As wKierunek
Dim Pkt As Long
m_Gniazdo(Index).GetData vData, vbString
pIP = Split(vData, "|")

If UBound(pIP) < 5 Then Exit Sub
If CLng(pIP(5)) <> Ktory Then Exit Sub

If pIP(0) = "Kladzie" Then
    
    If pIP(2) = "1" Then
        k = wPionowy
    Else
        k = wPoziomy
    End If
    graslow.svrKladzie pIP(1), CLng(pIP(3)), k, pIP(4)

ElseIf pIP(0) = "Litery" Then
    
    'graslow.UpdateStojak pIP(1)
    strData = "Litery|" & graslow.svrGetNewStojak(pIP(1))
    m_Gniazdo(Index).SendData strData

End If

FillGrid

End Sub

Private Sub m_Gniazdo_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
FillGrid
End Sub

Private Sub m_Gniazdo_SendComplete(Index As Integer)
FillGrid
End Sub

Private Sub m_Gniazdo_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
FillGrid
End Sub

Private Sub m_sckListener_Close()
FillGrid
End Sub

Private Sub m_sckListener_Connect()
FillGrid
End Sub

Private Sub m_sckListener_ConnectionRequest(ByVal requestID As Long)
Dim i As Long

For i = 0 To 3
    If m_Gniazdo(i).State <> 7 Then
        m_Gniazdo(i).Close
        m_Gniazdo(i).Accept requestID
        graslow.svrSetPlayerPresent i, True
        FillGrid
        Me.Show
        Exit Sub
    End If
Next i

FillGrid

End Sub

Private Sub m_sckListener_DataArrival(ByVal bytesTotal As Long)
FillGrid
End Sub

Private Sub m_sckListener_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
FillGrid
End Sub

Private Sub m_sckListener_SendComplete()
FillGrid
End Sub

Private Sub m_sckListener_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
FillGrid
End Sub

Public Function GetStatusText(pSocket As Winsock)

Dim X
X = pSocket.State

Select Case X
Case 0
    GetStatusText = "Off line"
Case 1
    GetStatusText = "On line"
Case 2
    GetStatusText = "Nas³uchiwanie"
Case 3
    GetStatusText = "Przygotowanie po³¹czenia"
Case 4
    GetStatusText = "Znajdowanie hosta"
Case 5
    GetStatusText = "HOST znaleziono"
Case 6
    GetStatusText = "Nawi¹zywanie po³¹czenia"
Case 7
    GetStatusText = "Po³¹czenie nawi¹zano"
Case 8
    GetStatusText = "Po³¹czenie jest zamykane"
Case Else
    GetStatusText = "???"

End Select

End Function
