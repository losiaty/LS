VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmKlient 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_Port 
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Text            =   "2595"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton m_btnConnect 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock m_Gniazdo 
      Left            =   960
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmKlient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Klient As Object
Private Sub m_btnConnect_Click()

If m_Gniazdo.State <> 7 Then ' jesli nie ma po³aczenia
    m_Gniazdo.Close ' najpierw na wszelki wypek zamykamy
    m_Gniazdo.RemoteHost = "127.0.0.1" ' nazwa serwera ' moze byc ip
    m_Gniazdo.RemotePort = txt_Port ' numer portu
    m_Gniazdo.Connect ' ³aczymy siê
    Me.Hide
End If

End Sub

Private Sub m_Gniazdo_Connect()
Dim n As Long
n = 0
m_Klient.GetLetters "_0_0_0_0_0_0_0"
End Sub

Private Sub m_Gniazdo_DataArrival(ByVal bytesTotal As Long)
Dim vData As Variant
Dim strData() As String

m_Gniazdo.GetData vData, vbString

strData = Split(vData, "|")

If strData(0) = "Litery" Then
    graslow.UpdateStojak (strData(1))
    graslow.Podstawka.Visible = True

ElseIf strData(0) = "" Then


End If

End Sub

Private Sub m_Gniazdo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim n As Long
n = 0

End Sub

Private Sub m_Gniazdo_SendComplete()

Dim n As Long
n = 0
End Sub
