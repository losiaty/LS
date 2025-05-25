VERSION 5.00
Begin VB.Form Historie 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Historie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lista As Tabela, w As Wiersz, e As Element
Private Sub Form_Activate()

Me.MousePointer = 0

End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set lista = New Tabela
Set lista.obiekt = Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set lista = Nothing

End Sub
