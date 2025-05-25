VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabela2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Linie As Collection
Public Cel As Object, MojaStrona As cStrona2

Public Property Get Height() As Single
Dim L As Clinia2, H As Single
For Each L In Linie
    H = H + L.Height + 1
Next L
Height = H
End Property

Public Function AddLine(Optional MinWysokosc As Single = 0, Optional Szerokosc As Single = 0, Optional Podkresl As Boolean, Optional Nadkresl As Boolean) As Clinia2
Dim NL As New Clinia2

Set NL.Cel = Cel
Set NL.MojaTabela = Me
NL.MinHeight = MinWysokosc
NL.Podkresl = Podkresl
NL.Nadkresl = Nadkresl
Linie.Add NL
Set AddLine = NL

End Function
Public Sub Drukuj(ByVal X As Single, ByVal Y As Single)

Dim L As Clinia2, py As Long, px As Long
py = Y
px = X
For Each L In Linie
   L.Drukuj px, py
   py = py + L.Height + 1
Next L

End Sub

Private Sub Class_Initialize()
Set Linie = New Collection
End Sub
