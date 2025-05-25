VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mWiersze As New Collection
Public obiekt As Object
Public Function AddWiersz(Optional Wysokoœæ As Single, Optional Nazwa As String) As Wiersz
Dim w As Wiersz

Set w = New Wiersz
w.wys = Wysokoœæ
If Nazwa = "" Then
    mWiersze.Add w
Else
    mWiersze.Add w, Nazwa
End If
Set AddWiersz = w
End Function
Public Sub Drukuj(X As Single, Y As Single)
Dim w As Wiersz, CY As Single
CY = Y
For Each w In mWiersze
    w.Drukuj X, CY, obiekt
    CY = CY + w.wys
Next w
    
End Sub
Public Function Wiersz(Index As Variant) As Wiersz
Set Wiersz = mWiersze.item(Index)
End Function
Public Function AddNewWiersz(Optional Wysokoœæ As Single, Optional Nazwa As String) As Wiersz
Dim w As Wiersz

Set w = New Wiersz
w.wys = Wysokoœæ
If Nazwa = "" Then
    mWiersze.Add w
Else
    mWiersze.Add w, Nazwa
End If
Set AddNewWiersz = w

End Function
Public Sub Clear()
Dim w As Wiersz
For Each w In mWiersze
    mWiersze.Remove (1)
Next w
End Sub

Private Sub Class_Terminate()
Set mWiersze = Nothing
Set w = Nothing
Set obiekt = Nothing
End Sub
