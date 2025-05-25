VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Wiersz
    Wyraz As String
    Punkty As String
End Type
Dim Item() As Wiersz, Licznik As Integer

Public Sub AddItem(Optional Slowo As String, Optional Punkty As String)
Licznik = Licznik + 1
ReDim Preserve Item(Licznik)
Item(Licznik).Wyraz = Slowo
Item(Licznik).Punkty = Punkty

End Sub

Public Sub Rysuj(Obiekt As Object, Optional FontName As String, Optional FontSize As Integer, Optional Kolor As ColorConstants)
Dim i As Integer, wWiersza As Single, px As Single, py As Single
Dim szer As Long, KT³a As ColorConstants
If FontName = "" Then FontName = "Arial"
If FontSize = 0 Then FontSize = 12
If Kolor = 0 Then
    KT³a = vbWhite
Else
    KT³a = Kolor
End If
Obiekt.Font.Name = FontName

Obiekt.Font.Size = FontSize
Obiekt.Font.Charset = 238
wWiersza = (Obiekt.TextHeight(Item(i).Wyraz))
py = 10
px = 50
szer = 4500

i = 1
    Obiekt.Line (px, py + i * wWiersza)-(px + szer, py + (i + 1) * wWiersza), vbYellow, BF
    Obiekt.CurrentX = px + 100
    Obiekt.CurrentY = py + i * wWiersza
    Obiekt.Print (Item(i).Wyraz)

For i = 2 To Licznik
    Obiekt.Line (px, py + i * wWiersza)-(px + szer, py + (i + 1) * wWiersza), KT³a, BF
    Obiekt.CurrentX = px + 100
    Obiekt.CurrentY = py + i * wWiersza
    Obiekt.Print (Item(i).Wyraz)
Next i
For i = 0 To Licznik
    Obiekt.CurrentX = px + 3800 - (Obiekt.TextWidth(Item(i).Punkty)) / 2
    Obiekt.CurrentY = py + i * wWiersza
    Obiekt.Print (Item(i).Punkty)
Next i

End Sub

Public Sub Clear()
ReDim Preserve Item(1)
Licznik = 1
End Sub

Public Sub Remove(Index As Integer)

If Index > Licznik Or Index < 1 Then
    Exit Sub
Else
    For i = Index To Licznik - 1
        Item(i).Wyraz = Item(i + 1).Wyraz
        Item(i).Punkty = Item(i + 1).Punkty
    Next i
End If
Licznik = Licznik - 1
ReDim Preserve Item(Licznik)
End Sub

Private Sub Class_Initialize()
ReDim Item(1)
Licznik = 1
Item(1).Punkty = "Punkty : "
Item(1).Wyraz = "Wyraz : "
End Sub
Public Property Get Count() As Long
Count = Licznik - 1
End Property

