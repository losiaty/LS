VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWyrazy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type tWyrazy
   Wyraz As String
   Punkty As Long
   Start As Long
   Kierunek As wKierunek
   NPola As Long
   IleLiter As Long
   Wart As Long
   Dlugosc As Long
End Type
Private Elementy() As tWyrazy, IleRek As Long, NrRek As Long

Public Function Element(Numer As Long) As cFullWyraz
Dim e As cFullWyraz
If Numer > IleRek Then
   Set Element = Nothing
   Exit Function
End If
Set e = New cFullWyraz
e.IleLiter = Elementy(Numer).IleLiter
e.Kierunek = Elementy(Numer).Kierunek
e.NPola = Elementy(Numer).NPola
e.Punkty = Elementy(Numer).Punkty
e.Start = Elementy(Numer).Start
e.Wart = Elementy(Numer).Wart
e.Wyraz = Elementy(Numer).Wyraz
e.Dlugosc = Elementy(Numer).Dlugosc

Set Element = e
Set e = Nothing

End Function

Public Sub Add(Wyraz As String, IleLiter As Long, Kierunek As wKierunek, NPola As Long, Start As Long, Punkty As Long, Wart As Long)

IleRek = IleRek + 1
ReDim Preserve Elementy(IleRek)
Elementy(IleRek).Dlugosc = Len(Wyraz)
Elementy(IleRek).Wyraz = Wyraz
Elementy(IleRek).IleLiter = IleLiter
Elementy(IleRek).Kierunek = Kierunek
Elementy(IleRek).NPola = NPola
Elementy(IleRek).Punkty = Punkty
Elementy(IleRek).Start = Start
Elementy(IleRek).Wart = Wart

End Sub

Public Sub AddPlace(nSize As Long)

ReDim Preserve Elementy(IleRek + nSize)

End Sub

Public Sub SetValidPlace()

ReDim Preserve Elementy(IleRek)

End Sub


Public Sub AddWyraz(Wyraz As String, IleLiter As Long, Kierunek As wKierunek, NPola As Long, Start As Long, Punkty As Long, Wart As Long)
Dim t As tWyrazy

t.Dlugosc = IleLiter
t.IleLiter = IleLiter
t.Kierunek = Kierunek
t.NPola = NPola
t.Punkty = Punkty
t.Start = Start
t.Wart = Wart
IleRek = IleRek + 1
t.Wyraz = Wyraz
Elementy(IleRek) = t

End Sub
Public Sub AddTab(Wyrazy() As String, IleLiter As Long, Kierunek As wKierunek, NPola As Long, Start As Long, Punkty As Long, Wart As Long)
Dim i As Long, j As Long, t As tWyrazy
ReDim Preserve Elementy(IleRek + UBound(Wyrazy))
t.Dlugosc = IleLiter
t.IleLiter = IleLiter
t.Kierunek = Kierunek
t.NPola = NPola
t.Punkty = Punkty
t.Start = Start
t.Wart = Wart
For i = 1 To UBound(Wyrazy)
   j = IleRek + i
   t.Wyraz = Wyrazy(i)
   Elementy(j) = t
Next i
IleRek = IleRek + UBound(Wyrazy)
End Sub
Public Sub AddObject(FullWyraz As cFullWyraz)
IleRek = IleRek + 1
ReDim Preserve Elementy(IleRek)
Elementy(IleRek).Dlugosc = FullWyraz.Dlugosc
Elementy(IleRek).Wyraz = FullWyraz.Wyraz
Elementy(IleRek).IleLiter = FullWyraz.IleLiter
Elementy(IleRek).Kierunek = FullWyraz.Kierunek
Elementy(IleRek).NPola = FullWyraz.NPola
Elementy(IleRek).Punkty = FullWyraz.Punkty
Elementy(IleRek).Start = FullWyraz.Start
Elementy(IleRek).Wart = FullWyraz.Wart

End Sub

Public Sub Clear()
ReDim Elementy(0)
IleRek = 0
End Sub

Public Property Get EOF() As Boolean
If NrRek > IleRek Then
   EOF = True
Else
   EOF = False
End If
End Property
Public Property Get BOF() As Boolean
If NrRek < 1 Then
   BOF = True
Else
   BOF = False
End If
End Property
Public Property Get RecordCount() As Long
RecordCount = IleRek
End Property
Public Sub MoveNext()
If NrRek <= IleRek Then
   NrRek = NrRek + 1
End If
End Sub
Public Sub MoveFirst()
   NrRek = 1
End Sub
Public Sub MoveLast()
   NrRek = IleRek
End Sub

Public Sub MovePrevious()
If NrRek > 0 Then
   NrRek = NrRek - 1
End If
End Sub

