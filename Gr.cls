VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Score As Long, Pseudo As String, CR As Long, CC As Long
Public iw As Long, il As Integer, Komp As Boolean, MaxPsz As Long
Public IloscWymian As Long
Dim Wyrazik() As String
Public MyKolek As Collection

Public Property Get wynik() As Long
wynik = Score
End Property
Public Property Let wynik(ByVal Value As Long)
Score = Value
End Property
Public Property Get imie() As String
imie = Pseudo
End Property
Public Property Let imie(ByVal Value As String)
Pseudo = Value
End Property
Public Property Get czasRuchu() As Long
czasRuchu = CR
End Property
Public Property Get MinPKT() As Long
Select Case MaxPsz
   Case 0
      MinPKT = 55
   Case 1
      MinPKT = 60
   Case 2
      MinPKT = 62
   Case 8
      MinPKT = 65
End Select
End Property
Public Property Let czasRuchu(ByVal Value As Long)
CR = Value
End Property
Public Property Get CzasCa³kowity() As Long
CzasCa³kowity = CC
End Property

Public Property Let CzasCa³kowity(ByVal Value As Long)
CC = Value
End Property

Public Property Get IleWyrazów() As Long
IleWyrazów = MyKolek.Count
End Property
Public Function Œrednia() As Single
If il = 0 Then
   Œrednia = 0
   Exit Function
End If
Œrednia = iw / il
End Function

Public Sub DodajWyrazy(kol As Collection, Pekt As Long)

If kol Is Nothing Then Exit Sub
Dim Ruch As Ruchy, S As String, item As Variant
Set Ruch = New Ruchy
il = il + 1
For Each item In kol
    S = S & item & " + "
    If Left(CStr(item), 4) <> " ***" Then
      iw = iw + 1
   End If
Next item
If Len(S) > 3 Then
   Ruch.S³owa = Left(S, Len(S) - 3)
Else
   Ruch.S³owa = ""
End If
Ruch.Punkty = Pekt
MyKolek.Add Ruch
End Sub

Public Sub ClearWyrazy()
Dim item As Variant
il = 0
iw = 0
For Each item In MyKolek
    MyKolek.Remove (1)
Next item
End Sub

Public Property Get Wyraz() As Collection
Set Wyraz = MyKolek
End Property
Private Sub Class_Initialize()
Set MyKolek = New Collection
End Sub
Public Property Get IleRuchów() As Integer
IleRuchów = il
End Property

Public Property Let IleRuchów(val As Integer)
il = val
End Property

Private Sub Class_Terminate()
'Set Ruch = Nothing
Set MyKolek = Nothing
End Sub
