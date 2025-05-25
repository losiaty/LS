VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sciezka As String, Nazwa As String, FSO As FileSystemObject
Public Ready As Boolean, Initialized As Boolean
Private Tabela() As String, IleRek As Long, RekL As Long, NPliku As Long
Private Dodaj() As String

Public Function Inicjuj(SciezkaT As String, NazwaT As String) As Boolean
Dim Plik As TextStream, Wyraz As String, Spli As String, Czas As Single
On Error GoTo BladInicjuj
ReDim Preserve Tabela(0)
Sciezka = SciezkaT
Nazwa = NazwaT
Set Plik = FSO.OpenTextFile(Sciezka & "\" & Nazwa & ".bmt", ForReading, True)
IleRek = 0
ReDim Tabela(500000)

While Plik.AtEndOfStream = False
   Wyraz = Plik.ReadLine
   IleRek = IleRek + 1
   'ReDim Preserve Tabela(IleRek)
   Tabela(IleRek) = Wyraz
Wend
ReDim Preserve Tabela(IleRek)
Inicjuj = True
Ready = True
Plik.Close
'Czas = Timer
 '  Spli = Join(Tabela, ",")
'MsgBox Timer - Czas
'Spli = ""
Initialized = True
Exit Function
BladInicjuj:
Inicjuj = False
Plik.Close

End Function

Public Property Get Pole(Numer As Long) As String
On Error GoTo BladPole
Pole = Tabela(Numer)
Exit Property
BladPole:
Pole = Null
End Property

Public Property Get RecordCount() As Long
RecordCount = IleRek
End Property

Public Function GetRekord(Numer As Long) As String
Dim swyraz As String
If EOF(NPliku) = False Then
   Get NPliku, Numer, swyraz
   GetRekord = swyraz
Else
   GetRekord = Null
End If
End Function
Private Sub QuickSortTekst(tbl() As String, X As Long, Y As Long, Optional Descending As Boolean = False)
Dim i As Long, j As Long, v As String, Temp As String

If Descending = False Then
   i = X
   j = Y
   v = tbl((X + Y) \ 2)
   While i <= j
      While StrComp(v, tbl(i), vbBinaryCompare) = 1
         i = i + 1
      Wend
      
      While StrComp(v, tbl(j), vbBinaryCompare) = -1
         j = j - 1
      Wend
      If i <= j Then
         Temp = tbl(i)
         tbl(i) = tbl(j)
         tbl(j) = Temp
         i = i + 1
         j = j - 1
      End If
   Wend

   If X < j Then QuickSortTekst tbl, X, j
   If i < Y Then QuickSortTekst tbl, i, Y
Else
   i = X
   j = Y
   v = tbl((X + Y) \ 2)
   While i <= j
      While StrComp(v, tbl(i), vbBinaryCompare) = -1
         i = i + 1
      Wend
      
      While StrComp(v, tbl(j), vbBinaryCompare) = 1
         j = j - 1
      Wend
      If i <= j Then
         Temp = tbl(i)
         tbl(i) = tbl(j)
         tbl(j) = Temp
         i = i + 1
         j = j - 1
      End If
   Wend
   If X < j Then QuickSortTekst tbl, X, j, True
   If i < Y Then QuickSortTekst tbl, i, Y, True
End If
End Sub

Public Sub Add(Wyraz As String)
If FindB2(Wyraz) > 0 Then Exit Sub
IleRek = IleRek + 1
ReDim Preserve Tabela(IleRek)
Tabela(IleRek) = Wyraz
QuickSortTekst Tabela, 1, IleRek
Aktualizuj
End Sub

Public Sub AddNew(Wyraz As String)
Dim i As Long, n As Long, Temp() As String, j As Long
If FindB(Wyraz) > 0 Then Exit Sub
IleRek = IleRek + 1
n = FindPrev(Wyraz)
ReDim Temp(IleRek)
For i = n + 1 To IleRek - 1
   Temp(i - n) = Tabela(i)
Next i
Tabela(n + 1) = Wyraz
ReDim Preserve Tabela(IleRek)
For i = n + 2 To IleRek
   Tabela(i) = Temp(i - n - 1)
Next i
Aktualizuj
End Sub


Public Sub Remove(Wyraz As String)
Dim i As Long
i = FindB(Wyraz)
If i > 0 Then
   Tabela(i) = Chr$(255)
   QuickSortTekst Tabela, 1, IleRek
   IleRek = IleRek - 1
   ReDim Preserve Tabela(IleRek)
   Aktualizuj
End If

End Sub

Public Function Find(Wyraz As String) As Long
Dim i As Long

For i = UBound(Tabela) To 1 Step -1
   If Tabela(i) = Wyraz Then
      Find = i
      Exit Function
   End If
Next i
Find = 0
End Function

Public Sub Aktualizuj()
Dim Plik As TextStream, i As Long, F1 As String, F2 As String
QuickSortTekst Tabela, 1, IleRek
F1 = Sciezka & "\" & Nazwa & "tmp.bmt"
F2 = Sciezka & "\" & Nazwa & ".bmt"
Set Plik = FSO.OpenTextFile(F1, ForWriting, True)
For i = 1 To IleRek
   Plik.WriteLine Tabela(i)
Next i
Plik.Close
FSO.DeleteFile F2, True
FSO.MoveFile F1, F2
CreateIndeks
Set Plik = Nothing
End Sub

Public Function FindB(Wyraz As String) As Long
Dim w1 As Long, w2 As Long, Temp As String, w As Long
w1 = 1
w2 = IleRek
w = (w1 + w2) \ 2

While Abs(w1 - w2) > 1
   If StrComp(Wyraz, Tabela(w), vbBinaryCompare) > 0 Then
      w1 = w
   Else
      w2 = w
   End If
   w = (w1 + w2) \ 2
   If StrComp(Wyraz, Tabela(w), vbBinaryCompare) = 0 Then
      FindB = w
      Exit Function
   End If
Wend

If StrComp(Wyraz, Tabela(w1), vbBinaryCompare) = 0 Then
   FindB = w1
   Exit Function
End If
If StrComp(Wyraz, Tabela(w2), vbBinaryCompare) = 0 Then
   FindB = w2
   Exit Function
End If

FindB = 0
   
End Function

Public Sub CreateIndeks()
Dim Plik As TextStream, plik2 As TextStream, a As String * 1, i As Long, n As Long
Dim znak As String * 1, Pocz(40) As Long, Znaczek(40) As String * 1, j As Long
Dim Wyrazik As String
Set Plik = FSO.OpenTextFile(Sciezka & "\" & Nazwa & ".bmt", ForReading, False)
Set plik2 = FSO.OpenTextFile(Sciezka & "\" & Nazwa & ".imt", ForWriting, True)

While Plik.AtEndOfStream = False
   Wyrazik = Plik.ReadLine
   a = Left(Wyrazik, 1)
   n = n + 1
   If znak <> a Then
      i = i + 1
      Pocz(i) = n
      Znaczek(i) = a
      znak = a
   End If
Wend

Pocz(i + 1) = n + 1
For j = 1 To i
   plik2.WriteLine Znaczek(j) & "," & Pocz(j) & "," & Pocz(j + 1) - 1
Next j

plik2.Close
Plik.Close
Set plik2 = Nothing
Set Plik = Nothing
End Sub

Public Function FindPrev(Wyraz As String) As Long
Dim w1 As Long, w2 As Long, Temp As String, w As Long


If UBound(Tabela) < 1 Then
   FindPrev = 0
   Exit Function
End If

If Tabela(1) = Wyraz Then
   FindPrev = 0
   Exit Function
End If

w1 = 1
w2 = IleRek
w = (w1 + w2) \ 2

While Abs(w1 - w2) > 1
   If StrComp(Wyraz, Tabela(w), vbBinaryCompare) > 0 Then
      w1 = w
   Else
      w2 = w
   End If
   w = (w1 + w2) \ 2
Wend

FindPrev = w1

End Function

Public Function FindB2(Wyraz As String) As Long
Dim i As Long

i = FindPrev(Wyraz)
If IleRek > i Then
   If Tabela(i + 1) = Wyraz Then
      FindB2 = i + 1
   Else
      FindB2 = 0
   End If
Else
  FindB2 = 0
End If

End Function

Public Sub DodajWyraz(Wyraz As String)
Dim i As Long
   i = UBound(Dodaj) + 1
   ReDim Preserve Dodaj(i)
   Dodaj(i) = Wyraz
End Sub

Public Sub Dodawaj()
Dim i As Long, IleDodaj As Long, OldIleRek As Long
IleDodaj = UBound(Dodaj)
OldIleRek = IleRek
IleRek = IleRek + IleDodaj
ReDim Preserve Tabela(IleRek)
For i = 1 To IleDodaj
   Tabela(OldIleRek + i) = Dodaj(i)
Next i
Aktualizuj
ReDim Dodaj(0)
End Sub

Private Sub Class_Initialize()
ReDim Dodaj(0)
End Sub
