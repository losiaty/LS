VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBazaDanych"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type tIndeks
   First As Long
   Last As Long
End Type
Private Type tWyrazy
   Wyraz As String
   Punkty As Long
   Start As Long
   Kierunek As wKierunek
   NPola As Long
   IleLiter As Long
   Wart As Long
End Type
'Private Type tPlus1
 '  Wyraz As String
  ' Dlugosc As Long
'End Type
Private Type tDoDo
   Wyraz As String
   NumerPola As Long
   Kierunek As wKierunek
End Type
Private IleRek As Long, AktRek As Long, IleWlasne As Long
Private FSO As FileSystemObject, IleDoDo As Long
Public Sciezka As String
Private Tabele As Collection
Private Indeksy As Collection
Private IleRekPlus12 As Long, IleRekPlus13 As Long, IleRekPlus14 As Long, IleRekPlus15 As Long, IleRekPlus16 As Long, IleRekPlus17 As Long, IleRekPlus18 As Long
Public Wyrazy As cWyrazy, AllWyrazy As cWyrazy, AllWyrazySort As cWyrazy
Private Wlasne() As String, DoDo() As tDoDo
Private Plus12() As String, Plus13() As String, Plus14() As String, Plus15() As String, Plus16() As String, Plus17() As String, Plus18() As String
Private First As Long, Last As Long, Rekordy() As String
Private Declare Function Lajk Lib "LowSlow.dll" (ByVal Wyraz As String, ByVal Warunek As String, ByVal Start As Long, ByVal Koniec As Long) As Long
Private Declare Function Zawiera Lib "LowSlow.dll" (ByVal Wyraz As String, ByVal LWyraz As Long, ByVal Warunek As String, ByVal LWarunek As Long) As Long
Public Property Get RecordCount() As Long
RecordCount = IleRek
End Property
Public Sub MoveNext()
If AktRek <= IleRek Then
   AktRek = AktRek + 1
End If
End Sub
Public Sub MoveFirst()
   AktRek = 1
End Sub
Public Sub MoveLast()
   AktRek = IleRek
End Sub

Public Sub MovePrevious()
If AktRek > 0 Then
   AktRek = AktRek - 1
End If
End Sub
Public Property Get EOF() As Boolean
'If AktRek > IleRek Then
'   EOF = True
'Else
'   EOF = False
'End If
EOF = IIf(AktRek > IleRek, True, False)
End Property
Public Property Get BOF() As Boolean
'If AktRek = 0 Then
'   BOF = True
'Else
'   BOF = False
'End If
BOF = IIf(AktRek = 0, True, False)
End Property
Public Property Get NoMatch() As Boolean
'If IleRek = 0 Then
'   NoMatch = True
'Else
'   NoMatch = False
'End If
NoMatch = IIf(IleRek = 0, True, False)
End Property

Private Sub Class_Initialize()
Dim i As Long
Set FSO = New FileSystemObject
End Sub
Private Sub Class_Terminate()
Set Tabele = Nothing
Set Indeksy = Nothing
Set FSO = Nothing
Set Wyrazy = Nothing
Set AllWyrazy = Nothing
Set AllWyrazySort = Nothing
End Sub
Private Sub Indeksuj(Tabela As String, znak As String)
Dim Plik42 As TextStream, a As String, i As Long, n As Long, spl() As String
Set Plik42 = FSO.OpenTextFile(Sciezka & "\" & Tabela & ".imt", ForReading, False)
While Plik42.AtEndOfStream = False
   a = Plik42.ReadLine
   spl = Split(a, ",")
   If spl(0) = znak Then
      First = spl(1)
      Last = spl(2)
      Plik42.Close
      Set Plik42 = Nothing
      Exit Sub
   End If
Wend
Plik42.Close
Set Plik42 = Nothing
End Sub

Public Property Get Rekord(Optional Numer As Long = 0) As String
If Numer < 0 Or Numer > IleRek Then
   Rekord = Null
   Exit Property
End If
If Numer = 0 Then
   Rekord = Rekordy(AktRek)
Else
   Rekord = Rekordy(Numer)
End If
   
End Function

Private Function StartP(Warunek As String) As Long
Dim i As Long
For i = 1 To Len(Warunek)
   If Mid(Warunek, i, 1) <> "_" Then
      StartP = i
      Exit Function
   End If
Next i

StartP = -1

End Function

Private Function Koniec(Warunek As String) As Long
Dim i As Long
For i = Len(Warunek) To 1 Step -1
   If Mid(Warunek, i, 1) <> "_" Then
      Koniec = i
      Exit Function
   End If
Next i

Koniec = -1

End Function

Public Sub MoveToRecord(Numer As Long)

AktRek = Numer

End Sub

Public Function Inicjuj(Katalog As String) As Boolean

Dim tabel As cTabela, i As Long, czasik As Single, indek As cIndeks
Sciezka = Katalog
czasik = Timer
Set Tabele = New Collection
Set Indeksy = New Collection
Set AllWyrazy = New cWyrazy
Set Wyrazy = New cWyrazy

For i = 15 To 2 Step -1
   Set tabel = New cTabela
   Set tabel.FSO = FSO
   tabel.Inicjuj Sciezka, "Sl" & CStr(i)
   Tabele.Add tabel, tabel.Nazwa
   Set tabel = Nothing
   Set indek = New cIndeks
   Set indek.FSO = FSO
   indek.Inicjuj "Sl" & CStr(i), Sciezka
   Indeksy.Add indek, "Sl" & CStr(i)
   Set indek = Nothing
   
   'MsgBox CStr(i)
Next i
End Function
Private Function CzyLajk(Warunek As String) As Boolean
Dim i As Long
For i = 1 To Len(Warunek)
   If Mid(Warunek, i, 1) <> "_" Then
      CzyLajk = True
      Exit Function
   End If
Next i
CzyLajk = False
End Function

Public Sub Add(Wyraz As String)
Dim Tabela As String
If Len(Wyraz) < 2 Or Len(Wyraz) > 15 Then Exit Sub
Tabela = "Sl" & CStr(Len(Wyraz))
Tabele(Tabela).Add UCase(Wyraz)
CreateIndeks (Tabela)
End Sub

Public Sub Remove(Wyraz As String)
Dim Tabela As String
If Len(Wyraz) < 2 Or Len(Wyraz) > 15 Then Exit Sub
Tabela = "Sl" & CStr(Len(Wyraz))
Tabele(Tabela).Remove UCase(Wyraz)
CreateIndeks (Tabela)
End Sub

Public Function FindB2(Wyraz As String) As Long
Dim tbl As String
If Len(Wyraz) < 2 Or Len(Wyraz) > 15 Then
   FindB2 = 0
   Exit Function
End If
tbl = "Sl" & CStr(Len(Wyraz))
FindB2 = Tabele(tbl).FindB2(UCase(Wyraz))
End Function

Public Sub CreateIndeks(Tabela As String)
Tabele(Tabela).CreateIndeks
End Sub

Public Function FindPrev(Wyraz As String) As Long
Dim tbl As String
If Len(Wyraz) < 2 Or Len(Wyraz) > 15 Then
   FindPrev = 0
   Exit Function
End If
tbl = "Sl" & CStr(Len(Wyraz))
FindPrev = Tabele(tbl).FindPrev(UCase(Wyraz))
End Function
Public Sub ClearWyrazy()
   Wyrazy.Clear
End Sub
Public Sub ClearAllWyrazy()
   AllWyrazy.Clear
End Sub
Public Sub Clear(Tabela As String)
   Tabele(Tabela).Clear
End Sub
Public Sub ClearPlus1()
   IleRekPlus12 = 0
   Erase Plus12
   IleRekPlus13 = 0
   Erase Plus13
   IleRekPlus14 = 0
   Erase Plus14
   IleRekPlus15 = 0
   Erase Plus15
   IleRekPlus16 = 0
   Erase Plus16
   IleRekPlus17 = 0
   Erase Plus17
   IleRekPlus18 = 0
   Erase Plus18
End Sub
Public Sub ClearWlasne()
   IleWlasne = 0
   Erase Wlasne
End Sub
Public Function Wyraz(Tabela As String, Numer As Long) As cFullWyraz
Set Wyraz = Tabele(Tabela).Element(Numer)
End Function
Public Sub InsertToAllWyrazy(Wyraz As String, Start As Long, Punkty As Long, Kier As wKierunek, Wart As Long, IleLiter As Long)
AllWyrazy.Add Wyraz, IleLiter, Kier, 0, Start, Punkty, Wart
End Sub
Public Sub InsertToWlasne(IleLiter As Long, WarTab() As String)
Dim Tabela As String, CurTab As cTabela
Dim i As Long, a As String, Warki() As String, U1 As Long, U2 As Long
   Tabela = "Sl" & CStr(IleLiter)
   Set CurTab = Tabele(Tabela)
   ReDim Preserve Wlasne(IleWlasne + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleWlasne = IleWlasne + 1
         Wlasne(IleWlasne) = a
      End If
   Next i
   ReDim Preserve Wlasne(IleWlasne)
Set CurTab = Nothing
End Sub
Private Function Warunki(Litery() As String) As String()
Dim i As Long, War() As String, j As Long, Lit(15) As String * 1
Dim n As Long, Znal As Boolean, k As Long, UL As Long, MaxL As Long
Dim Ml As Long

UL = UBound(Litery)
If UL = -1 Then
   ReDim War(0, 0)
   Warunki = War
   Exit Function
End If
For i = 0 To UL
   MaxL = Len(Litery(i))
   If MaxL > Ml Then Ml = MaxL
Next i
ReDim War(UL, Ml - 1)

For k = 0 To UL
   n = 0
   For i = 1 To Len(Litery(k))
      Znal = False
      For j = 1 To n
         If Lit(j) = Mid(Litery(k), i, 1) Then
            War(k, j - 1) = War(k, j - 1) & Lit(j) & "*"
            Znal = True
            Exit For
         End If
      Next j
      If Znal = False Then
         n = n + 1
         Lit(n) = Mid(Litery(k), i, 1)
         War(k, n - 1) = "*" & Lit(j) & "*"
      End If
   Next i
Next k

k = 0
For i = 0 To UBound(War, 1)
   Znal = False
   For j = UBound(War, 2) To 0 Step -1
      If War(i, j) <> "" Then
         Znal = True
         Exit For
      End If
   Next j
   If Znal = True Then
      If j > k Then k = j
   End If
Next i
ReDim Preserve War(UL, k)
Warunki = War
End Function

Private Function SQL3(Wyraz As String, LenWyraz As Long, Warki() As String, U1 As Long, U2 As Long) As Boolean
Dim k As Long, Chwila As Boolean, w As String
Dim m As Long
   
   For k = 0 To U1
      Chwila = True
      For m = 0 To U2
         w = Warki(k, m)
         If w <> "" Then
            If (Wyraz Like w) = False Then
               Chwila = False
               Exit For
            End If
         Else
            Exit For
         End If
      Next m
      If Chwila = True Then
         SQL3 = True
         Exit Function
      End If
   Next k

SQL3 = False
   
End Function

Public Sub InsertToAllWyrazy2(IleLiter As Long, WarTab() As String, Optional Znaczki As String = "")
Dim Tabela As String, CurTab As cTabela
Dim i As Long, a As String, Warki() As String, U1 As Long, U2 As Long
   
   Tabela = "Sl" & CStr(IleLiter)
   Set CurTab = Tabele(Tabela)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         AllWyrazy.Add a, -1, -1, -1, -1, -1, -1
      End If
   Next i
Set CurTab = Nothing

End Sub
Public Sub InsertToWyrazy2(IleLiter As Long, Warunek As String, Kierunek As wKierunek, Start As Long, PlusJeden As Boolean, WarTab() As String)
Dim Tabela As String, CurTab As cTabela
Dim Wyraziki() As String
Dim i As Long, a As String, Warki() As String, U1 As Long, U2 As Long
Dim Znacz As String, iFirst As Long, iLast As Long, IleWyraziki As Long

If PlusJeden = False Then
   Tabela = "Sl" & CStr(IleLiter)
   Set CurTab = Tabele(Tabela)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   iFirst = StartP(Warunek)
   iLast = Koniec(Warunek)
      
  ' ReDim Wyraziki(CurTab.RecordCount)
    Wyrazy.AddPlace CurTab.RecordCount
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            'Wyraziki(IleWyraziki) = a
    
            Wyrazy.AddWyraz a, IleLiter, Kierunek, -1, Start, -1, -1
         End If
      End If
   Next i

    Wyrazy.SetValidPlace
'   If IleWyraziki > 0 Then
'      ReDim Preserve Wyraziki(IleWyraziki)
'      Wyrazy.AddTab Wyraziki, IleLiter, Kierunek, -1, Start, -1, -1
'   End If
   Set CurTab = Nothing
Else
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   iFirst = StartP(Warunek)
   iLast = Koniec(Warunek)

Select Case IleLiter

Case 2
   ReDim Wyraziki(IleRekPlus12)
   For i = 1 To IleRekPlus12
      a = Plus12(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
Case 3
   ReDim Wyraziki(IleRekPlus13)
   For i = 1 To IleRekPlus13
      a = Plus13(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
Case 4
   ReDim Wyraziki(IleRekPlus14)
   For i = 1 To IleRekPlus14
      a = Plus14(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
Case 5
   ReDim Wyraziki(IleRekPlus15)
   For i = 1 To IleRekPlus15
      a = Plus15(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
Case 6
   ReDim Wyraziki(IleRekPlus16)
   For i = 1 To IleRekPlus16
      a = Plus16(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
Case 7
   ReDim Wyraziki(IleRekPlus17)
   For i = 1 To IleRekPlus17
      a = Plus17(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
Case 8
   ReDim Wyraziki(IleRekPlus18)
   For i = 1 To IleRekPlus18
      a = Plus18(i)
      If Lajk(a, Warunek, iFirst, iLast) <> 0 Then
         If SQL3(a, IleLiter, Warki, U1, U2) = True Then
            'Wyrazy.Add a, IleLiter, Kierunek, -1, Start, -1, -1
            IleWyraziki = IleWyraziki + 1
            Wyraziki(IleWyraziki) = a
         End If
      End If
   Next i
   End Select
   If IleWyraziki > 0 Then
      ReDim Preserve Wyraziki(IleWyraziki)
      Wyrazy.AddTab Wyraziki, IleLiter, Kierunek, -1, Start, -1, -1
   End If
End If

End Sub

Public Sub InsertToPlus1(IleLiter As Long, WarTab() As String)
Dim Tabela As String, CurTab As cTabela
Dim i As Long, a As String, Warki() As String, U1 As Long, U2 As Long

Tabela = "Sl" & CStr(IleLiter)
Set CurTab = Tabele(Tabela)
Select Case IleLiter

Case 2
   ReDim Preserve Plus12(IleRekPlus12 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus12 = IleRekPlus12 + 1
         Plus12(IleRekPlus12) = a
      End If
   Next i
   ReDim Preserve Plus12(IleRekPlus12)
Case 3
   ReDim Preserve Plus13(IleRekPlus13 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus13 = IleRekPlus13 + 1
         Plus13(IleRekPlus13) = a
      End If
   Next i
   ReDim Preserve Plus13(IleRekPlus13)
Case 4
   ReDim Preserve Plus14(IleRekPlus14 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus14 = IleRekPlus14 + 1
         Plus14(IleRekPlus14) = a
      End If
   Next i
   ReDim Preserve Plus14(IleRekPlus14)
Case 5
   ReDim Preserve Plus15(IleRekPlus15 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus15 = IleRekPlus15 + 1
         Plus15(IleRekPlus15) = a
      End If
   Next i
   ReDim Preserve Plus15(IleRekPlus15)
Case 6
   ReDim Preserve Plus16(IleRekPlus16 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus16 = IleRekPlus16 + 1
         Plus16(IleRekPlus16) = a
      End If
   Next i
   ReDim Preserve Plus16(IleRekPlus16)
Case 7
   ReDim Preserve Plus17(IleRekPlus17 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus17 = IleRekPlus17 + 1
         Plus17(IleRekPlus17) = a
      End If
   Next i
   ReDim Preserve Plus17(IleRekPlus17)
Case 8
   ReDim Preserve Plus18(IleRekPlus18 + CurTab.RecordCount)
   Warki = Warunki(WarTab())
   U1 = UBound(Warki, 1)
   U2 = UBound(Warki, 2)
   For i = 1 To CurTab.RecordCount
      a = CurTab.Pole(i)
      If SQL3(a, IleLiter, Warki, U1, U2) = True Then
         IleRekPlus18 = IleRekPlus18 + 1
         Plus18(IleRekPlus18) = a
      End If
   Next i
   ReDim Preserve Plus18(IleRekPlus18)
End Select

Set CurTab = Nothing

End Sub

Public Function ExtDlugosc(Minimum As Boolean) As Long

Dim iMax As Long, iMin As Long, i As Long, a As String
iMin = 8
For i = 1 To IleWlasne
   a = Wlasne(i)
   If Minimum = False Then
      If Len(a) > iMax Then iMax = Len(a)
      If iMax = 7 Then Exit For
   Else
      If Len(a) < iMin Then iMin = Len(a)
      If iMin = 2 Then Exit For
   End If
Next i

If Minimum = False Then
   ExtDlugosc = iMax
Else
   ExtDlugosc = iMin
End If
End Function

Public Function OpenDoDoLitera(NPola As Long, Kierunek As wKierunek) As cMojRecordset
Dim i As Long, r As New cMojRecordset
For i = 1 To IleDoDo
   If DoDo(i).NumerPola = NPola Then
      If DoDo(i).Kierunek = Kierunek Then
         r.Add DoDo(i).Wyraz
      End If
   End If
Next i

Set OpenDoDoLitera = r
Set r = Nothing

End Function

Public Sub AddDoDo(Wyraz As String, NPola As Long, Kierunek As wKierunek)
IleDoDo = IleDoDo + 1
ReDim Preserve DoDo(IleDoDo)
DoDo(IleDoDo).Kierunek = Kierunek
DoDo(IleDoDo).NumerPola = NPola
DoDo(IleDoDo).Wyraz = Wyraz
End Sub
Public Sub ClearDoDo()
Erase DoDo
IleDoDo = 0
End Sub
Public Property Get CountDoDo() As Long
CountDoDo = IleDoDo
End Property
Public Function OpenDoDoNPola(Kierunek As wKierunek, Optional DistinctRow As Boolean = False) As cMojRecordset
Dim i As Long, r As New cMojRecordset, j As Long, Bylo As Boolean

If DistinctRow = False Then
   For i = 1 To IleDoDo
      If DoDo(i).Kierunek = Kierunek Then
         r.Add CStr(DoDo(i).NumerPola)
      End If
   Next i
Else
   For i = 1 To IleDoDo
      If DoDo(i).Kierunek = Kierunek Then
         If r.RecordCount > 0 Then
            Bylo = False
            r.MoveFirst
            Do While r.EOF = False And Bylo = False
               If r.Element() = CStr(DoDo(i).NumerPola) Then
                  Bylo = True
                  Exit Do
               End If
               r.MoveNext
            Loop
            If Bylo = False Then
               r.Add CStr(DoDo(i).NumerPola)
            End If
         Else
            r.Add CStr(DoDo(i).NumerPola)
         End If
      End If
   Next i
End If

Set OpenDoDoNPola = r
Set r = Nothing

End Function
Public Function OpenWlasneWyraz(IleLiter As Long, NumerMid As Long, LiteryMid() As String) As cMojRecordset
Dim i As Long, r As New cMojRecordset, Jest As Boolean, j As Long

If NumerMid > 0 Then
   For i = 1 To IleWlasne
      If Len(Wlasne(i)) = IleLiter Then
         Jest = False
         For j = 1 To UBound(LiteryMid)
            If Mid(Wlasne(i), NumerMid, 1) = LiteryMid(j) Then
               Jest = True
               Exit For
            End If
         Next j
         If Jest = True Then
            r.Add Wlasne(i)
         End If
      End If
   Next i
Else
   For i = 1 To IleWlasne
      If Len(Wlasne(i)) = IleLiter Then
         r.Add Wlasne(i)
      End If
   Next i
End If
Set OpenWlasneWyraz = r
Set r = Nothing

End Function

Public Function OpenMidTabela(Tabela As String, Wzorzec As String, NumerMid As Long, LiteryMid() As String) As cMojRecordset
Dim i As Long, r As New cMojRecordset, Jest As Boolean, j As Long
Dim t As cTabela, Indeks As cIndeks, iFirst As Long, iLast As Long
Dim Indeksik As cIndeksik, LW As Long, St As Long, Ko As Long

Set t = Tabele(Tabela)
Set Indeks = Indeksy(Tabela)

If Left(Wzorzec, 1) = "_" Then
   iFirst = 1
   iLast = t.RecordCount
Else
   On Error GoTo NieMaIndeksu
   Set Indeksik = Indeks.Znaczki(Left(Wzorzec, 1))
   On Error GoTo 0
   iFirst = Indeksik.First
   iLast = Indeksik.Last
End If

LW = Len(Wzorzec)
St = StartP(Wzorzec)
Ko = Koniec(Wzorzec)

If UBound(LiteryMid) <> 0 Then
   For i = iFirst To iLast
      If Lajk(t.Pole(i), Wzorzec, St, Ko) <> 0 Then
         Jest = False
         For j = 0 To UBound(LiteryMid)
            If Mid(t.Pole(i), NumerMid, 1) = LiteryMid(j) Then
               Jest = True
               Exit For
            End If
         Next j
         If Jest = True Then
            r.Add Mid(t.Pole(i), NumerMid, 1)
         End If
      End If
   Next i
Else
   For i = iFirst To iLast
      If Lajk(t.Pole(i), Wzorzec, St, Ko) <> 0 Then
         r.Add Mid(t.Pole(i), NumerMid, 1)
      End If
   Next i
End If

NieMaIndeksu:
Set t = Nothing
Set Indeks = Nothing
Set OpenMidTabela = r
Set r = Nothing

End Function

Public Sub SortujAW(MaxPkt As Long)
Dim i As Long
Set AllWyrazySort = New cWyrazy
AllWyrazySort.Clear
For i = 1 To AllWyrazy.RecordCount
   If AllWyrazy.Element(i).Punkty = MaxPkt Then
      AllWyrazySort.AddObject AllWyrazy.Element(i)
   End If
Next i
End Sub

Public Function Plus1Count() As Long
Plus1Count = IleRekPlus12 + IleRekPlus13 + IleRekPlus14 + IleRekPlus15 + IleRekPlus16 + IleRekPlus17 + IleRekPlus18
End Function

Public Function CountWlasne() As Long
CountWlasne = IleWlasne
End Function

Public Sub DodajWyraz(Wyraz As String)
Dim Tabela As String, i As Long, t As cTabela
i = Len(Wyraz)
If i < 2 Or i > 15 Then Exit Sub
Tabela = "Sl" & CStr(i)
Set t = Tabele(Tabela)
t.DodajWyraz Wyraz
End Sub

Public Sub AktualizujAll()
Dim t As cTabela

For Each t In Tabele
   t.Dodawaj
   t.Aktualizuj
Next
End Sub

Public Sub InsertToWyrazyCPP(IleLiter As Long, Warunek As String, Kierunek As wKierunek, Start As Long, PlusJeden As Boolean, WarTab() As String)


End Sub
