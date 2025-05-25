Attribute VB_Name = "Module1"
Option Explicit
Enum wKierunek
   wBrak = 0
   wPionowy = 1
   wPoziomy = 2
   wSama = 3
End Enum
Enum Kieruj
   Pionowy = 0
   Poziomy = 1
End Enum
Public Type tPole
   Numer As Long
   Mnoznik As Long
End Type
Private Type tDanePola
   RazyL As Long
   RazyS As Long
   Kolor As ColorConstants
   Opis As String
End Type
Public Type Tpunktacja
   Punkty As Long
   Wartosc As Long
   IleLiter As Long
   Wsp As Long
End Type
Public Type tAllWyrazy
   Wyraz As String
   Start As Long
   Punkty As Long
   Kierunek As wKierunek
   IleLiter As Long
   Wart As Long
End Type
Public Type tDoDo
   Slowo As String
   Start As Long
   Kierunek As wKierunek
End Type
Public Type tWlasne
   Wyraz As String
   Dlugosc As Long
End Type

Global Const bok As Long = 33
Global Const px As Long = 5
Global Const py As Long = 5
Global Const Prz As Long = 1
Global Const MinLiter As Long = 7
Global MojaBaza As cBazaDanych, JuzKoniec As Boolean
Global BylLad As Boolean, PlanszaWART(0 To 224) As Long, PlanszaBLANK(0 To 224) As Long
Global IleWymian As Long, AllWyrazy() As tAllWyrazy, Wyrazy() As tAllWyrazy, Plus1() As String
Global Le¿¹ce As Collection, IleLiter As Long, CzasCTak As Boolean, CzasRTak As Boolean
Global Defalt(1 To 4) As String
Global Jêzyki As Collection, Jêzyk As tJezyk, Wart(255) As Long, JuzGramy As Boolean
Global kKierunek As Kieruj, max(4) As Long, WymianaMydla As Boolean
Global Kolkon(225) As Long, Mydlo As String, Ktory As Long
Global Wolne(100) As Boolean, Pole(225) As Klocek3D, plytka(7) As Klocek3D
Global OldMydlo As String, Stojak(1 To 4, 0 To 6) As String, Worek(100) As String
Global Czas(4) As Long, FMax As Long, RMax As Long, Zero As Long
Global PLit() As String, DFMax As Long, DRMax As Long, IleGraczy As Long, Gracz(4) As String
Global Omin As Long, Demo As Long
Global PLR(4) As GR, BrakAnswer As Integer, Czasy As Collection
Global Lst As Tabela, w As Wiersz, e As Element, KoniecGry As Boolean
Global ConStr As String, MaxPsz As Long, Slowa As Collection
Global m_Server As CServer
Global m_Klient As CKlient
'Global m_CBaza As LITDLLEXTLib.Literak

Public Declare Function WarunkujADO Lib "LowSlow.dll" (ByVal Litery As String) As String
Public Declare Function PolujCPP Lib "LowSlow.dll" (ByRef PolaTAB As tPole, ByVal Dlugosc As Long, ByVal Ile As Long, ByVal Start As Long, ByVal Kierunek As wKierunek, PlanszaTAG As Long, PlanszaRAZYS As Long, PlanszaRAZYL As Long) As Long
Public Declare Function VirtPunkty Lib "LowSlow.dll" (ByVal IleLiter As Long, ByRef Lezace As Long, ByVal IleLezy As Long, ByVal Kierunek As wKierunek, ByRef PlanszaWART As Long, ByRef PlanszaRAZYL As Long, ByRef PlanszaRAZYS As Long, ByRef PlanszaTAG As Long, ByRef PlanszaLICZ As Long) As Tpunktacja
Public Declare Function QSrtPOLA Lib "LowSlow.dll" (ByRef tbl As tPole, ByVal Y As Long, ByVal Y As Long, ByVal Descending As Long) As Long
Public Declare Function MaxWynik Lib "LowSlow.dll" (KPola As tPole, KLitery As Long, ByVal Dlugosc As Long, ByVal Start As Long, ByVal IlePol As Long, ByVal Kierunek As wKierunek, PlanszaRAZYS As Long, PlanszaTAG As Long, PlanszaWART As Long, PlanszaRAZYL As Long) As Long
Public Declare Function wx Lib "LowSlow.dll" Alias "WX" (ByVal Liczba As Long) As Long
Public Declare Function wy Lib "LowSlow.dll" Alias "WY" (ByVal Liczba As Long) As Long
Public Declare Function ciagiem Lib "LowSlow.dll" Alias "CIAGIEM" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function VirtUstalWzorzec2 Lib "LowSlow.dll" (ByVal Numer As Long, ByVal Kierunek As wKierunek, ByVal nLiter As Long, ByVal VMaxPsz As Long, ByRef PlanszaTAG As Long, ByVal Kaption As String, ByRef Zakaz As Long) As String
Public Declare Function VirtMinusST2 Lib "LowSlow.dll" (ByVal NPola As Long, ByVal Kierunek As wKierunek, ByRef PlanszaTAG As Long) As Long
Public Declare Function CzySzukac2 Lib "LowSlow.dll" (ByVal Numer As Long, ByVal Kierunek As wKierunek, ByVal Ile As Long, ByRef PlanszaTAG As Long, ByRef Zakaz As Long) As Long
Public Declare Function JakiStart3 Lib "LowSlow.dll" (ByVal NPola As Long, ByVal Kierunek As wKierunek, ByVal Dlugosc As Long, ByVal LW As Long, ByRef PlanszaTAG As Long, ByRef Zakaz As Long) As Long
Public Declare Function WzorkujCPP Lib "LowSlow.dll" (ByVal NPola As Long, ByVal Kierunek As wKierunek, ByVal Prawo As Long, ByRef PlanszaTAG As Long, ByVal Kaption As String) As String

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Declare Function GrajWave Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Public Const Swp_Nosize = &H1
Public Const Swp_Nomove = &H2
Public Const Swp_NoActivate = &H10
Public Const Hwnd_TopMost = -1
Public Const Hwnd_NoTopMost = -2
Public Const Swp_ShowWindow = &H40
Sub Main()

'If autoryzuj = False Then
   'MsgBox "Numer rejestaracyjny nieprawid³owy"
   'End
'End If
'Unload Autory

'Set m_CBaza = New LITDLLEXTLib.Literak

'm_CBaza.AddWord "Dzien dobry"

On Error GoTo 0
Call Zaczynaj
Call LadujJêzyki
graslow.Show

End Sub

Public Sub Zaczynaj()
Dim i As Long, j As Long, tdp As tDanePola

For i = 1 To 4
    Set PLR(i) = New GR
Next i

For i = 0 To 224
    Set Pole(i) = New Klocek3D
    Pole(i).bok = bok
    Set Pole(i).Kartka = graslow.T³o
    Pole(i).MojNumer = i
    Pole(i).Prz = Prz
    Pole(i).px = px
    Pole(i).py = py
    Pole(i).Pltk = False
    tdp = Koloruj(i)
    Pole(i).Inicjuj tdp.RazyL, tdp.RazyS, tdp.Kolor, tdp.Opis
Next i

For i = 0 To 6
   Set plytka(i) = New Klocek3D
   plytka(i).MojNumer = i
   plytka(i).bok = bok
   plytka(i).Prz = Prz + 1
   plytka(i).Pltk = True
   plytka(i).px = 0
   plytka(i).py = 0
   Set plytka(i).Kartka = graslow.Podstawka
   tdp = Koloruj(i)
   Pole(i).Inicjuj tdp.RazyL, tdp.RazyS, tdp.Kolor, tdp.Opis
Next i

Mydlo = Chr$(32)
IleLiter = 100
graslow.T³o.Left = 0
graslow.T³o.Top = 0
graslow.T³o.Width = 15 * (bok + Prz) + 8
graslow.T³o.Height = 15 * (bok + Prz) + 9

For i = 0 To 224
   Pole(i).PustePole
Next i

Worek(98) = Mydlo
Worek(99) = Mydlo
Set Le¿¹ce = New Collection

Set Lst = New Tabela

graslow.Podstawka.Width = 7 * (bok + Prz + 1) - Prz - 1
Pokazuj False

Defalt(1) = GetSetting("£ów S³ów", "Gracze", "PLR1", "Gracz 1")
Defalt(2) = GetSetting("£ów S³ów", "Gracze", "PLR2", "Gracz 2")
Defalt(3) = GetSetting("£ów S³ów", "Gracze", "PLR3", "Gracz 3")
Defalt(4) = GetSetting("£ów S³ów", "Gracze", "PLR4", "Gracz 4")

For i = 1 To 3
   graslow.player(i).MousePointer = graslow.player(0).MousePointer
   graslow.player(i).MouseIcon = graslow.player(0).MouseIcon
   graslow.player(i).Left = graslow.player(0).Left
Next i
For i = 1 To 4
   graslow.AkcjeSub(i).Enabled = False
Next i

IleGraczy = 0
Ktory = 1
graslow.Podstawka.Left = 100
graslow.Podstawka.Top = 520
graslow.OdczytUstaw
End Sub

Public Sub NowaGra()
Dim i As Long, Tmp As String, j As Long

If IleGraczy Then
    i = InfoBox("Czy na pewno chcesz rozpocz¹æ grê od nowa ?", True)
    If i = 2 Then Exit Sub
End If

graslow.Zegar.Enabled = False
For i = 0 To 224
   Pole(i).PustePole
Next i
Demo = 0
Erase PLR
'For i = 1 To 4
'   Set PLR(i) = New GR
'Next i

Zacz.Show vbModal

If IleGraczy = 0 Then
    Exit Sub
End If


For i = 0 To 3
   graslow.player(i).Visible = False
   graslow.wynik(i).Visible = False
   graslow.Kolorek(i).Visible = False
   graslow.cas(i).Visible = False
Next i

For i = 1 To IleGraczy
   Tmp = ImiêGracza("Proszê podaæ Imiê/Nazwisko/Pseudonim zawodnika numer " & i, Defalt(i), i - 1)
   Set PLR(i) = New GR
   PLR(i).imie = Left(Tmp, Len(Tmp) - 1)
   graslow.wynik(i - 1).Caption = PLR(i).wynik
   graslow.player(i - 1).Caption = PLR(i).imie
   graslow.cas(i - 1).Caption = graslow.FormaCzasu(FMax - PLR(i).CzasCa³kowity)
   graslow.player(i - 1).Caption = PLR(i).imie
   graslow.player(i - 1).Visible = True
   graslow.wynik(i - 1).Visible = True
   graslow.Kolorek(i - 1).Visible = True
   graslow.cas(i - 1).Visible = True
   If Right(Tmp, 1) = "1" Then
       PLR(i).Komp = True
       PLR(i).MaxPsz = MaxPsz
   Else
       PLR(i).Komp = False
   End If
   If PLR(i).imie = "" Then PLR(i).imie = Defalt(i)
   Defalt(i) = PLR(i).imie
   SaveSetting "£ów S³ów", "Gracze", "PLR" & i, PLR(i).imie
   graslow.player(i - 1).FontBold = False
   graslow.Kolorek(i - 1).Picture = graslow.Obrazki.ListImages.item(i + 3).Picture
   graslow.wynik(i - 1).FontBold = False
   graslow.cas(i - 1).FontBold = False
Next i

If Zacz.Check1.Value = 1 Then
   IleWymian = Zacz.IleW
Else
   IleWymian = 99
End If

Call Pokazuj(True)
graslow.IleLitCap.Caption = CStr(IleLiter)
DoEvents
For i = 0 To 6
    plytka(i).PustePole
    plytka(i).Po³ó¿ "_", False
    plytka(i).ZatwierdŸ 3
    For j = 1 To 4
      Stojak(j, i) = "_"
   Next j
Next i

For i = 0 To 99
   Wolne(i) = False
Next i

Randomize Timer
Ktory = Int(Rnd() * IleGraczy) + 1
IleLiter = 100
For i = 1 To IleGraczy
    For j = 0 To 6
        If Left(Stojak(i, j), 1) = "_" Or Stojak(i, j) = "" Then Stojak(i, j) = graslow.Losuj()
    Next j
Next i

graslow.MousePointer = 1

If JuzGramy = False Then
    If Not Jêzyk Is Nothing Then
        graslow.Caption = "£ów S³ów - " & Jêzyk.Nazwa
    End If
    graslow.DefMyd1.Clear

    If m_Klient Is Nothing Then
        For i = 0 To UBound(PLit) - 2
            graslow.DefMyd1.AddItem PLit(i)
        Next i
    
        graslow.DefMyd1.ListIndex = 0
    End If
End If
JuzGramy = True

Set Czasy = Nothing
Set Czasy = New Collection

KoniecGry = False
Zero = 0
Omin = 0

graslow.MousePointer = 0
graslow.IleLitCap.BackColor = vbGreen

If m_Klient Is Nothing Then
    graslow.Caption = "£ów S³ów - " & Jêzyk.Nazwa
    Call graslow.Punktuj
End If
Call graslow.Anuluj
Call graslow.CzytajHS
graslow.Zegar.Interval = 1000
graslow.Zegar.Enabled = False
Call graslow.NextPlayer

End Sub

Public Function InfoBox(Tekst As String, Optional Nie As Boolean, Optional Dopisz As Boolean) As Integer

If Nie = False And Dopisz = False Then
    Brak.tak.Caption = "OK"
    Brak.tak.Left = Brak.Width / 2 - Brak.tak.Width / 2
    Brak.Komunikat.Height = 2000
    Brak.tak.Top = Brak.Height - Brak.tak.Height - 500
Else
    Brak.tak.Caption = "Tak"
    Brak.tak.Left = 840
    Brak.tak.Top = Brak.Nie.Top
End If
Brak.Nie.Visible = Nie
Brak.Dopisz.Visible = Dopisz
Brak.Komunikat.Caption = Tekst
Brak.Show vbModal
InfoBox = BrakAnswer

End Function
Public Sub KoniecLS()
Dim item As Form
graslow.ZapiszUstaw
Set MojaBaza = Nothing
Set Jêzyki = Nothing
Set Czasy = Nothing
Set Le¿¹ce = Nothing
Set Jêzyk = Nothing
Set Slowa = Nothing
Set e = Nothing
Set w = Nothing
Set Lst = Nothing
Set MojaBaza = Nothing
For Each item In Forms
   Unload item
Next item
End
End Sub
Public Function ImiêGracza(Komunikat As String, imie As String, nrGR As Long) As String
Dim k As Long
PodaName.Komunikat.Caption = Komunikat
PodaName.imie.Text = imie
PodaName.imie.SelStart = 0
PodaName.imie.SelLength = Len(imie)
PodaName.nrGR = nrGR
PodaName.Show vbModal
If PodaName.kom = True Then
    k = 1
Else
    k = 0
End If
ImiêGracza = PodaName.im & k
End Function

Public Sub LadujJêzyki()
Dim hj As Long, Jezyk As tJezyk, a As String, b As String, c As String
Dim i As Long, d As String
hj = FreeFile
Open App.Path & "\Jezyki.cfg" For Input As hj
Set Jêzyki = New Collection
While Not EOF(hj)
    Input #hj, a, b, c, d
    Set Jezyk = New tJezyk
    Jezyk.Nazwa = a
    Jezyk.Plik = b
    Jezyk.S³ownik = c
    Jezyk.Klucz = d
    Jêzyki.Add Jezyk, Jezyk.Klucz
Wend
Set Jezyk = Nothing
Close hj

End Sub

Public Sub Pokazuj(Stan As Boolean)
Dim i As Long
For i = 1 To 4
graslow.AkcjeSub(i).Enabled = Stan
Next i
graslow.Zegar.Enabled = False
graslow.Timer1.Enabled = False
graslow.Podstawka.Visible = Stan
graslow.MenuUstaw.Enabled = Stan
graslow.MenuSlowa1.Enabled = Stan
graslow.MenuSave.Enabled = Stan
graslow.MenuAkcje.Enabled = Stan
graslow.anulujemy.Visible = Stan
graslow.Solve.Visible = Stan
graslow.DefMyd1.Visible = Stan
graslow.Label1.Visible = Stan
graslow.Kolejka.Visible = Stan
graslow.wymiana.Visible = Stan
graslow.lblRekord.Visible = Stan
graslow.NazwaGracza.Visible = Stan
graslow.WynikGracza.Visible = Stan
graslow.IleLitCap.Visible = Stan
graslow.EfektyTu.Visible = Stan


End Sub

Public Function Koloruj(Numer As Long) As tDanePola
Dim Kolor(0 To 3) As ColorConstants, Bck As ColorConstants
Dim tdp As tDanePola
Kolor(1) = RGB(255, 64, 0)
Kolor(0) = RGB(255, 160, 150)
Kolor(2) = RGB(64, 128, 255)
Kolor(3) = RGB(128, 192, 255)
Select Case Numer
    Case 112, 16, 32, 48, 64, 28, 42, 56, 70, 208, 192, 176, 160, 154, 168, 182, 196: tdp.Kolor = Kolor(0): tdp.Opis = "2S": tdp.RazyS = 2: tdp.RazyL = 1
    Case 0, 7, 14, 224, 210, 217, 105, 119: tdp.Kolor = Kolor(1): tdp.Opis = "3S": tdp.RazyS = 3: tdp.RazyL = 1
    Case 20, 24, 80, 84, 76, 88, 140, 144, 136, 148, 200, 204: tdp.Kolor = Kolor(2): tdp.Opis = "3L": tdp.RazyL = 3: tdp.RazyS = 1
    Case 3, 11, 36, 38, 52, 45, 59, 92, 122, 108, 96, 98, 126, 128, 132, 102, 221, 213, 165, 172, 179, 186, 188, 116: tdp.Kolor = Kolor(3): tdp.Opis = "2L": tdp.RazyL = 2: tdp.RazyS = 1
    Case Else: tdp.Kolor = RGB(0, 128, 0): tdp.Opis = "": tdp.RazyS = 1: tdp.RazyL = 1
End Select

Koloruj = tdp
End Function

Public Sub ZapiszDom1(IG As Long, iw As Long, nrJezyk As Long, CRO As Boolean, CCO As Boolean, WO As Boolean, CR As Long, CC As Long)
Dim i As Long
SaveSetting "£ów S³ów", "Ustawienia", "IleGraczy", CStr(IG)
SaveSetting "£ów S³ów", "Ustawienia", "CzasRuchuO", CStr(Abs(CLng(CRO)))
SaveSetting "£ów S³ów", "Ustawienia", "CzasCa³yO", CStr(Abs(CLng(CCO)))
SaveSetting "£ów S³ów", "Ustawienia", "WymianyO", CStr(Abs(CLng(WO)))
SaveSetting "£ów S³ów", "Ustawienia", "IleWymian", CStr(iw)
SaveSetting "£ów S³ów", "Ustawienia", "CzasRuchu", CStr(CR)
SaveSetting "£ów S³ów", "Ustawienia", "CzasCa³y", CStr(CC)
SaveSetting "£ów S³ów", "Ustawienia", "Jêzyk", CStr(nrJezyk)
End Sub

Public Sub ZapiszDom2(TypGracza() As Boolean, LevelGracza() As Long)
Dim i As Long
For i = 0 To UBound(TypGracza)
   SaveSetting "£ów S³ów", "Ustawienia", "TypGracza" & CStr(i + 1), CStr(TypGracza(i))
Next i
For i = 0 To UBound(LevelGracza)
   SaveSetting "£ów S³ów", "Ustawienia", "LevelGracza" & CStr(i + 1), CStr(LevelGracza(i))
Next i

End Sub
