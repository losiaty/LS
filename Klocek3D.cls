VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Klocek3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Kartka As PictureBox, bok As Single, Prz As Single
Public Drukarka As Object, Blank As Boolean
'Public Litery(100) As String, WartoscLitery(100) As Long
Public px As Single, py As Single, Tag As Boolean, licz As Boolean
Public NumerObrazka As Long, Pltk As Boolean
Private MeRazyS As Long, MeRazyL As Long, MNumer As Long
Private PLit(32) As String, MyOpis As String
Private MyFont As String, PrzX As Long, PrzY As Long, MyKolor As ColorConstants
Public Litera As String
Public Property Get Caption() As String
If Pltk = False Then
    Caption = Litera
Else
    If Blank = True Then
        Caption = Mydlo
    Else
        Caption = Litera
    End If
End If
End Property
Public Property Get Wartoœæ() As Long

If Blank = True Or Litera = "" Then
    Wartoœæ = 0
Else
    Wartoœæ = Wart(Asc(Litera))
End If

End Property
Public Property Get RazyS() As Long
RazyS = MeRazyS
End Property
Public Property Get RazyL() As Long
RazyL = MeRazyL
End Property
Public Sub PustePole()
Dim MyX As Single, MyY As Single

NumerObrazka = 0
Blank = False
MyX = px + (bok + Prz) * (wx(MojNumer) - 1)
MyY = py + (bok + Prz) * (wy(MojNumer) - 1)

Kartka.Line (MyX, MyY)-(MyX + bok - 1, MyY + bok - 1), MyKolor, BF
Kartka.FontSize = 17
Kartka.Font = "Arial"
Kartka.Font.Bold = True
Kartka.ForeColor = vbWhite
Kartka.CurrentX = MyX + (bok - Kartka.TextWidth(MyOpis)) / 2
Kartka.CurrentY = MyY + (bok - Kartka.TextHeight(MyOpis)) / 2
Kartka.Print MyOpis
Tag = False
licz = False
Litera = ""
End Sub
Private Function wy(Numer As Long) As Long
wy = Numer \ 15 + 1
End Function
Private Function wx(Numer As Long) As Long
wx = ((Numer + 1) - (Numer \ 15) * 15)
End Function
Private Function ciagiem(ByVal wx As Long, ByVal wy As Long) As Long
ciagiem = wx + (wy * 15) - 16
End Function
Public Property Get MojNumer() As Long
MojNumer = MNumer
End Property
Public Property Let MojNumer(ByVal v As Long)
MNumer = v
End Property
Private Sub LoadKolory()
'Kolor(0) = vbYellow
'Kolor(1) = RGB(255, 128, 128)
'Kolor(2) = RGB(160, 160, 255)
'Kolor(3) = RGB(64, 255, 64)
'Kolor(1) = RGB(255, 64, 0)
'Kolor(0) = RGB(255, 160, 150)
'Kolor(2) = RGB(64, 128, 255)
'Kolor(3) = RGB(128, 192, 255)
End Sub
Private Sub Class_Initialize()
LoadKolory
MyFont = "Courier New"
End Sub
Public Sub Po³ó¿(Lit As String, JestMydlo As Boolean, Optional ikonuj As Boolean = False)
Dim MyX As Single, MyY As Single, Pkt As Long, NewBok As Long, LitX As Long, LitY As Long
Dim kl1 As ColorConstants, kl2 As ColorConstants, kl3 As ColorConstants

NumerObrazka = 1
Blank = JestMydlo
If ikonuj = False Then
   NewBok = bok - 2
Else
   NewBok = bok
End If
If Lit <> "" Then Litera = Lit
If Litera = "_" Then
    Exit Sub
End If
'kl1 = RGB(250, 180, 130)
'kl2 = RGB(175, 120, 85)
'kl3 = RGB(80, 60, 40)
kl1 = vbWhite
kl2 = RGB(128, 128, 128)
kl3 = vbBlack
Kartka.ForeColor = vbBlack
Randomize Timer
If ikonuj = False Then
   If Pltk = True Then
      PrzY = 2
   Else
      PrzY = Int(Rnd() * 3)
   End If
   PrzX = Int(Rnd() * 3)
Else
   PrzY = 0
   PrzX = 0
End If

MyX = px + (bok + Prz) * (wx(MojNumer) - 1) + PrzX
MyY = py + (bok + Prz) * (wy(MojNumer) - 1) + PrzY
'Kartka.PaintPicture graslow.Obrazki.ListImages.item(NumerObrazka).Picture, MyX, MyY, NewBok - 2, NewBok - 2
Kartka.PaintPicture graslow.Obrazki.ListImages.item(2).Picture, MyX, MyY, NewBok - 2, NewBok - 2
'Kartka.PaintPicture graslow.Obrazki.ListImages.item(NumerObrazka).Picture, MyX, MyY, NewBok - 2, NewBok - 2
Kartka.Line (MyX, MyY)-(MyX + 1, MyY + NewBok - 1), kl1, BF
Kartka.Line (MyX, MyY)-(MyX + NewBok - 1, MyY + 1), kl1, BF
Kartka.Line (MyX, MyY + NewBok - 1)-(MyX + NewBok - 1, MyY + NewBok - 1), kl3, BF
Kartka.Line (MyX + NewBok - 1, MyY)-(MyX + NewBok - 1, MyY + NewBok - 1), kl3, BF
Kartka.Line (MyX + NewBok - 2, MyY + 1)-(MyX + NewBok - 2, MyY + NewBok - 2), kl2, BF
Kartka.Line (MyX + 1, MyY + NewBok - 2)-(MyX + NewBok - 2, MyY + NewBok - 2), kl2, BF
'Kartka.FontSize = 19
'Kartka.FontName = MyFont
'Kartka.Font.Bold = True
'Kartka.Font.Charset = 238
'Kartka.CurrentX = MyX + (NewBok - Kartka.TextWidth(Litera)) / 2 - 3
'Kartka.CurrentY = MyY + 2
'
'If Blank = True Then
'    Kartka.ForeColor = RGB(128, 128, 128)
'    Pkt = 0
'Else
'    Kartka.ForeColor = vbBlack
'    Pkt = (Wart(Asc(Litera)))
'End If
'
'Kartka.Print (Litera)
'Kartka.FontSize = 10
'Tag = True
'licz = True
'
'Kartka.CurrentX = MyX + NewBok - Kartka.TextWidth(Pkt) - 6
'Kartka.CurrentY = MyY + NewBok - Kartka.TextHeight(Pkt) - 1
'If Blank Then Exit Sub
'Kartka.Print Pkt
'If Wart(Asc(Litera)) = 0 Then Exit Sub

Kartka.FontSize = 19
Kartka.FontName = MyFont
Kartka.Font.Bold = True
Kartka.Font.Charset = 238

LitX = MyX + (NewBok - Kartka.TextWidth(Litera)) / 2 - 4
LitY = MyY + 2

Kartka.ForeColor = vbWhite
Kartka.CurrentX = LitX + 1
Kartka.CurrentY = LitY + 1
Kartka.Print (Litera)

Kartka.CurrentX = LitX + 1
Kartka.CurrentY = LitY
Kartka.Print (Litera)

Kartka.CurrentX = LitX
Kartka.CurrentY = LitY + 1
Kartka.Print (Litera)

Kartka.CurrentX = LitX
Kartka.CurrentY = LitY

If Blank = True Then
    Kartka.ForeColor = RGB(64, 128, 64)
    Pkt = 0
Else
    Kartka.ForeColor = RGB(80, 90, 0) 'vbBlack
    Pkt = (Wart(Asc(Litera)))
End If

Kartka.Print (Litera)

Kartka.FontSize = 11
Kartka.FontBold = True
Kartka.FontName = "Times New Roman"
Kartka.CurrentX = MyX + NewBok - Kartka.TextWidth(Wart(Asc(Litera))) - 6
Kartka.CurrentY = MyY + NewBok - Kartka.TextHeight(Wart(Asc(Litera))) - 1
Tag = True
licz = True

'Set ob = Nothing
If Blank Then Exit Sub
Kartka.Print Pkt


End Sub
Public Property Get Font() As String
Font = MyFont
End Property
Public Property Let Font(ByVal vNewValue As String)
MyFont = vNewValue
End Property
Public Sub ZatwierdŸ(Optional ByVal NObrazka As Long = 2, Optional ikonuj As Boolean = False)
Dim MyX As Single, MyY As Single, ob As StdPicture, i As Long
Dim NewBok As Long, Pkt As Long, LitX As Long, LitY As Long
If ikonuj = False Then
   NewBok = bok - 2
Else
   NewBok = bok
End If
NumerObrazka = NObrazka
If Litera = "_" Then
    PustyStojak
    Exit Sub
End If
Kartka.ForeColor = vbBlack
If NumerObrazka = 2 Then
    Randomize Timer
    i = Int(Rnd() * 4 + 1)
    Set ob = graslow.Klocki.ListImages.item(i).Picture
Else
   Set ob = graslow.Obrazki.ListImages.item(NumerObrazka).Picture
End If
'PrzX = Int(Rnd * 1 + 1)
'PrzY = Int(Rnd * 1 + 1)
MyX = px + (bok + Prz) * (wx(MojNumer) - 1) + PrzX
MyY = py + (bok + Prz) * (wy(MojNumer) - 1) + PrzY

Kartka.PaintPicture ob, MyX, MyY, NewBok - 2, NewBok - 2
Kartka.Line (MyX, MyY)-(MyX + 1, MyY + NewBok - 1), vbWhite, BF
Kartka.Line (MyX, MyY)-(MyX + NewBok - 1, MyY + 1), vbWhite, BF
Kartka.Line (MyX, MyY + NewBok - 1)-(MyX + NewBok - 1, MyY + NewBok - 1), vbBlack, BF
Kartka.Line (MyX + NewBok - 1, MyY)-(MyX + NewBok - 1, MyY + NewBok - 1), vbBlack, BF
Kartka.Line (MyX + NewBok - 2, MyY + 1)-(MyX + NewBok - 2, MyY + NewBok - 2), RGB(128, 128, 128), BF
Kartka.Line (MyX + 1, MyY + NewBok - 2)-(MyX + NewBok - 2, MyY + NewBok - 2), RGB(128, 128, 128), BF
Kartka.FontSize = 19
Kartka.FontName = MyFont
Kartka.Font.Bold = True
Kartka.Font.Charset = 238

LitX = MyX + (NewBok - Kartka.TextWidth(Litera)) / 2 - 4
LitY = MyY + 2

Kartka.ForeColor = vbWhite
Kartka.CurrentX = LitX + 1
Kartka.CurrentY = LitY + 1
Kartka.Print (Litera)

Kartka.CurrentX = LitX + 1
Kartka.CurrentY = LitY
Kartka.Print (Litera)

Kartka.CurrentX = LitX
Kartka.CurrentY = LitY + 1
Kartka.Print (Litera)

Kartka.CurrentX = LitX
Kartka.CurrentY = LitY

If Blank = True Then
    Kartka.ForeColor = RGB(64, 128, 64)
    Pkt = 0
Else
    Kartka.ForeColor = RGB(80, 90, 0) 'vbBlack
    Pkt = (Wart(Asc(Litera)))
End If

Kartka.Print (Litera)

Kartka.FontSize = 11
Kartka.FontBold = True
Kartka.FontName = "Times New Roman"
Kartka.CurrentX = MyX + NewBok - Kartka.TextWidth(Wart(Asc(Litera))) - 6
Kartka.CurrentY = MyY + NewBok - Kartka.TextHeight(Wart(Asc(Litera))) - 1
licz = False
Set ob = Nothing
If Blank Then Exit Sub
Kartka.Print Pkt

End Sub

Public Sub PustyStojak()
Dim ob As StdPicture, Plik As String, MyX As Long, MyY As Long
Set ob = graslow.Podstawki.ListImages.item(MojNumer + 1).Picture
MyX = px + (bok + Prz) * (wx(MojNumer) - 1)
MyY = py + (bok + Prz) * (wy(MojNumer) - 1)
Kartka.PaintPicture ob, MyX, MyY, bok + 2, bok + 1
End Sub
Public Sub Drukuj()
Dim MyX As Single, MyY As Single, ob As Object, i As Long
MyX = px + (bok + Prz) * (wx(MojNumer) - 1)
MyY = py + (bok + Prz) * (wy(MojNumer) - 1)
Drukarka.Line (MyX, MyY)-(MyX + 1, MyY + bok - 1), vbWhite, BF
Drukarka.FontSize = 14
Drukarka.FontName = MyFont
Drukarka.Font.Bold = False
Drukarka.Font.Charset = 238
Drukarka.CurrentX = MyX + (bok - Drukarka.TextWidth(Litera)) / 2 - 3
Drukarka.CurrentY = MyY + 2
Drukarka.Print (Litera)
Drukarka.FontSize = 10
Drukarka.CurrentX = MyX + bok - Drukarka.TextWidth(Wart(Asc(Litera))) - 6
Drukarka.CurrentY = MyY + bok - Drukarka.TextHeight(Wart(Asc(Litera))) - 1
If Wart(Asc(Litera)) = 0 Then Exit Sub
Drukarka.Print (Wart(Asc(Litera)))
End Sub

Private Sub Class_Terminate()
Set Kartka = Nothing
Set Drukarka = Nothing
End Sub
Public Sub Inicjuj(RL As Long, RS As Long, Kolor As ColorConstants, Opis As String)
MeRazyL = RL
MeRazyS = RS
MyKolor = Kolor
MyOpis = Opis
End Sub

Public Sub Zaslon()
''kartka.zaslonka.picture
'Dim MyX As Long, MyY As Long
'MyX = px + (bok + Prz) * (wx(MojNumer) - 1)
'MyY = py + (bok + Prz) * (wy(MojNumer) - 1)
'graslow.Zaslonka.Left = MyX
'graslow.Zaslonka.Top = MyY
'graslow.Zaslonka.Width = bok
'graslow.Zaslonka.Height = bok
'Set graslow.Zaslonka.Picture = graslow.Podstawki.ListImages.item(MojNumer + 1).Picture
'graslow.Zaslonka.Visible = True
End Sub
