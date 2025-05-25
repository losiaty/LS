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
Private Type Elementy
    Text As String
    FontSize As Integer
    BackColor As ColorConstants
    ForeColor As ColorConstants
    Align As AlignConstants
End Type
Private Type Formaty
    FontSize As Integer
    FontItalic As Boolean
    FontBold As Boolean
    FontUnderline As Boolean
    FontName As String
End Type
    
Dim Licznik As Integer, Szer() As Single
Private Obiekcik As Object, Forma(15) As Formaty, Item(15, 2) As Elementy
Public px As Single, py As Single
Dim IleKol As Long

Public Sub Rysuj(obiekt As Object)
Dim i As Integer, wid As Single, widt As Integer, l As Long, j As Long
Dim Kolumna As Long, Wiersz As Long, kolorek As ColorConstants
Dim Szerok() As Single, DFKolor As ColorConstants, DBKolor As ColorConstants
Dim kolorekT As ColorConstants
ReDim Szerok(IleKol)
Set Obiekcik = obiekt
DFKolor = obiekt.ForeColor
DBKolor = obiekt.BackColor
'On Error Resume Next
widt = 0
For Kolumna = 0 To IleKol
    If Kolumna > 0 Then Szerok(Kolumna) = Szerok(Kolumna - 1) + Szer(Kolumna)
    widt = widt + Szer(Kolumna)
    heigt = 0
    For Wiersz = 0 To Licznik
        If Wiersz > 0 Then heigt = heigt + (obiekt.TextHeight(Item(Wiersz - 1, 0).Text))
        obiekt.FontBold = Forma(Wiersz).FontBold
        obiekt.FontSize = Forma(Wiersz).FontSize
        obiekt.FontItalic = Forma(Wiersz).FontItalic
        obiekt.FontUnderline = Forma(Wiersz).FontUnderline
        If Forma(Wiersz).FontName = "" Then
            obiekt.Font = "Times New Roman CE"
        Else
            obiekt.Font = Forma(Wiersz).FontName
        End If
        obiekt.Font.Charset = 238
        If Not Item(Wiersz, Kolumna).ForeColor = 0 Then obiekt.ForeColor = Item(Wiersz, Kolumna).ForeColor
        If Item(Wiersz, Kolumna).BackColor = 0 Then
            kolorekT = vbWhite
        Else
            kolorekT = Item(Wiersz, Kolumna).BackColor
        End If
        
        
        
        obiekt.Line (px + widt - Szer(Kolumna), py + heigt + (obiekt.TextHeight(Item(Wiersz, Kolumna).Text)))-(px + widt, py + heigt), kolorekT, BF
        
        obiekt.CurrentX = 2 + px + widt - Szer(Kolumna)
        obiekt.CurrentY = py + heigt
      '  If Item(Wiersz, Kolumna).ForeColor = 0 Then
       '     Obiekt.ForeColor = vbBlack
       ' Else
            obiekt.ForeColor = Item(Wiersz, Kolumna).ForeColor
        'End If
        obiekt.Print (Item(Wiersz, Kolumna).Text)
    Next Wiersz
Next Kolumna
End Sub

Public Sub Clear()
Dim i As Long
On Error Resume Next
'For k = 0 To 15
'    For l = 0 To 2
'        Item(k, l).Text = ""
'    Next l
'Next k
For j = 1 To 2
    For t = 1 To 15
        Item(t, j).Text = ""
    Next t
Next j
Licznik = 0
IleKol = 0
Obiekcik.Cls
End Sub

Public Sub Remove(Index As Integer)

'If Index > Licznik Or Index < 0 Then
'    Exit Sub
'Else
'    For i = Index To Licznik - 1
'        Item(i).Wyraz = Item(i + 1).Wyraz
'        Item(i).Punkty = Item(i + 1).Punkty
'    Next i
'End If
'Licznik = Licznik - 1
'ReDim Preserve Item(Licznik)
End Sub

Private Sub Class_Initialize()
For i = 0 To 15
    Forma(i).FontSize = 14
Next i
Licznik = 0
IleKol = 0
End Sub
Public Sub AddColumn(NazwaKolumny As String, SzerokoœæKolumny As Single, Optional Align As AlignConstants)
Dim k As Long
If Align = 0 Then Align = vbAlignLeft
For k = 0 To 15
    Item(k, IleKol).Align = Align
Next k
IleKol = IleKol + 1
ReDim Preserve Szer(IleKol)
Szer(IleKol) = SzerokoœæKolumny
Item(0, IleKol).Text = NazwaKolumny
End Sub

Public Property Get Count() As Long
Count = Licznik
End Property

Public Sub AddItem(Wiersz As Single, Kolumna As Single, Text As String, Optional BackColor As ColorConstants = vbWhite, Optional ForeColor As ColorConstants = vbBlack)
If Kolumna > IleKol Then Exit Sub
If Wiersz > Licznik Then Licznik = Wiersz
Item(Wiersz, Kolumna).ForeColor = ForeColor
If Not BackColor = 0 Then Item(Wiersz, Kolumna).BackColor = BackColor
Item(Wiersz, Kolumna).Text = Text
End Sub
Public Sub Formatuj(Wiersz As Long, Optional FontName As String, Optional FontSize As Integer = 14, Optional FontBold As Boolean, Optional FontItalic As Boolean = False, Optional FontUnderline As Boolean = False)
 Forma(Wiersz).FontUnderline = FontUnderline
 Forma(Wiersz).FontItalic = FontItalic
 If Not FontName = "" Then Forma(Wiersz).FontName = FontName
 If Not FontSize = 0 Then Forma(Wiersz).FontSize = FontSize
 Forma(Wiersz).FontBold = FontBold
End Sub
