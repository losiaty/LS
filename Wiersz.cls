VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wiersz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mElementy As New Collection
Public wys As Single

Public Sub Drukuj(X As Single, Y As Single, obiekt As Object)

Dim e As Element, CX As Single
CX = X

For Each e In mElementy
    e.Drukuj CX, Y, wys, obiekt
    CX = CX + e.Szer
Next e

End Sub

Public Function AddElement(Optional szerokoœæ As Single, Optional Nazwa As String) As Element

Dim e As Element

Set e = New Element
e.Szer = szerokoœæ
If Nazwa = "" Then
    mElementy.Add e
Else
    mElementy.Add e, Nazwa
End If
Set AddElement = e

End Function

Public Function AddNewElement(Optional Tekst As String, Optional szerokoœæ As Single, Optional horizontalCellSpacing As Single = 1, Optional verticalCellSpacing As Single = 0, Optional Nazwa As String, Optional eBackColor As ColorConstants = vbWhite, Optional eForeColor As ColorConstants = vbBlack, Optional eAlignH As AlignConstants = vbAlignLeft, Optional eAlignV As AlignConstants = vbAlignTop, Optional eFontName As String = "Times New Roman CE", Optional eFontSize As Integer = 10, Optional eFontItalic As Boolean, Optional eFontBold As Boolean, Optional eFontUnderline As Boolean) As Element
Dim e As Element
Set e = New Element
e.eAlignH = eAlignH
e.eFontName = eFontName
e.eAlignV = eAlignV
e.vCs = verticalCellSpacing
e.hCs = horizontalCellSpacing
e.eBackColor = eBackColor
e.eFontBold = eFontBold
e.eFontItalic = eFontItalic
e.eFontSize = eFontSize
e.eFontUnderline = eFontUnderline
e.eForeColor = eForeColor
e.Szer = szerokoœæ
e.Tekst = Tekst

If Nazwa = "" Then
    mElementy.Add e
Else
    mElementy.Add e, Nazwa
End If
Set AddNewElement = e

End Function
