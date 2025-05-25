VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Tekst As String
Public eBackColor As ColorConstants, eForeColor As ColorConstants, eAlignH As AlignConstants
Public eFontSize As Integer, eFontItalic As Boolean, eFontBold As Boolean, eFontUnderline As Boolean
Public vCs As Single, hCs As Single, eAlignV As AlignConstants, eFontName As String
Public Szer As Single, Obrazek As Boolean, Imag As StdPicture
Public Sub Drukuj(px As Single, py As Single, wys As Single, ByVal obiekt As Object)
Dim pxAL As Single, pyAL As Single
    
If Obrazek = False Then
   obiekt.Font = eFontName
   obiekt.FontBold = eFontBold
   If eFontSize Then obiekt.FontSize = eFontSize
   obiekt.FontItalic = eFontItalic
   obiekt.FontUnderline = eFontUnderline
   obiekt.ForeColor = eForeColor
   obiekt.Font.Charset = 238
   If eBackColor > -1 Then obiekt.Line (px, py)-(px + Szer, py + wys), eBackColor, BF
    
   Select Case eAlignH
        
   Case vbAlignNone:
      pxAL = px + Szer / 2 - obiekt.TextWidth(Tekst) / 2
   Case vbAlignLeft
      pxAL = px + hCs
   Case vbAlignRight
      pxAL = px + Szer - obiekt.TextWidth(Tekst) - hCs
   End Select
    
   Select Case eAlignV
    
   Case vbAlignTop:
      pyAL = py + vCs
   Case vbAlignBottom
      pyAL = py + wys - obiekt.TextHeight(Tekst) - vCs
   Case vbAlignNone
      pyAL = py + wys / 2 - obiekt.TextHeight(Tekst) / 2
   End Select
   
   obiekt.CurrentX = pxAL
   obiekt.CurrentY = pyAL
   obiekt.Print Tekst
Else

   Select Case eAlignH
        
   Case vbAlignNone:
      pxAL = px + Szer / 2 - obiekt.TextWidth(Tekst) / 2
   Case vbAlignLeft
      pxAL = px + hCs
   Case vbAlignRight
      pxAL = px + Szer - obiekt.TextWidth(Tekst) - hCs
   End Select
    
   Select Case eAlignV
    
   Case vbAlignTop:
      pyAL = py + vCs
   Case vbAlignBottom
      pyAL = py + wys - obiekt.TextHeight(Tekst) - vCs
   Case vbAlignNone
      pyAL = py + wys / 2 - obiekt.TextHeight(Tekst) / 2
   End Select
   
   graslow.PaintPicture Imag, pxAL, pyAL, Szer, wys
End If
End Sub






