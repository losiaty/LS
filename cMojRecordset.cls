VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMojRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private IleRek As Long, AktRek As Long, Rekordy() As String

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

Public Sub Add(Wyraz As String)
IleRek = IleRek + 1
ReDim Preserve Rekordy(IleRek)
AktRek = IleRek
Rekordy(IleRek) = Wyraz
End Sub

Public Sub Clear()
Erase Rekordy
AktRek = 0
IleRek = 0

End Sub

Public Function Element() As String
   Element = Rekordy(AktRek)
End Function
