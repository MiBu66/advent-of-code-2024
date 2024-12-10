Attribute VB_Name = "Day02"
Option Explicit
Dim Daten As String
Dim DatArr

Sub AdventDay02a()
Dim i As Integer, Report
Dim Summe As Long
  Daten = HoleDaten(1, True)  'Lese die Vorgabedaten
  DatArr = Split(Daten, vbCrLf)  'Zerlege die Daten zeilenweise
  Summe = 0
  For i = 0 To UBound(DatArr)
    Report = Split(DatArr(i), " ")   'zerlege in die einzelnen Nummern
    If CheckReport(Report) Then
      Summe = Summe + 1
    End If
  Next i
  MsgBox Summe
End Sub

Sub AdventDay02b()
Dim i As Integer, Report, QReport
Dim Summe As Long, GlobalSafe As Boolean
Dim p As Integer
  Daten = HoleDaten(1, True)  'Lese die Vorgabedaten
  DatArr = Split(Daten, vbCrLf)  'Zerlege die Daten zeilenweise
  Summe = 0
  For i = 0 To UBound(DatArr)
    Report = Split(DatArr(i), " ")   'zerlege in die einzelnen Nummern
    If CheckReport(Report) Then      'Prüfe, ob Vorgaben erfüllt werden
      GlobalSafe = True              'wenn ja, dann ist es sicher
    Else                             'wenn nicht
      GlobalSafe = False
      For p = 0 To UBound(Report)
        QReport = Report
        ArrayRemoveElement QReport, p   'entferne einfach nacheinander ein Element nach dem nächsten (brute force)
        If CheckReport(QReport) Then    'wenn's dann funktioniert, dann ist es sicher
          GlobalSafe = True
          Exit For
        End If
      Next p
    End If
    If GlobalSafe Then
      Summe = Summe + 1
    End If
  Next i
  MsgBox Summe
End Sub

Sub ArrayRemoveElement(ByRef Arr, ByVal ElementNr)
Dim i As Integer
  For i = ElementNr To UBound(Arr) - 1
    Arr(i) = Arr(i + 1)
  Next i
  ReDim Preserve Arr(UBound(Arr) - 1)
End Sub

Function CheckReport(RepIn) As Boolean              'prüft, ob auf- oder absteigend und Sprung >1 und <3
Dim l As Integer, safe As Boolean, q As Integer
  l = UBound(RepIn)
  safe = True
  If CInt(RepIn(0)) > CInt(RepIn(l)) Then           'absteigende Folge
    For q = 0 To l - 1
      If CInt(RepIn(q)) - CInt(RepIn(q + 1)) < 1 Or CInt(RepIn(q)) - CInt(RepIn(q + 1)) > 3 Then
        safe = False
        Exit For
      End If
    Next q
  Else                                               'aufsteigende Folge
    For q = 0 To l - 1
      If CInt(RepIn(q + 1)) - CInt(RepIn(q)) < 1 Or CInt(RepIn(q + 1)) - CInt(RepIn(q)) > 3 Then
        safe = False
        Exit For
      End If
    Next q
  End If
  CheckReport = safe
End Function
