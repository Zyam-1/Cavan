VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UrineResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolUrineResults As New Collection

Private mUrineResults As New UrineResults


Public Function Add(Ux As UrineResult) As UrineResult

10    mcolUrineResults.Add Ux, Ux.TestName

20    Set Add = Ux

End Function

Public Sub CheckForAllResults()

      Dim Ux As UrineResult
      Dim UxNew As UrineResult
      Dim Found As Boolean
      Dim Result As String
      Dim Added As Integer

10    On Error GoTo CheckForAllResults_Error

20    Found = False

30    For Each Ux In mcolUrineResults
40      If Ux.TestName = "WBC" Then
50        For Each UxNew In mcolUrineResults
60          If UxNew.TestName = "WCC" Then
70            Found = True
80            Exit For
90          End If
100       Next
110       If Not Found Then
120         Set UxNew = New UrineResult
130         UxNew.Result = GetWBCValue(Ux.Result)
140         UxNew.SampleID = Ux.SampleID
150         UxNew.TestName = "WCC"
160         UrineResults.Save UxNew
170         UrineResults.Add UxNew
180         Exit For
190       End If
200     End If
210   Next

220   For Each Ux In mcolUrineResults
230     If Ux.TestName = "RBC" Then
240       For Each UxNew In mcolUrineResults
250         If UxNew.TestName = "RCC" Then
260           Found = True
270           Exit For
280         End If
290       Next
300       If Not Found Then
310         Set UxNew = New UrineResult
320         UxNew.Result = GetPlussesOrNil(Ux.Result)
330         UxNew.SampleID = Ux.SampleID
340         UxNew.TestName = "RCC"
350         UrineResults.Save UxNew
360         UrineResults.Add UxNew
370         Exit For
380       End If
390     End If
400   Next

410   For Each Ux In mcolUrineResults
420     If Ux.TestName = "BAC" Then
430       For Each UxNew In mcolUrineResults
440         If UxNew.TestName = "Bacteria" Then
450           Found = True
460           Exit For
470         End If
480       Next
490       If Not Found Then
500         Set UxNew = New UrineResult
510         UxNew.Result = GetPlussesOrNil(Ux.Result)
520         UxNew.SampleID = Ux.SampleID
530         UxNew.TestName = "Bacteria"
540         UrineResults.Save UxNew
550         UrineResults.Add UxNew
560         Exit For
570       End If
580     End If
590   Next

600   For Each Ux In mcolUrineResults
610     If Ux.TestName = "CRY" Then
620       For Each UxNew In mcolUrineResults
630         If UxNew.TestName = "Crystals" Then
640           Found = True
650           Exit For
660         End If
670       Next
680       If Not Found Then
690         Result = GetPlusses(Ux.Result)
700         If Result <> "" Then
710           Set UxNew = New UrineResult
720           UxNew.Result = Result
730           UxNew.SampleID = Ux.SampleID
740           UxNew.TestName = "Crystals"
750           UrineResults.Save UxNew
760           UrineResults.Add UxNew
770           Exit For
780         End If
790       End If
800     End If
810   Next

820   For Each Ux In mcolUrineResults
830     If Ux.TestName = "HYA" Then
840       For Each UxNew In mcolUrineResults
850         If InStr(UxNew.TestName, "Casts") > 0 Then
860           Found = True
870           Exit For
880         End If
890       Next
900       If Not Found Then
910         Result = GetPlusses(Ux.Result)
920         If Result <> "" Then
930           Set UxNew = New UrineResult
940           UxNew.Result = "Casts - Hyalin " & Result
950           UxNew.SampleID = Ux.SampleID
960           UxNew.TestName = "Casts"
970           UrineResults.Save UxNew
980           UrineResults.Add UxNew
990           Exit For
1000        End If
1010      End If
1020    End If
1030  Next

1040  For Each Ux In mcolUrineResults
1050    If Ux.TestName = "PAT" Then
1060      For Each UxNew In mcolUrineResults
1070        If InStr(UxNew.TestName, "Casts") > 0 Then
1080          Found = True
1090          Exit For
1100        End If
1110      Next
1120      If Not Found Then
1130        Result = GetPlusses(Ux.Result)
1140        If Result <> "" Then
1150          Set UxNew = New UrineResult
1160          UxNew.Result = "Casts - Pathological " & Result
1170          UxNew.SampleID = Ux.SampleID
1180          UxNew.TestName = "Casts"
1190          UrineResults.Save UxNew
1200          UrineResults.Add UxNew
1210          Exit For
1220        End If
1230      End If
1240    End If
1250  Next

1260  Added = -1

1270  For Each Ux In mcolUrineResults
1280    If Ux.TestName = "EPI" Then
1290      For Each UxNew In mcolUrineResults
1300        If InStr(UxNew.Result, "Epith") > 0 Then
1310          Found = True
1320          Exit For
1330        End If
1340      Next
1350      If Not Found Then
1360        Result = GetPlusses(Ux.Result)
1370        If Result <> "" Then
1380          Set UxNew = New UrineResult
1390          UxNew.Result = "Epithelial Cells " & Result
1400          UxNew.SampleID = Ux.SampleID
1410          Added = Added + 1
1420          UxNew.TestName = "Misc" & Format(Added)
1430          UrineResults.Save UxNew
1440          UrineResults.Add UxNew
1450          Exit For
1460        End If
1470      End If
1480    End If
1490  Next

1500  For Each Ux In mcolUrineResults
1510    If Ux.TestName = "YEA" Then
1520      For Each UxNew In mcolUrineResults
1530        If InStr(UxNew.Result, "Yeasts") > 0 Then
1540          Found = True
1550          Exit For
1560        End If
1570      Next
1580      If Not Found Then
1590        Result = GetPlusses(Ux.Result)
1600        If Result <> "" Then
1610          Set UxNew = New UrineResult
1620          UxNew.Result = "Yeasts " & Result
1630          UxNew.SampleID = Ux.SampleID
1640          Added = Added + 1
1650          UxNew.TestName = "Misc" & Format(Added)
1660          UrineResults.Save UxNew
1670          UrineResults.Add UxNew
1680          Exit For
1690        End If
1700      End If
1710    End If
1720  Next

1730  Exit Sub

CheckForAllResults_Error:

Dim strES As String
Dim intEL As Integer

1740  intEL = Erl
1750  strES = Err.Description
1760  LogError "UrineResults", "CheckForAllResults", intEL, strES

End Sub

Public Function Count() As Long
  
10    Count = mcolUrineResults.Count

End Function





Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM UrineResults50 WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND TestName = '" & TestName & "'"
30    Cnxn(0).Execute sql

40    mcolUrineResults.Remove TestName

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "UrineResults", "Delete", intEL, strES, sql

End Sub





Public Property Get UrineResults() As UrineResults
  
10    Set UrineResults = mUrineResults

End Property



Public Function Item(ByVal TestName As String) As UrineResult
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolUrineResults.Item(TestName)

End Function





Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Ux As UrineResult

10    On Error GoTo Load_Error

20    sql = "SELECT G.SampleID, G.TestName, G.Result, COALESCE(G.HealthLink, 0) Healthlink, " & _
            "G.UserName, G.DateTimeOfRecord, COALESCE(P.Valid, 0) Valid, COALESCE(P.Printed, 0) Printed, " & _
            "P.ValidatedBy, P.ValidatedDateTime, P.PrintedBy, P.PrintedDateTime " & _
            "FROM UrineResults50 G LEFT JOIN PrintValidLog P " & _
            "ON G.SampleID = P.SampleID " & _
            "WHERE G.SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Ux = New UrineResult
70      With Ux
80        .SampleID = tb!SampleID & ""
90        .TestName = tb!TestName & ""
100       .Result = tb!Result & ""
110       .HealthLink = tb!HealthLink
120       .Valid = tb!Valid
130       .Printed = tb!Printed
140       .ValidatedBy = tb!ValidatedBy & ""
150       .ValidatedDateTime = tb!ValidatedDateTime & ""
160       .PrintedBy = tb!PrintedBy & ""
170       .PrintedDateTime = tb!PrintedDateTime & ""
180       .UserName = tb!UserName & ""
190       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
200       mcolUrineResults.Add Ux, .TestName
210     End With
220     tb.MoveNext
230   Loop


240   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "UrineResults", "Load", intEL, strES, sql


End Sub

Public Sub LoadSedimax(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Ux As UrineResult
      Dim SedimexResultsExist As Boolean

10    On Error GoTo LoadSedimax_Error

20    sql = "SELECT SampleID, Code TestName, Result " & _
            "FROM SedimaxResults " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    SedimexResultsExist = False
60    Do While Not tb.EOF
70      Set Ux = New UrineResult
80      With Ux
90        .SampleID = tb!SampleID & ""
100       .TestName = tb!TestName & ""
110       .Result = tb!Result & ""
120       mcolUrineResults.Add Ux, .TestName
130     End With
140     SedimexResultsExist = True
150     tb.MoveNext
160   Loop


170   Exit Sub

LoadSedimax_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "UrineResults", "LoadSedimax", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolUrineResults.[_NewEnum]

End Function


Public Sub Save(ByVal Fx As UrineResult)

      Dim sql As String
      Dim SID As Long

10    On Error GoTo Save_Error

20    If Val(Fx.SampleID) < sysOptMicroOffset(0) Then
30      SID = Val(Fx.SampleID) + sysOptMicroOffset(0)
40    Else
50      SID = Val(Fx.SampleID)
60    End If

70    If Trim$(Fx.Result) = "" Then
80      sql = "Delete from UrineResults50 where " & _
              "SampleID = '" & SID & "' " & _
              "and TestName = '" & Fx.TestName & "'"
90    Else
100     sql = "IF EXISTS (SELECT * FROM UrineResults50 WHERE " & _
              "           SampleID = '" & SID & "' " & _
              "           AND TestName = '" & Fx.TestName & "') " & _
              "  UPDATE UrineResults50 " & _
              "  SET Result = '" & Fx.Result & "', " & _
              "  UserName = '" & AddTicks(Fx.UserName) & "', " & _
              "  HealthLink = '" & Fx.HealthLink & "' " & _
              "  WHERE SampleID = '" & SID & "' " & _
              "  AND TestName = '" & Fx.TestName & "' " & _
              "ELSE " & _
              "  INSERT INTO UrineResults50 (SampleID, TestName, Result, UserName, HealthLink) " & _
              "  VALUES " & _
              "  ('" & SID & "', " & _
              "  '" & Fx.TestName & "', " & _
              "  '" & Fx.Result & "', " & _
              "  '" & AddTicks(Fx.UserName) & "', " & _
              "  '" & Fx.HealthLink & "')"
110   End If
120   Cnxn(0).Execute sql

      '      sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
      '            "           SampleID = '" & SID & "' " & _
      '            "           AND Department = 'M') " & _
      '            "  UPDATE PrintValidLog " & _
      '            "  SET Valid = '" & Fx.Valid & "', " & _
      '            "  Printed = '" & Fx.Printed & "', " & _
      '            "  ValidatedBy = '" & AddTicks(Fx.ValidatedBy) & "', " & _
      '            "  PrintedBy = '" & AddTicks(Fx.PrintedBy) & "', " & _
      '            "  ValidatedDateTime = '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
      '            "  PrintedDateTime = '" & Format$(Fx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "' " & _
      '            "  WHERE SampleID = '" & SID & "' " & _
      '            "  AND Department = 'M' " & _
      '            "ELSE " & _
      '            "  INSERT INTO PrintValidLog " & _
      '            "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
      '            "  VALUES " & _
      '            "  ('" & SID & "', 'M', '" & Fx.Printed & "', '" & Fx.Valid & "', " & _
      '            "  '" & AddTicks(Fx.PrintedBy) & "', '" & Format$(Fx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
      '            "  '" & AddTicks(Fx.ValidatedBy) & "', '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
      '      Cnxn(0).Execute sql

130      Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "UrineResults", "Save", intEL, strES, sql

End Sub




