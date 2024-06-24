VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sensitivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolSensitivities As New Collection

Private mSensitivities As New Sensitivities

Public Sub Add(ByVal sx As Sensitivity)

10    On Error Resume Next

20    mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode

End Sub


Public Function Count() As Long
  
10    Count = mcolSensitivities.Count

End Function



Public Sub Delete(ByVal sx As Sensitivity)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM Sensitivities WHERE " & _
            "SampleID = '" & sx.SampleID & "' " & _
            "AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
            "AND AntibioticCode = '" & sx.AntibioticCode & "'"
30    Cnxn(0).Execute sql

40    mcolSensitivities.Remove CStr(sx.IsolateNumber) & sx.AntibioticCode

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "Sensitivities", "Delete", intEL, strES, sql


End Sub


Public Property Get Sensitivities() As Sensitivities
  
10    Set Sensitivities = mSensitivities

End Property





Public Function Item(ByVal IsolateNumber As Integer, ByVal AntibioticCode As String) As Sensitivity
        
      Dim I As Integer

10    On Error GoTo Item_Error

20    Set Item = Nothing
30    For I = 1 To mcolSensitivities.Count
40        If mcolSensitivities.Item(I).IsolateNumber = IsolateNumber And mcolSensitivities.Item(I).AntibioticCode = AntibioticCode Then
50            Set Item = mcolSensitivities(I)
60        End If
70    Next I


80    Exit Function

Item_Error:

       Dim strES As String
       Dim intEL As Integer

90     intEL = Erl
100    strES = Err.Description
110    LogError "Sensitivities", "Item", intEL, strES
          
End Function




Public Sub Load(ByVal SampleIDWithOffset As String)

      Dim tb As Recordset
      Dim sql As String
      Dim sx As Sensitivity

10    On Error GoTo Load_Error

20    sql = "SELECT S.SampleID, S.IsolateNumber, LTRIM(RTRIM(S.AntibioticCode)) AntibioticCode, " & _
            "LTRIM(RTRIM(A.AntibioticName)) AntibioticName, LTRIM(RTRIM(COALESCE(S.Result, ''))) Result, " & _
            "COALESCE(S.Report, 0) Report, LTRIM(RTRIM(COALESCE(S.CPOFlag, ''))) CPOFlag, S.RunDate, " & _
            "S.RunDateTime, LTRIM(RTRIM(COALESCE(S.RSI, ''))) RSI, " & _
            "COALESCE(S.UserCode, '') UserCode, COALESCE(S.Forced, 0) Forced, COALESCE(S.Secondary, 0) Secondary, " & _
            "COALESCE(S.Valid, 0) Valid, S.AuthoriserCode, DateTimeOfRecord " & _
            "FROM Sensitivities S JOIN Antibiotics A " & _
            "ON S.AntibioticCode = A.Code " & _
            "WHERE S.SampleID = '" & SampleIDWithOffset & "' " & _
            "ORDER BY A.ListOrder"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set sx = New Sensitivity
70      With sx
80        .SampleID = SampleIDWithOffset
90        .AntibioticCode = tb!AntibioticCode & ""
100       .AntibioticName = tb!AntibioticName & ""
110       .AuthoriserCode = tb!AuthoriserCode & ""
120       .CPOFlag = tb!CPOFlag & ""
130       .Forced = tb!Forced
140       .IsolateNumber = tb!IsolateNumber
150       .Report = tb!Report
160       .Result = tb!Result
170       .RSI = tb!RSI
180       If Not IsNull(tb!Rundate) Then
190         .Rundate = tb!Rundate
200       End If
210       If Not IsNull(tb!RunDateTime) Then
220         .RunDateTime = tb!RunDateTime
230       End If
240       .Secondary = tb!Secondary
250       .UserCode = tb!UserCode
260       .Valid = tb!Valid
270       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
280       mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode
290     End With
300     tb.MoveNext
310   Loop
    
320   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "Sensitivities", "Load", intEL, strES, sql


End Sub

Public Sub LoadRepeats(ByVal SampleIDWithOffset As String)

      Dim tb As Recordset
      Dim sql As String
      Dim sx As Sensitivity

10    On Error GoTo LoadRepeats_Error

20    sql = "SELECT S.SampleID, S.IsolateNumber, LTRIM(RTRIM(S.AntibioticCode)) AntibioticCode, " & _
            "LTRIM(RTRIM(A.AntibioticName)) AntibioticName, LTRIM(RTRIM(COALESCE(S.Result, ''))) Result, " & _
            "COALESCE(S.Report, 0) Report, LTRIM(RTRIM(COALESCE(S.CPOFlag, ''))) CPOFlag, S.RunDate, " & _
            "S.RunDateTime, LTRIM(RTRIM(COALESCE(S.RSI, ''))) RSI, " & _
            "COALESCE(S.UserCode, '') UserCode, COALESCE(S.Forced, 0) Forced, COALESCE(S.Secondary, 0) Secondary, " & _
            "COALESCE(S.Valid, 0) Valid, S.AuthoriserCode, DateTimeOfRecord " & _
            "FROM SensitivitiesRepeats S JOIN Antibiotics A " & _
            "ON S.AntibioticCode = A.Code " & _
            "WHERE S.SampleID = '" & SampleIDWithOffset & "' " & _
            "ORDER BY A.ListOrder"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set sx = New Sensitivity
70      With sx
80        .SampleID = SampleIDWithOffset
90        .AntibioticCode = tb!AntibioticCode & ""
100       .AntibioticName = tb!AntibioticName & ""
110       .AuthoriserCode = tb!AuthoriserCode & ""
120       .CPOFlag = tb!CPOFlag & ""
130       .Forced = tb!Forced
140       .IsolateNumber = tb!IsolateNumber
150       .Report = tb!Report
160       .Result = tb!Result
170       .RSI = tb!RSI
180       If Not IsNull(tb!Rundate) Then
190         .Rundate = tb!Rundate
200       End If
210       If Not IsNull(tb!RunDateTime) Then
220         .RunDateTime = tb!RunDateTime
230       End If
240       .Secondary = tb!Secondary
250       .UserCode = tb!UserCode
260       .Valid = tb!Valid
270       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
280       mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode
290     End With
300     tb.MoveNext
310   Loop

320   Exit Sub

LoadRepeats_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "Sensitivities", "LoadRepeats", intEL, strES, sql

End Sub

Public Sub LoadArchive(ByVal SampleIDWithOffset As String)

      Dim tb As Recordset
      Dim sql As String
      Dim sx As Sensitivity

10    On Error GoTo LoadArchive_Error

20    sql = "SELECT S.SampleID, S.IsolateNumber, LTRIM(RTRIM(S.AntibioticCode)) AntibioticCode, " & _
            "LTRIM(RTRIM(A.AntibioticName)) AntibioticName, LTRIM(RTRIM(COALESCE(S.Result, ''))) Result, " & _
            "COALESCE(S.Report, 0) Report, LTRIM(RTRIM(COALESCE(S.CPOFlag, ''))) CPOFlag, S.RunDate, " & _
            "S.RunDateTime, LTRIM(RTRIM(COALESCE(S.RSI, ''))) RSI, " & _
            "COALESCE(S.UserCode, '') UserCode, COALESCE(S.Forced, 0) Forced, COALESCE(S.Secondary, 0) Secondary, " & _
            "COALESCE(S.Valid, 0) Valid, S.AuthoriserCode, DateTimeOfRecord " & _
            "FROM SensitivitiesArc S JOIN Antibiotics A " & _
            "ON S.AntibioticCode = A.Code " & _
            "WHERE S.SampleID = '" & SampleIDWithOffset & "' " & _
            "ORDER BY A.ListOrder"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set sx = New Sensitivity
70      With sx
80        .SampleID = SampleIDWithOffset
90        .AntibioticCode = tb!AntibioticCode & ""
100       .AntibioticName = tb!AntibioticName & ""
110       .AuthoriserCode = tb!AuthoriserCode & ""
120       .CPOFlag = tb!CPOFlag & ""
130       .Forced = tb!Forced
140       .IsolateNumber = tb!IsolateNumber
150       .Report = tb!Report
160       .Result = tb!Result
170       .RSI = tb!RSI
180       If Not IsNull(tb!Rundate) Then
190         .Rundate = tb!Rundate
200       End If
210       If Not IsNull(tb!RunDateTime) Then
220         .RunDateTime = tb!RunDateTime
230       End If
240       .Secondary = tb!Secondary
250       .UserCode = tb!UserCode
260       .Valid = tb!Valid
270       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
280       .ArchivedBy = tb!ArchivedBy & ""
290       .ArchiveDateTime = Format(tb!ArchiveDateTime, "dd/MM/yy HH:mm")
300       mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode
310     End With
320     tb.MoveNext
330   Loop

340   Exit Sub

LoadArchive_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "Sensitivities", "LoadArchive", intEL, strES, sql

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolSensitivities.[_NewEnum]

End Function




Public Sub Save()

      Dim sql As String
      Dim sx As Sensitivity

10    On Error GoTo Save_Error

20    For Each sx In mcolSensitivities
  
30      sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
              "           SampleID = '" & sx.SampleID & "' " & _
              "           AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "           AND AntibioticCode = '" & sx.AntibioticCode & "') " & _
              "  UPDATE Sensitivities " & _
              "  SET Result = '" & sx.Result & "', " & _
              "  Report = '" & sx.Report & "', " & _
              "  CPOFlag = '" & sx.CPOFlag & "', " & _
              "  RunDate = '" & Format$(sx.Rundate, "dd/MMM/yyyy") & "', " & _
              "  RunDateTime = getdate(), " & _
              "  RSI = '" & sx.RSI & "', " & _
              "  UserCode = '" & sx.UserCode & "', " & _
              "  Forced = '" & sx.Forced & "', " & _
              "  Secondary = '" & sx.Secondary & "', " & _
              "  Valid = '" & sx.Valid & "', " & _
              "  AuthoriserCode = '" & sx.AuthoriserCode & "' " & _
              "  WHERE SampleID = '" & sx.SampleID & "' " & _
              "  AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
        "  AND AntibioticCode = '" & sx.AntibioticCode & "' "
40      sql = sql & "ELSE " & _
              "  INSERT INTO Sensitivities " & _
              "  (SampleID, IsolateNumber, AntibioticCode, Result, Report, CPOFlag, RunDate, " & _
              "  RunDateTime, RSI, UserCode, Forced, Secondary, Valid, AuthoriserCode) " & _
              "  VALUES " & _
              "  ('" & sx.SampleID & "', " & _
              "  '" & sx.IsolateNumber & "', " & _
              "  '" & sx.AntibioticCode & "', " & _
              "  '" & sx.Result & "', " & _
              "  '" & sx.Report & "', " & _
              "  '" & sx.CPOFlag & "', " & _
              "  '" & Format$(sx.Rundate, "dd/MMM/yyyy") & "', " & _
              "  '" & Format$(sx.RunDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
              "  '" & sx.RSI & "', " & _
              "  '" & sx.UserCode & "', " & _
              "  '" & sx.Forced & "', " & _
              "  '" & sx.Secondary & "', " & _
              "  '" & sx.Valid & "', " & _
              "  '" & sx.AuthoriserCode & "')"

50      Cnxn(0).Execute sql

60    Next

70    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "Sensitivities", "Save", intEL, strES, sql

End Sub


Public Sub ForceUnForce(ByVal sx As Sensitivity, ByVal Force As Integer)

Dim sql As String

10    On Error GoTo ForceUnForce_Error

20      sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
              "           SampleID = '" & sx.SampleID & "' " & _
              "           AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "           AND AntibioticCode = '" & sx.AntibioticCode & "') " & _
              "  UPDATE Sensitivities " & _
              "  SET Forced = '" & Force & "' " & _
              "  WHERE SampleID = '" & sx.SampleID & "' " & _
              "  AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "  AND AntibioticCode = '" & sx.AntibioticCode & "' "

30      Cnxn(0).Execute sql

40    Exit Sub

ForceUnForce_Error:

Dim strES As String
Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "Sensitivities", "ForceUnForce", intEL, strES, sql

End Sub

Public Sub SetSecondary(ByVal sx As Sensitivity, ByVal Secondary As Integer)

Dim sql As String

10    On Error GoTo SetSecondary_Error

20      sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
              "           SampleID = '" & sx.SampleID & "' " & _
              "           AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "           AND AntibioticCode = '" & sx.AntibioticCode & "') " & _
              "  UPDATE Sensitivities " & _
              "  SET Secondary = '" & Secondary & "' " & _
              "  WHERE SampleID = '" & sx.SampleID & "' " & _
              "  AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "  AND AntibioticCode = '" & sx.AntibioticCode & "' "

30      Cnxn(0).Execute sql

40    Exit Sub

SetSecondary_Error:

Dim strES As String
Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "Sensitivities", "SetSecondary", intEL, strES, sql

End Sub


