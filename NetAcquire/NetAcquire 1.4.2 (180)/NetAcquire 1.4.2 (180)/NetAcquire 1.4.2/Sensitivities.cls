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

24240     On Error Resume Next

24250     mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode

End Sub


Public Function Count() As Long
        
24260     Count = mcolSensitivities.Count

End Function



Public Sub Delete(ByVal sx As Sensitivity)
        
          Dim sql As String

24270     On Error GoTo Delete_Error

24280     sql = "DELETE FROM Sensitivities WHERE " & _
              "SampleID = '" & sx.SampleID & "' " & _
              "AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "AND AntibioticCode = '" & sx.AntibioticCode & "'"
24290     Cnxn(0).Execute sql

24300     mcolSensitivities.Remove CStr(sx.IsolateNumber) & sx.AntibioticCode

24310     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

24320     intEL = Erl
24330     strES = Err.Description
24340     LogError "Sensitivities", "Delete", intEL, strES, sql


End Sub


Public Property Get Sensitivities() As Sensitivities
        
24350     Set Sensitivities = mSensitivities

End Property





Public Function Item(ByVal IsolateNumber As Integer, ByVal AntibioticCode As String) As Sensitivity
        
          Dim i As Integer

24360     On Error GoTo Item_Error

24370     Set Item = Nothing
24380     For i = 1 To mcolSensitivities.Count
24390         If mcolSensitivities.Item(i).IsolateNumber = IsolateNumber And mcolSensitivities.Item(i).AntibioticCode = AntibioticCode Then
24400             Set Item = mcolSensitivities(i)
24410         End If
24420     Next i


24430     Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

24440     intEL = Erl
24450     strES = Err.Description
24460     LogError "Sensitivities", "Item", intEL, strES
          
End Function




Public Sub Load(ByVal SampleIDWithOffset As String)

          Dim tb As Recordset
          Dim sql As String
          Dim sx As Sensitivity

24470     On Error GoTo Load_Error

24480     sql = "SELECT S.SampleID, S.IsolateNumber, LTRIM(RTRIM(S.AntibioticCode)) AntibioticCode, " & _
              "LTRIM(RTRIM(A.AntibioticName)) AntibioticName, LTRIM(RTRIM(COALESCE(S.Result, ''))) Result, " & _
              "COALESCE(S.Report, 0) Report, LTRIM(RTRIM(COALESCE(S.CPOFlag, ''))) CPOFlag, S.RunDate, " & _
              "S.RunDateTime, LTRIM(RTRIM(COALESCE(S.RSI, ''))) RSI, " & _
              "COALESCE(S.UserCode, '') UserCode, COALESCE(S.Forced, 0) Forced, COALESCE(S.Secondary, 0) Secondary, " & _
              "COALESCE(S.Valid, 0) Valid, S.AuthoriserCode, DateTimeOfRecord " & _
              "FROM Sensitivities S JOIN Antibiotics A " & _
              "ON S.AntibioticCode = A.Code " & _
              "WHERE S.SampleID = '" & SampleIDWithOffset & "' " & _
              "ORDER BY A.ListOrder"
24490     Set tb = New Recordset
24500     RecOpenClient 0, tb, sql
24510     Do While Not tb.EOF
24520         Set sx = New Sensitivity
24530         With sx
24540             .SampleID = SampleIDWithOffset
24550             .AntibioticCode = tb!AntibioticCode & ""
24560             .AntibioticName = tb!AntibioticName & ""
24570             .AuthoriserCode = tb!AuthoriserCode & ""
24580             .CPOFlag = tb!CPOFlag & ""
24590             .Forced = tb!Forced
24600             .IsolateNumber = tb!IsolateNumber
24610             .Report = tb!Report
24620             .Result = tb!Result
24630             .RSI = tb!RSI
24640             If Not IsNull(tb!Rundate) Then
24650                 .Rundate = tb!Rundate
24660             End If
24670             If Not IsNull(tb!RunDateTime) Then
24680                 .RunDateTime = tb!RunDateTime
24690             End If
24700             .Secondary = tb!Secondary
24710             .UserCode = tb!UserCode
24720             .Valid = tb!Valid
24730             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
24740             mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode
24750         End With
24760         tb.MoveNext
24770     Loop
          
24780     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

24790     intEL = Erl
24800     strES = Err.Description
24810     LogError "Sensitivities", "Load", intEL, strES, sql


End Sub

Public Sub LoadRepeats(ByVal SampleIDWithOffset As String)

          Dim tb As Recordset
          Dim sql As String
          Dim sx As Sensitivity

24820     On Error GoTo LoadRepeats_Error

24830     sql = "SELECT S.SampleID, S.IsolateNumber, LTRIM(RTRIM(S.AntibioticCode)) AntibioticCode, " & _
              "LTRIM(RTRIM(A.AntibioticName)) AntibioticName, LTRIM(RTRIM(COALESCE(S.Result, ''))) Result, " & _
              "COALESCE(S.Report, 0) Report, LTRIM(RTRIM(COALESCE(S.CPOFlag, ''))) CPOFlag, S.RunDate, " & _
              "S.RunDateTime, LTRIM(RTRIM(COALESCE(S.RSI, ''))) RSI, " & _
              "COALESCE(S.UserCode, '') UserCode, COALESCE(S.Forced, 0) Forced, COALESCE(S.Secondary, 0) Secondary, " & _
              "COALESCE(S.Valid, 0) Valid, S.AuthoriserCode, DateTimeOfRecord " & _
              "FROM SensitivitiesRepeats S JOIN Antibiotics A " & _
              "ON S.AntibioticCode = A.Code " & _
              "WHERE S.SampleID = '" & SampleIDWithOffset & "' " & _
              "ORDER BY A.ListOrder"
24840     Set tb = New Recordset
24850     RecOpenClient 0, tb, sql
24860     Do While Not tb.EOF
24870         Set sx = New Sensitivity
24880         With sx
24890             .SampleID = SampleIDWithOffset
24900             .AntibioticCode = tb!AntibioticCode & ""
24910             .AntibioticName = tb!AntibioticName & ""
24920             .AuthoriserCode = tb!AuthoriserCode & ""
24930             .CPOFlag = tb!CPOFlag & ""
24940             .Forced = tb!Forced
24950             .IsolateNumber = tb!IsolateNumber
24960             .Report = tb!Report
24970             .Result = tb!Result
24980             .RSI = tb!RSI
24990             If Not IsNull(tb!Rundate) Then
25000                 .Rundate = tb!Rundate
25010             End If
25020             If Not IsNull(tb!RunDateTime) Then
25030                 .RunDateTime = tb!RunDateTime
25040             End If
25050             .Secondary = tb!Secondary
25060             .UserCode = tb!UserCode
25070             .Valid = tb!Valid
25080             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
25090             mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode
25100         End With
25110         tb.MoveNext
25120     Loop

25130     Exit Sub

LoadRepeats_Error:

          Dim strES As String
          Dim intEL As Integer

25140     intEL = Erl
25150     strES = Err.Description
25160     LogError "Sensitivities", "LoadRepeats", intEL, strES, sql

End Sub

Public Sub LoadArchive(ByVal SampleIDWithOffset As String)

          Dim tb As Recordset
          Dim sql As String
          Dim sx As Sensitivity

25170     On Error GoTo LoadArchive_Error

25180     sql = "SELECT S.SampleID, S.IsolateNumber, LTRIM(RTRIM(S.AntibioticCode)) AntibioticCode, " & _
              "LTRIM(RTRIM(A.AntibioticName)) AntibioticName, LTRIM(RTRIM(COALESCE(S.Result, ''))) Result, " & _
              "COALESCE(S.Report, 0) Report, LTRIM(RTRIM(COALESCE(S.CPOFlag, ''))) CPOFlag, S.RunDate, " & _
              "S.RunDateTime, LTRIM(RTRIM(COALESCE(S.RSI, ''))) RSI, " & _
              "COALESCE(S.UserCode, '') UserCode, COALESCE(S.Forced, 0) Forced, COALESCE(S.Secondary, 0) Secondary, " & _
              "COALESCE(S.Valid, 0) Valid, S.AuthoriserCode, DateTimeOfRecord, S.ArchivedBy " & _
              "FROM SensitivitiesArc S JOIN Antibiotics A " & _
              "ON S.AntibioticCode = A.Code " & _
              "WHERE S.SampleID = '" & SampleIDWithOffset & "' " & _
              "ORDER BY A.ListOrder"
25190     Set tb = New Recordset
25200     RecOpenClient 0, tb, sql
25210     Do While Not tb.EOF
25220         Set sx = New Sensitivity
25230         With sx
25240             .SampleID = SampleIDWithOffset
25250             .AntibioticCode = tb!AntibioticCode & ""
25260             .AntibioticName = tb!AntibioticName & ""
25270             .AuthoriserCode = tb!AuthoriserCode & ""
25280             .CPOFlag = tb!CPOFlag & ""
25290             .Forced = tb!Forced
25300             .IsolateNumber = tb!IsolateNumber
25310             .Report = tb!Report
25320             .Result = tb!Result
25330             .RSI = tb!RSI
25340             If Not IsNull(tb!Rundate) Then
25350                 .Rundate = tb!Rundate
25360             End If
25370             If Not IsNull(tb!RunDateTime) Then
25380                 .RunDateTime = tb!RunDateTime
25390             End If
25400             .Secondary = tb!Secondary
25410             .UserCode = tb!UserCode
25420             .Valid = tb!Valid
25430             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
25440             .ArchivedBy = tb!ArchivedBy & ""
25450             .ArchiveDateTime = Format(tb!ArchiveDateTime, "dd/MM/yy HH:mm")
25460             mcolSensitivities.Add sx, CStr(sx.IsolateNumber) & sx.AntibioticCode
25470         End With
25480         tb.MoveNext
25490     Loop

25500     Exit Sub

LoadArchive_Error:

          Dim strES As String
          Dim intEL As Integer

25510     intEL = Erl
25520     strES = Err.Description
25530     LogError "Sensitivities", "LoadArchive", intEL, strES, sql

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
25540     Set NewEnum = mcolSensitivities.[_NewEnum]

End Function




Public Sub Save()

          Dim sql As String
          Dim sx As Sensitivity

25550     On Error GoTo Save_Error

25560     For Each sx In mcolSensitivities
        
25570         sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
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
                  "  AuthoriserCode = '" & sx.AuthoriserCode & "', DateTimeOfRecord = GetDate() " & _
                  "  WHERE SampleID = '" & sx.SampleID & "' " & _
                  "  AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
                  "  AND AntibioticCode = '" & sx.AntibioticCode & "' "
25580         sql = sql & "ELSE " & _
                  "  INSERT INTO Sensitivities " & _
                  "  (SampleID, IsolateNumber, AntibioticCode, Result, Report, CPOFlag, RunDate, " & _
                  "  RunDateTime, RSI, UserCode, Forced, Secondary, Valid, AuthoriserCode, DateTimeOfRecord) " & _
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
                  "  '" & sx.AuthoriserCode & "', GetDate())"

25590         Cnxn(0).Execute sql

25600     Next

25610     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

25620     intEL = Erl
25630     strES = Err.Description
25640     LogError "Sensitivities", "Save", intEL, strES, sql

End Sub


Public Sub ForceUnForce(ByVal sx As Sensitivity, ByVal Force As Integer)

          Dim sql As String

25650     On Error GoTo ForceUnForce_Error

25660     sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
              "           SampleID = '" & sx.SampleID & "' " & _
              "           AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "           AND AntibioticCode = '" & sx.AntibioticCode & "') " & _
              "  UPDATE Sensitivities " & _
              "  SET Forced = '" & Force & "' " & _
              "  WHERE SampleID = '" & sx.SampleID & "' " & _
              "  AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "  AND AntibioticCode = '" & sx.AntibioticCode & "' "

25670     Cnxn(0).Execute sql

25680     Exit Sub

ForceUnForce_Error:

          Dim strES As String
          Dim intEL As Integer

25690     intEL = Erl
25700     strES = Err.Description
25710     LogError "Sensitivities", "ForceUnForce", intEL, strES, sql

End Sub

Public Sub SetSecondary(ByVal sx As Sensitivity, ByVal Secondary As Integer)

          Dim sql As String

25720     On Error GoTo SetSecondary_Error

25730     sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
              "           SampleID = '" & sx.SampleID & "' " & _
              "           AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "           AND AntibioticCode = '" & sx.AntibioticCode & "') " & _
              "  UPDATE Sensitivities " & _
              "  SET Secondary = '" & Secondary & "' " & _
              "  WHERE SampleID = '" & sx.SampleID & "' " & _
              "  AND IsolateNumber = '" & sx.IsolateNumber & "' " & _
              "  AND AntibioticCode = '" & sx.AntibioticCode & "' "

25740     Cnxn(0).Execute sql

25750     Exit Sub

SetSecondary_Error:

          Dim strES As String
          Dim intEL As Integer

25760     intEL = Erl
25770     strES = Err.Description
25780     LogError "Sensitivities", "SetSecondary", intEL, strES, sql

End Sub



