Attribute VB_Name = "modAutoval"
Option Explicit

Public Function CheckAutoComments(ByVal SampleID As String, ByVal Index As Integer) As String

          Dim tb As Recordset
          Dim sql As String
          Dim ShortDisc As String
          Dim Discipline As String
          Dim RetVal As String

57600     On Error GoTo CheckAutoComments_Error

57610     RetVal = ""

57620     If Index = 2 Then
57630         Discipline = "Biochemistry"
57640         ShortDisc = "Bio"
57650     Else
57660         Discipline = "Coagulation"
57670         ShortDisc = "Coag"
57680     End If

57690     sql = "SELECT 'Output' = " & _
                "CASE WHEN ISNUMERIC(R.Result) = 1 AND R.Result <> '.' " & _
              "  THEN " & _
              "    CASE " & _
              "      WHEN Criteria = 'Present' THEN A.Comment " & _
              "      WHEN Criteria = 'Equal to' AND CONVERT(float, R.Result) = CONVERT(float, A.Value0) THEN A.Comment " & _
              "      WHEN Criteria = 'Less than' AND CONVERT(float, R.Result) < CONVERT(float, A.Value0) THEN A.Comment " & _
              "      WHEN Criteria = 'Greater than' AND CONVERT(float, R.Result) > CONVERT(float, A.Value0) THEN A.Comment " & _
              "      WHEN Criteria = 'Between' AND CONVERT(float, R.Result) > CONVERT(float, A.Value0) AND CONVERT(float, R.Result) < CONVERT(float, A.Value1) THEN A.Comment " & _
              "      WHEN Criteria = 'Not between' AND (CONVERT(float, R.Result) < CONVERT(float, A.Value0) OR CONVERT(float, R.Result) > CONVERT(float, A.Value1)) THEN A.Comment " & _
              "      ELSE '' " & _
              "    END " & _
              "  ELSE " & _
              "    CASE " & _
              "      WHEN Criteria = 'Contains Text' AND CHARINDEX( A.Value0, R.Result) > 0 THEN A.Comment " & _
              "      WHEN Criteria = 'Starts with' AND LEFT(R.Result, 1) = A.Value0 THEN A.Comment " & _
              "      ELSE '' " & _
              "    END " & _
                "END "
57700     sql = sql & "FROM AutoComments A JOIN " & ShortDisc & "Results R ON " & _
                "R.Code = (SELECT TOP 1 Code FROM " & ShortDisc & "TestDefinitions " & _
              "          WHERE ShortName = A.Parameter " & _
              "          AND InUse = 1 ) " & _
                "WHERE A.Discipline = '" & Discipline & "' " & _
                "AND R.SampleID = '" & SampleID & "'"

57710     Set tb = New Recordset
57720     RecOpenClient 0, tb, sql
57730     Do While Not tb.EOF
57740         If Trim$(tb!Output & "") <> "" Then
57750             RetVal = RetVal & tb!Output & vbCrLf
57760         End If
57770         tb.MoveNext
57780     Loop

57790     CheckAutoComments = Trim$(RetVal)

57800     Exit Function

CheckAutoComments_Error:

          Dim strES As String
          Dim intEL As Integer

57810     intEL = Erl
57820     strES = Err.Description
57830     LogError "modAutoval", "CheckAutoComments", intEL, strES, sql

End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckAutoCommentsMicro
' Author    : Masood
' Date      : 01/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CheckAutoCommentsMicro(ByVal SampleID As String) As String
57840     On Error GoTo CheckAutoCommentsMicro_Error
          Dim tb As New ADODB.Recordset
          Dim Qry As String
          Dim Coments As String

57850     Qry = ""
57860     Qry = Qry & " SELECT     D.DoB, D.Ward, D.DateTimeDemographics, D.Age, ISNULL(SD.Site,'') AS Site, I.OrganismGroup, I.OrganismName, I.IsolateNumber" & vbNewLine
57870     Qry = Qry & " FROM         demographics AS D INNER JOIN Isolates AS I ON D.SampleID = I.SampleID INNER JOIN SiteDetails50 AS SD ON D.SampleID = SD.SampleID" & vbNewLine
57880     Qry = Qry & " WHERE     D.SampleID = '" & SampleID & "'" & vbNewLine


57890     Set tb = New Recordset
57900     RecOpenServer 0, tb, Qry
57910     Do While Not tb.EOF
57920         Coments = FindMicroComents(tb!OrganismName, tb!Site, tb!Ward, Format(tb!DoB, "yyyy-mm-dd"), Format(tb!DateTimeDemographics, "yyyy-mm-dd"), SampleID)
57930         If Coments <> "" Then
57940             CheckAutoCommentsMicro = IIf((CheckAutoCommentsMicro = ""), "", CheckAutoCommentsMicro & vbCrLf) & Coments
57950         End If
57960         tb.MoveNext
57970     Loop
57980     CheckAutoCommentsMicro = Trim(CheckAutoCommentsMicro)

57990     Exit Function


CheckAutoCommentsMicro_Error:

          Dim strES As String
          Dim intEL As Integer

58000     intEL = Erl
58010     strES = Err.Description
58020     LogError "modAutoval", "CheckAutoCommentsMicro", intEL, strES, Qry
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckMicroCom
' Author    : Masood
' Date      : 06/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FindMicroComents(OrganismName As String, Site As String, PatientLocation As String, DoB As String, SampleDate As String, SampleID As String)

          Dim sql As String
          Dim AgeinDays As Double
          Dim rs As New ADODB.Recordset
58030     On Error GoTo CheckMicroCom_Error

          '    AgeinDays = CalcAgeToDays(Val(PatientAge), 0, 0)
      '20        AgeinDays = DateDiff("d", DoB, SampleDate)
58040     AgeinDays = DateDiff("d", Format(DoB, "dd/mmm/yyyy hh:mm:ss"), Format(SampleDate, "dd/mmm/yyyy hh:mm:ss"))
58050     FindMicroComents = ""
58060     sql = "SELECT * from MicroAutoCommentAlert " & vbNewLine
58070     sql = sql & " Where " & vbNewLine
58080     sql = sql & " ISNULL(OrganismName,'" & OrganismName & "') = '" & OrganismName & "'" & vbNewLine
58090     sql = sql & " AND ISNULL(Site,'" & Site & "') = '" & Site & "'" & vbNewLine
58100     sql = sql & " AND ISNULL(PatientLocation,'" & PatientLocation & "') = '" & PatientLocation & "'" & vbNewLine

          '    sql = sql & " AND ( ISNULL(PatientAgeFrom," & Val(PatientAge) & ") <= " & Val(PatientAge) & "" & vbNewLine
          '    sql = sql & " AND ISNULL(PatientAgeTo," & Val(PatientAge) & ") >= " & Val(PatientAge) & " )" & vbNewLine

58110     sql = sql & " AND ( ISNULL(PatientAgeFrom," & Val(AgeinDays) & ") <= " & Val(AgeinDays) & "" & vbNewLine
58120     sql = sql & " AND ISNULL(PatientAgeTo," & Val(AgeinDays) & ") >= " & Val(AgeinDays) & " )" & vbNewLine

58130     sql = sql & " AND ISNULL(DateStart,'" & Format(SampleDate, "yyyy-mm-dd") & "') <= '" & Format(SampleDate, "yyyy-mm-dd") & "'" & vbNewLine
58140     sql = sql & " AND ISNULL(DateEnd,'" & Format(SampleDate, "yyyy-mm-dd") & "') >= '" & Format(SampleDate, "yyyy-mm-dd") & "'" & vbNewLine

58150     Set rs = New Recordset
58160     RecOpenServer 0, rs, sql
58170     Do While Not rs.EOF
58180         FindMicroComents = IIf((FindMicroComents = ""), "", FindMicroComents & vbCrLf) & rs!Comment
58190         If rs!PhoneAlert = True Then
58200             Call CheckIfMustPhoneMicro(SampleID, "Microbiology", Site, rs!PhoneAlertDateTime & "")
58210         End If
58220         rs.MoveNext
58230     Loop

58240     Exit Function


CheckMicroCom_Error:

          Dim strES As String
          Dim intEL As Integer

58250     intEL = Erl
58260     strES = Err.Description
58270     LogError "modAutoval", "CheckMicroCom", intEL, strES, sql

End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckIfMustPhoneMicro
' Author    : Masood
' Date      : 07/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CheckIfMustPhoneMicro(ByVal SampleID As String, ByVal Discipline As String, ByVal Site As String, ByVal PhoneAlertDateTime As String)

          Dim tb As Recordset
          Dim sql As String

58280     On Error GoTo CheckIfMustPhoneMicro_Error


58290     sql = "IF EXISTS (SELECT * FROM PhoneLog " & _
              "           WHERE SampleID = '" & SampleID & "' " & _
              "           AND Discipline LIKE '%" & Left$(Discipline, 1) & "%') " & _
              "  DELETE FROM PhoneAlert WHERE " & _
              "  SampleID = '" & SampleID & "' " & _
              "  AND Discipline = '" & Discipline & "' " & _
                "ELSE " & _
              "  IF NOT EXISTS (SELECT SampleID FROM PhoneAlert WHERE " & _
              "                 Discipline = '" & Discipline & "' " & _
              "                 AND Parameter = '" & Site & "' " & _
              "                 AND SampleID = '" & SampleID & "') " & _
              "    INSERT INTO PhoneAlert " & _
              "    (SampleID, Discipline, Parameter,PhoneAlertDateTime) VALUES " & _
              "    ('" & SampleID & "', " & _
              "     '" & Discipline & " ', " & _
              "     '" & Site & "'," & _
              "     '" & IIf((PhoneAlertDateTime = ""), Now, Format(PhoneAlertDateTime, "YYYY-MM-DD")) & "'" & _
                ")"
58300     Cnxn(0).Execute sql



58310     Exit Sub


CheckIfMustPhoneMicro_Error:

          Dim strES As String
          Dim intEL As Integer

58320     intEL = Erl
58330     strES = Err.Description
58340     LogError "modAutoval", "CheckIfMustPhoneMicro", intEL, strES, sql

End Sub





Public Function BuildAutoValSQL() As String

          Dim sql As String
          Dim n As Integer
          Dim Parameter As String

58350     On Error GoTo BuildAutoValSQL_Error

58360     sql = ""
58370     For n = 1 To 24
58380         Parameter = Choose(n, "WBC", "RBC", "Hct", "Hgb", "MCH", "MCHC", "MCV", _
                                 "Plt", "MPV", "PDW", "PLCR", _
                                 "BasA", "BasP", "EosA", "EosP", _
                                 "MonoA", "MonoP", "NeutA", "NeutP", "LymA", "LymP", _
                                 "RDWCV", "RDWSD", "Ret")
58390         sql = sql & BuildInsert(Parameter)
58400     Next

58410     BuildAutoValSQL = sql

58420     Exit Function

BuildAutoValSQL_Error:

          Dim strES As String
          Dim intEL As Integer

58430     intEL = Erl
58440     strES = Err.Description
58450     LogError "modAutoval", "BuildAutoValSQL", intEL, strES, sql


End Function

Private Function BuildInsert(ByVal Parameter As String) As String

          Dim sql As String

58460     On Error GoTo BuildInsert_Error

58470     sql = "INSERT INTO #TempAutoVal " & _
                "SELECT '" & Parameter & "', SampleID, " & _
                "CASE ISNUMERIC(" & Parameter & ") " & _
              "  WHEN 1 THEN " & Parameter & " " & _
              "  ELSE CAST(0 AS real) " & _
                "END FROM HaemResults WHERE " & _
                "SampleID IN (SELECT SampleID FROM HaemResults WHERE " & _
              "             COALESCE(Valid, 0) = 0 " & _
              "             AND RunDate BETWEEN DATEADD(day, -1, getdate()) " & _
              "             AND getdate() ) "

58480     BuildInsert = sql

58490     Exit Function

BuildInsert_Error:

          Dim strES As String
          Dim intEL As Integer

58500     intEL = Erl
58510     strES = Err.Description
58520     LogError "modAutoval", "BuildInsert", intEL, strES, sql


End Function


Public Function BuildSelectAutoValSQL() As String

          Dim sql As String

58530     On Error GoTo BuildSelectAutoValSQL_Error

58540     sql = "SELECT DISTINCT SampleID ,'Failure' AutoVal " & _
                "FROM #TempAutoVal R JOIN HaemAutoVal A " & _
                "ON A.Parameter COLLATE DATABASE_DEFAULT = R.Parameter COLLATE DATABASE_DEFAULT " & _
                "WHERE A.Include = 1 " & _
                "AND ( V  < Low OR V  > High OR V = 0) " & _
                "UNION " & _
                "SELECT DISTINCT SampleID ,'Pass' AutoVal FROM HaemResults WHERE " & _
                "COALESCE(Valid, 0) = 0 " & _
                "AND RunDate BETWEEN DATEADD(day, -1, getdate()) AND getdate() " & _
                "AND SampleID  NOT IN ( " & _
              "    SELECT DISTINCT SampleID FROM #TempAutoVal R JOIN HaemAutoVal A " & _
              "    ON A.Parameter COLLATE DATABASE_DEFAULT = R.Parameter COLLATE DATABASE_DEFAULT " & _
              "    WHERE A.Include = 1 " & _
              "    AND ( V < Low OR V > High OR V = 0 ) )"

58550     BuildSelectAutoValSQL = sql

58560     Exit Function

BuildSelectAutoValSQL_Error:

          Dim strES As String
          Dim intEL As Integer

58570     intEL = Erl
58580     strES = Err.Description
58590     LogError "modAutoval", "BuildSelectAutoValSQL", intEL, strES, sql


End Function

Public Sub CheckAutoVal()

          Dim sql As String
          Dim n As Integer
          Dim Parameter As String

58600     On Error GoTo CheckAutoVal_Error

58610     sql = "IF OBJECT_ID('tempdb..#TempAutoVal') IS NOT NULL  DROP TABLE #TempAutoVal " & _
              " CREATE Table #TempAutoVal (Parameter nvarchar(50), " & _
              "                               SampleID nvarchar(50), " & _
              "                               V real ) "

58620     For n = 1 To 24
58630         Parameter = Choose(n, "WBC", "RBC", "Hct", "Hgb", "MCH", "MCHC", "MCV", _
                                 "Plt", "MPV", "PDW", "PLCR", _
                                 "BasA", "BasP", "EosA", "EosP", _
                                 "MonoA", "MonoP", "NeutA", "NeutP", "LymA", "LymP", _
                                 "RDWCV", "RDWSD", "RetA")
58640         sql = sql & BuildInsert(Parameter)
58650     Next
58660     Cnxn(0).Execute sql

58670     Exit Sub

CheckAutoVal_Error:

          Dim strES As String
          Dim intEL As Integer

58680     intEL = Erl
58690     strES = Err.Description
58700     LogError "modAutoval", "CheckAutoVal", intEL, strES, sql

End Sub


