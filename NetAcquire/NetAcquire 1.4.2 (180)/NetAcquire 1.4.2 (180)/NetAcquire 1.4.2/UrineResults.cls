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

27810     mcolUrineResults.Add Ux, Ux.TestName

27820     Set Add = Ux

End Function

Public Sub CheckForAllResults()

          Dim Ux As UrineResult
          Dim UxNew As UrineResult
          Dim Found As Boolean
          Dim Result As String
          Dim Added As Integer

27830     On Error GoTo CheckForAllResults_Error

27840     Found = False

27850     For Each Ux In mcolUrineResults
27860         If Ux.TestName = "WBC" Then
27870             For Each UxNew In mcolUrineResults
27880                 If UxNew.TestName = "WCC" Then
27890                     Found = True
27900                     Exit For
27910                 End If
27920             Next
27930             If Not Found Then
27940                 Set UxNew = New UrineResult
27950                 UxNew.Result = GetWBCValue(Ux.Result)
27960                 UxNew.SampleID = Ux.SampleID
27970                 UxNew.TestName = "WCC"
27980                 UrineResults.Save UxNew
27990                 UrineResults.Add UxNew
28000                 Exit For
28010             End If
28020         End If
28030     Next

28040     For Each Ux In mcolUrineResults
28050         If Ux.TestName = "RBC" Then
28060             For Each UxNew In mcolUrineResults
28070                 If UxNew.TestName = "RCC" Then
28080                     Found = True
28090                     Exit For
28100                 End If
28110             Next
28120             If Not Found Then
28130                 Set UxNew = New UrineResult
28140                 UxNew.Result = GetPlussesOrNil(Ux.Result)
28150                 UxNew.SampleID = Ux.SampleID
28160                 UxNew.TestName = "RCC"
28170                 UrineResults.Save UxNew
28180                 UrineResults.Add UxNew
28190                 Exit For
28200             End If
28210         End If
28220     Next

28230     For Each Ux In mcolUrineResults
28240         If Ux.TestName = "BAC" Then
28250             For Each UxNew In mcolUrineResults
28260                 If UxNew.TestName = "Bacteria" Then
28270                     Found = True
28280                     Exit For
28290                 End If
28300             Next
28310             If Not Found Then
28320                 Set UxNew = New UrineResult
28330                 UxNew.Result = GetPlussesOrNil(Ux.Result)
28340                 UxNew.SampleID = Ux.SampleID
28350                 UxNew.TestName = "Bacteria"
28360                 UrineResults.Save UxNew
28370                 UrineResults.Add UxNew
28380                 Exit For
28390             End If
28400         End If
28410     Next

28420     For Each Ux In mcolUrineResults
28430         If Ux.TestName = "CRY" Then
28440             For Each UxNew In mcolUrineResults
28450                 If UxNew.TestName = "Crystals" Then
28460                     Found = True
28470                     Exit For
28480                 End If
28490             Next
28500             If Not Found Then
28510                 Result = GetPlusses(Ux.Result)
28520                 If Result <> "" Then
28530                     Set UxNew = New UrineResult
28540                     UxNew.Result = Result
28550                     UxNew.SampleID = Ux.SampleID
28560                     UxNew.TestName = "Crystals"
28570                     UrineResults.Save UxNew
28580                     UrineResults.Add UxNew
28590                     Exit For
28600                 End If
28610             End If
28620         End If
28630     Next

28640     For Each Ux In mcolUrineResults
28650         If Ux.TestName = "HYA" Then
28660             For Each UxNew In mcolUrineResults
28670                 If InStr(UxNew.TestName, "Casts") > 0 Then
28680                     Found = True
28690                     Exit For
28700                 End If
28710             Next
28720             If Not Found Then
28730                 Result = GetPlusses(Ux.Result)
28740                 If Result <> "" Then
28750                     Set UxNew = New UrineResult
28760                     UxNew.Result = "Casts - Hyalin " & Result
28770                     UxNew.SampleID = Ux.SampleID
28780                     UxNew.TestName = "Casts"
28790                     UrineResults.Save UxNew
28800                     UrineResults.Add UxNew
28810                     Exit For
28820                 End If
28830             End If
28840         End If
28850     Next

28860     For Each Ux In mcolUrineResults
28870         If Ux.TestName = "PAT" Then
28880             For Each UxNew In mcolUrineResults
28890                 If InStr(UxNew.TestName, "Casts") > 0 Then
28900                     Found = True
28910                     Exit For
28920                 End If
28930             Next
28940             If Not Found Then
28950                 Result = GetPlusses(Ux.Result)
28960                 If Result <> "" Then
28970                     Set UxNew = New UrineResult
28980                     UxNew.Result = "Casts - Pathological " & Result
28990                     UxNew.SampleID = Ux.SampleID
29000                     UxNew.TestName = "Casts"
29010                     UrineResults.Save UxNew
29020                     UrineResults.Add UxNew
29030                     Exit For
29040                 End If
29050             End If
29060         End If
29070     Next

29080     Added = -1

29090     For Each Ux In mcolUrineResults
29100         If Ux.TestName = "EPI" Then
29110             For Each UxNew In mcolUrineResults
29120                 If InStr(UxNew.Result, "Epith") > 0 Then
29130                     Found = True
29140                     Exit For
29150                 End If
29160             Next
29170             If Not Found Then
29180                 Result = GetPlusses(Ux.Result)
29190                 If Result <> "" Then
29200                     Set UxNew = New UrineResult
29210                     UxNew.Result = "Epithelial Cells " & Result
29220                     UxNew.SampleID = Ux.SampleID
29230                     Added = Added + 1
29240                     UxNew.TestName = "Misc" & Format(Added)
29250                     UrineResults.Save UxNew
29260                     UrineResults.Add UxNew
29270                     Exit For
29280                 End If
29290             End If
29300         End If
29310     Next

29320     For Each Ux In mcolUrineResults
29330         If Ux.TestName = "YEA" Then
29340             For Each UxNew In mcolUrineResults
29350                 If InStr(UxNew.Result, "Yeasts") > 0 Then
29360                     Found = True
29370                     Exit For
29380                 End If
29390             Next
29400             If Not Found Then
29410                 Result = GetPlusses(Ux.Result)
29420                 If Result <> "" Then
29430                     Set UxNew = New UrineResult
29440                     UxNew.Result = "Yeasts " & Result
29450                     UxNew.SampleID = Ux.SampleID
29460                     Added = Added + 1
29470                     UxNew.TestName = "Misc" & Format(Added)
29480                     UrineResults.Save UxNew
29490                     UrineResults.Add UxNew
29500                     Exit For
29510                 End If
29520             End If
29530         End If
29540     Next

29550     Exit Sub

CheckForAllResults_Error:

          Dim strES As String
          Dim intEL As Integer

29560     intEL = Erl
29570     strES = Err.Description
29580     LogError "UrineResults", "CheckForAllResults", intEL, strES

End Sub

Public Function Count() As Long
        
29590     Count = mcolUrineResults.Count

End Function





Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
        
          Dim sql As String

29600     On Error GoTo Delete_Error

29610     sql = "DELETE FROM UrineResults50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND TestName = '" & TestName & "'"
29620     Cnxn(0).Execute sql

29630     mcolUrineResults.Remove TestName

29640     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

29650     intEL = Erl
29660     strES = Err.Description
29670     LogError "UrineResults", "Delete", intEL, strES, sql

End Sub





Public Property Get UrineResults() As UrineResults
        
29680     Set UrineResults = mUrineResults

End Property



Public Function Item(ByVal TestName As String) As UrineResult
Attribute Item.VB_UserMemId = 0
        
29690     On Error Resume Next

29700     Set Item = mcolUrineResults.Item(TestName)

End Function





Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ux As UrineResult

29710     On Error GoTo Load_Error

29720     sql = "SELECT G.SampleID, G.TestName, G.Result, COALESCE(G.HealthLink, 0) Healthlink, " & _
              "G.UserName, G.DateTimeOfRecord, COALESCE(P.Valid, 0) Valid, COALESCE(P.Printed, 0) Printed, " & _
              "P.ValidatedBy, P.ValidatedDateTime, P.PrintedBy, P.PrintedDateTime " & _
              "FROM UrineResults50 G LEFT JOIN PrintValidLog P " & _
              "ON G.SampleID = P.SampleID " & _
              "WHERE G.SampleID = '" & SampleID & "'"
29730     Set tb = New Recordset
29740     RecOpenClient 0, tb, sql
29750     Do While Not tb.EOF
29760         Set Ux = New UrineResult
29770         With Ux
29780             .SampleID = tb!SampleID & ""
29790             .TestName = tb!TestName & ""
29800             .Result = tb!Result & ""
29810             .HealthLink = tb!HealthLink
29820             .Valid = tb!Valid
29830             .Printed = tb!Printed
29840             .ValidatedBy = tb!ValidatedBy & ""
29850             .ValidatedDateTime = tb!ValidatedDateTime & ""
29860             .PrintedBy = tb!PrintedBy & ""
29870             .PrintedDateTime = tb!PrintedDateTime & ""
29880             .UserName = tb!UserName & ""
29890             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
29900             mcolUrineResults.Add Ux, .TestName
29910         End With
29920         tb.MoveNext
29930     Loop


29940     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

29950     intEL = Erl
29960     strES = Err.Description
29970     LogError "UrineResults", "Load", intEL, strES, sql


End Sub

Public Sub LoadSedimax(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ux As UrineResult
          Dim SedimexResultsExist As Boolean

29980     On Error GoTo LoadSedimax_Error

29990     sql = "SELECT SampleID, Code TestName, Result " & _
              "FROM SedimaxResults " & _
              "WHERE SampleID = '" & SampleID & "'"
30000     Set tb = New Recordset
30010     RecOpenClient 0, tb, sql
30020     SedimexResultsExist = False
30030     Do While Not tb.EOF
30040         Set Ux = New UrineResult
30050         With Ux
30060             .SampleID = tb!SampleID & ""
30070             .TestName = tb!TestName & ""
30080             .Result = tb!Result & ""
30090             mcolUrineResults.Add Ux, .TestName
30100         End With
30110         SedimexResultsExist = True
30120         tb.MoveNext
30130     Loop


30140     Exit Sub

LoadSedimax_Error:

          Dim strES As String
          Dim intEL As Integer

30150     intEL = Erl
30160     strES = Err.Description
30170     LogError "UrineResults", "LoadSedimax", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
30180     Set NewEnum = mcolUrineResults.[_NewEnum]

End Function


Public Sub Save(ByVal Fx As UrineResult)

          Dim sql As String
          Dim SID As Long

30190     On Error GoTo Save_Error

          '20    If Val(Fx.SampleID) < sysOptMicroOffset(0) Then
          '30      SID = Val(Fx.SampleID) + sysOptMicroOffset(0)
          '40    Else
30200     SID = Val(Fx.SampleID)
          '60    End If

30210     If Trim$(Fx.Result) = "" Then
              '80      Sql = "Delete from UrineResults50 where " & _
              '              "SampleID = '" & SID & "' " & _
              '              "and TestName = '" & Fx.TestName & "'"
              '        Exit Sub
30220     Else
30230         sql = "IF EXISTS (SELECT * FROM UrineResults50 WHERE " & _
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
30240         Cnxn(0).Execute sql
30250     End If


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
          '      Cnxn(0).Execute Sql

30260     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

30270     intEL = Erl
30280     strES = Err.Description
30290     LogError "UrineResults", "Save", intEL, strES, sql

End Sub


