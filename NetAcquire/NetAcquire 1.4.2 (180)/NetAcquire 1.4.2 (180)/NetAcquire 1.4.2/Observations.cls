VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Observations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolObservations As New Collection
Private mObservations As New Observations

Public Sub Clear()

          Dim n As Long

22360     For n = 1 To mcolObservations.Count
22370         mcolObservations.Remove 1
22380     Next

End Sub

Public Property Get Observations() As Observations
        
22390     Set Observations = mObservations

End Property


Public Function Item(ByVal X As Long) _
          As Observation

22400     Set Item = mcolObservations(X)

End Function



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
22410     Set NewEnum = mcolObservations.[_NewEnum]

End Function




Public Sub Remove(ByVal Index As Integer)

22420     mcolObservations.Remove (Index)

End Sub


Public Function Load(ByVal SampleID As String, ParamArray Disc()) As Observations
          'Paramarray is one or a combination of :-
          'Demographic Biochemistry Haematology Coagulation Immunology BloodGas
          'Semen Microcs MicroIdent MicroGeneral MicroConsultant Film
          'Endocrinology Histology Cytology CSFFluid
          'ImmunologyA ImmunologyB ImmunologyC

          Dim OBs As New Observations
          Dim OB As Observation
          Dim sql As String
          Dim tb As New Recordset
          Dim n As Integer
          Dim s As String

22430     On Error GoTo Load_Error

22440     sql = "SELECT * FROM Observations WHERE " & _
              "SampleID = '" & Val(SampleID) & "' AND ("
22450     s = ""
22460     For n = LBound(Disc) To UBound(Disc)
22470         s = s & "Discipline = '" & Disc(n) & "' OR "
22480     Next
22490     s = Left$(s, Len(s) - 3) & ")"
22500     sql = sql & s

22510     Set tb = New Recordset
22520     RecOpenClient 0, tb, sql

22530     Do While Not tb.EOF
22540         Set OB = New Observation
22550         With OB
22560             .SampleID = tb!SampleID & ""
22570             .Discipline = tb!Discipline & ""
22580             .Comment = tb!Comment & ""
22590             .DateTimeOfRecord = Format$(tb!DateTimeOfRecord, "dd/MM/yyyy HH:nn:ss")
22600             .UserName = tb!UserName & ""
22610             OBs.Add OB
22620         End With
22630         tb.MoveNext
22640     Loop

22650     If OBs.Count <> 0 Then
22660         Set Load = OBs
22670     Else
22680         Set Load = Nothing
22690     End If

22700     Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

22710     intEL = Erl
22720     strES = Err.Description
22730     LogError "Observations", "Load", intEL, strES, sql


End Function
Public Function Save(ByVal SampleID As String, ByVal OverWrite As Boolean, ParamArray Disc()) As Observations

          'OverWrite = True : Any existing comment is replaced.
          'OverWrite = False : New comment is concatenated

          'Paramarray is one or a combination of :-
          'Demographic Biochemistry Haematology Coagulation Immunology BloodGas
          'Semen Microcs MicroIdent MicroGeneral MicroConsultant Film
          'Endocrinology Histology Cytology CSFFluid
          'ImmunologyA ImmunologyB ImmunologyC
          'eg "Haematology","Haem Comment","Biochemistry","Bio Comment" etc

          Dim sql As String
          Dim tb As Recordset
          Dim tbD As Recordset
          Dim n As Integer
          Dim Discipline As String
          Dim Comment As String
          Dim l_Dis As String

22740     On Error GoTo Save_Error
22750     l_Dis = ""
22760     For n = LBound(Disc) To UBound(Disc) Step 2
22770         Discipline = Disc(n)
22780         Comment = AddTicks(Disc(n + 1))

              '+++ Junaid 22-05-2024
22790         If Discipline = "Biochemistry" Then
22800             l_Dis = "BioResults"
22810         ElseIf Discipline = "Coagulation" Then
22820             l_Dis = "CoagResults"
22830         ElseIf Discipline = "Haematology" Then
22840             l_Dis = "HaemResults"
22850         End If
22860         If l_Dis <> "" Then
22870             sql = "Select * from " & l_Dis & " Where SampleID = '" & SampleID & "'"
22880             Set tbD = New Recordset
22890             RecOpenClient 0, tbD, sql
22900             If Not tbD Is Nothing Then
22910                 If Not tbD.EOF Then
22920                 Else
22930                     sql = "DELETE FROM Observations " & _
                              "WHERE SampleID = '" & SampleID & "' " & _
                              "AND Discipline = '" & Discipline & "'"
22940                     Cnxn(0).Execute sql
22950                 End If
22960             Else
22970                 sql = "DELETE FROM Observations " & _
                          "WHERE SampleID = '" & SampleID & "' " & _
                          "AND Discipline = '" & Discipline & "'"
22980                 Cnxn(0).Execute sql
22990             End If
23000         End If
              '--- Junaid
              '50      If Comment = "" Then
              ''60          Sql = "DELETE FROM Observations " & _
              ''                  "WHERE SampleID = '" & SampleID & "' " & _
              ''                  "AND Discipline = '" & Discipline & "'"
              ''        Exit Function
              '
              '70      Else
23010         sql = "IF EXISTS (SELECT * FROM Observations " & _
                  "WHERE SampleID = '" & SampleID & "' " & _
                  "AND Discipline = '" & Discipline & "') " & _
                  "  UPDATE Observations "
23020         If OverWrite Then
23030             sql = sql & "  SET Comment = '" & Comment & "' ,UserName = '" & AddTicks(UserName) & "' ,DateTimeOfRecord = getdate()"
23040         Else
23050             sql = sql & "  SET Comment = Comment + ' " & Comment & "' ,UserName = '" & AddTicks(UserName) & "' ,DateTimeOfRecord = getdate()"
23060         End If
23070         sql = sql & "  WHERE SampleID = '" & SampleID & "' " & _
                  "  AND Discipline = '" & Discipline & "' " & _
                  "ELSE " & _
                  "  INSERT INTO Observations " & _
                  "  (SampleID, Discipline, Comment, UserName,DateTimeOfRecord ) " & _
                  "  VALUES " & _
                  "  ('" & SampleID & "', " & _
                  "   '" & Discipline & "', " & _
                  "   '" & Comment & "', " & _
                  "   '" & AddTicks(UserName) & "',getdate())"
23080         Cnxn(0).Execute sql
                
              '+++ Junaid 20-01-2024
23090         If Comment = "" Then
23100             sql = "Select * from Observations WHERE SampleID = '" & SampleID & "' AND Discipline = '" & Discipline & "'"
23110             Set tb = New Recordset
23120             RecOpenClient 0, tb, sql
23130             If Not tb Is Nothing Then
23140                 If Not tb.EOF Then
23150                     If ConvertNull(tb!Comment, "") = "" Then
23160                         sql = "DELETE FROM Observations " & _
                                  "WHERE SampleID = '" & SampleID & "' " & _
                                  "AND Discipline = '" & Discipline & "'"
23170                         Cnxn(0).Execute sql
23180                     End If
23190                 End If
23200             End If
23210         End If
              '--- Junaid
              '150     End If

23220     Next

23230     Exit Function

Save_Error:

          Dim strES As String
          Dim intEL As Integer

23240     intEL = Erl
23250     strES = Err.Description
23260     LogError "Observations", "Save", intEL, strES, sql


End Function
Public Function Count() As Long
        
23270     Count = mcolObservations.Count

23280     Exit Function

End Function



Public Function Add(OB As Observation) As Observation
Attribute Add.VB_MemberFlags = "40"

23290     mcolObservations.Add OB

23300     Set Add = OB

End Function


