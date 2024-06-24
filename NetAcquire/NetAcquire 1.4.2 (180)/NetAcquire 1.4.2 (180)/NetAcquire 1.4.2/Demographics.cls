VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Demographics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolDemographics As New Collection

Private mDemographics As New Demographics

Public Sub Add(ByVal dx As Demographic)

10400     mcolDemographics.Add dx, dx.SampleID

End Sub
Public Function Count() As Long
        
10410     Count = mcolDemographics.Count

End Function
Public Property Get Demographics() As Demographics
        
10420     Set Demographics = mDemographics

End Property

Public Function Item(ByVal Index As Long) As Demographic
Attribute Item.VB_UserMemId = 0
        
10430     On Error Resume Next

10440     Set Item = mcolDemographics.Item(Index)

End Function


Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String

10450     On Error GoTo Load_Error

10460     sql = "SELECT TOP 1 SampleID, Chart, PatName, Age, Sex, RunDate, " & _
              "DoB, " & _
              "Addr0, Addr1, Ward, Clinician, GP, SampleDate, ClDetails, Hospital, " & _
              "COALESCE(RooH, 1) RooH, COALESCE(FAXed, 0) Faxed, COALESCE(Fasting, 0) Fasting, " & _
              "DateTimeDemographics, DateTimeHaemPrinted, " & _
              "DateTimeBioPrinted, DateTimeCoagPrinted, " & _
              "COALESCE(Pregnant, 0) Pregnant, AandE, RecDate, " & _
              "RecordDateTime, [Operator], Category, Username, COALESCE(Urgent, 0) Urgent, " & _
              "COALESCE(Valid, 0) Valid, COALESCE(ForMicro, 0) ForMicro, " & _
              "COALESCE(SentToEMedRenal, 0) SentToEmedRenal, AssID, SurName, ForeName, ExtSampleID, COALESCE(Healthlink, 0) HealthLink " & _
              " ,LabNo " & _
              "FROM Demographics " & _
              "WHERE SampleID = '" & SampleID & "'"
10470     Set tb = New Recordset
10480     RecOpenClient 0, tb, sql
10490     Do While Not tb.EOF
10500         Populate tb
10510         tb.MoveNext
10520     Loop

10530     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

10540     intEL = Erl
10550     strES = Err.Description
10560     LogError "Demographics", "Load", intEL, strES, sql

End Sub
Public Sub LoadAllByEntryDate(ByVal DateOfInterest As String)

          Dim tb As Recordset
          Dim sql As String

10570     On Error GoTo LoadAllByEntryDate_Error

10580     sql = "SELECT SampleID, Chart, PatName, Age, Sex, RunDate, " & _
              "DoB, " & _
              "Addr0, Addr1, Ward, Clinician, GP, SampleDate, ClDetails, Hospital, " & _
              "COALESCE(RooH, 1) RooH, COALESCE(FAXed, 0) Faxed, COALESCE(Fasting, 0) Fasting, " & _
              "DateTimeDemographics, DateTimeHaemPrinted, " & _
              "DateTimeBioPrinted, DateTimeCoagPrinted, " & _
              "COALESCE(Pregnant, 0) Pregnant, AandE, RecDate, " & _
              "RecordDateTime, [Operator], Category, Username, COALESCE(Urgent, 0) Urgent, " & _
              "COALESCE(Valid, 0) Valid, COALESCE(ForMicro, 0) ForMicro, " & _
              "COALESCE(SentToEMedRenal, 0) SentToEmedRenal, AssID, SurName, ForeName, ExtSampleID, COALESCE(Healthlink, 0) HealthLink " & _
              " ,LabNo " & _
              "FROM Demographics " & _
              "WHERE DateTimeDemographics BETWEEN '" & Format$(DateOfInterest, "dd/MMM/yyyy") & "' " & _
              "                               AND '" & Format$(DateOfInterest, "dd/MMM/yyyy") & " 23:59'"
10590     Set tb = New Recordset
10600     RecOpenClient 0, tb, sql
10610     Do While Not tb.EOF
10620         Populate tb
10630         tb.MoveNext
10640     Loop

10650     Exit Sub

LoadAllByEntryDate_Error:

          Dim strES As String
          Dim intEL As Integer

10660     intEL = Erl
10670     strES = Err.Description
10680     LogError "Demographics", "LoadAllByEntryDate", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
10690     Set NewEnum = mcolDemographics.[_NewEnum]

End Function






Private Sub Populate(ByVal tb As Recordset)

10700     On Error GoTo Populate_Error

          Dim dx As New Demographic
10710     With dx
10720         .SampleID = tb!SampleID & ""
10730         .Chart = tb!Chart & ""
10740         .LabNo = Val(tb!LabNo & "")
10750         .PatName = tb!PatName & ""
10760         .Age = tb!Age & ""
10770         .Sex = tb!Sex & ""
10780         .DoB = tb!DoB & ""
10790         .Addr0 = tb!Addr0 & ""
10800         .Addr1 = tb!Addr1 & ""
10810         .Ward = tb!Ward & ""
10820         .Clinician = tb!Clinician & ""
10830         .GP = tb!GP & ""
10840         .ClDetails = tb!ClDetails & ""
10850         .Hospital = tb!Hospital & ""
10860         .RooH = tb!RooH
10870         .FAXed = tb!FAXed
10880         .Fasting = tb!Fasting
10890         .Pregnant = tb!Pregnant
10900         .Operator = tb!Operator & ""
10910         .Category = tb!Category & ""
10920         .UserName = tb!UserName & ""
10930         .Urgent = tb!Urgent
10940         .Valid = tb!Valid
10950         .ForMicro = tb!ForMicro
10960         .SentToEMedRenal = tb!SentToEMedRenal
10970         .AssID = tb!AssID & ""
10980         .SurName = tb!SurName & ""
10990         .ForeName = tb!ForeName & ""
11000         .ExtSampleID = tb!ExtSampleID & ""
11010         .HealthLink = tb!HealthLink
11020         .AandE = tb!AandE & ""
          
11030         .Rundate = tb!Rundate & ""
11040         .SampleDate = tb!SampleDate & ""
11050         .RecDate = tb!RecDate & ""
11060         .DateTimeDemographics = tb!DateTimeDemographics & ""
11070         .DateTimeHaemPrinted = tb!DateTimeHaemPrinted & ""
11080         .DateTimeBioPrinted = tb!DateTimeBioPrinted & ""
11090         .DateTimeCoagPrinted = tb!DateTimeCoagPrinted & ""
11100         .RecordDateTime = tb!RecordDateTime & ""
11110         mcolDemographics.Add dx
11120     End With

11130     Exit Sub

Populate_Error:

          Dim strES As String
          Dim intEL As Integer

11140     intEL = Erl
11150     strES = Err.Description
11160     LogError "Demographics", "Populate", intEL, strES

End Sub

Public Sub Save()

          Dim sql As String

11170     On Error GoTo Save_Error

          '20    sql = "Delete from Demographics50 where " & _
          '            "SampleID = '" & SampleID & "'"
          '30    Cnxn(0).Execute Sql
          '
          '40    For Each Fx In mcolDemographics
          '
          '50      sql = "INSERT INTO Demographics50 (SampleID, Request, UserName) " & _
          '              "VALUES " & _
          '              "('" & Fx.SampleID & "', " & _
          '              " '" & Fx.Request & "', " & _
          '              " '" & Fx.Username & "')"
          '60      Cnxn(0).Execute Sql
          '70    Next

11180     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

11190     intEL = Erl
11200     strES = Err.Description
11210     LogError "Demographics", "Save", intEL, strES, sql

End Sub






