VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HaemResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolHaemResults As New Collection
Private mHaemResults As New HaemResults

Public Function Add(BR As BIEResult) As BIEResult

17700     mcolHaemResults.Add BR

17710     Set Add = BR

End Function

Public Property Get HaemResults() As HaemResults
        
17720     Set HaemResults = mHaemResults

End Property


Public Sub Clear()

          Dim n As Integer

17730     For n = 1 To mcolHaemResults.Count
17740         mcolHaemResults.Remove 1
17750     Next

End Sub


Public Function Count() As Long
        
17760     Count = mcolHaemResults.Count

End Function


Public Function Item(ByVal X As Long) _
          As BIEResult
        
17770     On Error Resume Next

17780     Set Item = mcolHaemResults(X)

End Function




Public Function Load(ByVal SampleID As String) _
          As HaemResults

          Dim BRs As New HaemResults
          Dim BR As BIEResult
          Dim tb As Recordset
          Dim tbRange As Recordset
          Dim tbNewIDX As Recordset
          Dim sql As String
          Dim DaysOld As Long
          Dim SelectNormalRange As String
          Dim SelectFlagRange As String
          Dim SelectSexLow As String
          Dim SelectSexHigh As String
          Dim Sex As String
          Dim DoB As String

17790     If Val(SampleID) = 0 Then
17800         Set BRs = Nothing
17810         Exit Function
17820     End If

17830     sql = "SELECT DoB, Sex FROM Demographics WHERE " & _
              "SampleID = '" & SampleID & "'"
17840     Set tb = Cnxn(0).Execute(sql)
17850     If Not tb.EOF Then
17860         DoB = tb!DoB & ""
17870         Sex = tb!Sex & ""
17880     End If

17890     If IsDate(DoB) Then
17900         DoB = Format$(DoB, "dd/mmm/yyyy")
17910         DaysOld = DateDiff("d", DoB, Now)
17920     End If
          
17930     Select Case Left$(UCase$(Trim$(Sex)), 1)
              Case "M": SelectSexLow = "Male": SelectSexHigh = "Male"
17940         Case "F": SelectSexLow = "Female": SelectSexHigh = "Female"
17950         Case Else: SelectSexLow = "Female": SelectSexHigh = "Male"
17960     End Select
        
17970     SelectNormalRange = " COALESCE(" & SelectSexLow & "Low, 0) Low, " & _
              " COALESCE(" & SelectSexHigh & "High, 9999) High, "
17980     SelectFlagRange = " COALESCE(Flag" & SelectSexLow & "Low, 0) FlagLow, " & _
              " COALESCE(Flag" & SelectSexHigh & "High, 9999) FlagHigh, "

17990     sql = "SELECT R.DefIndex, X.NormalLow Low, X.NormalHigh High, X.FlagLow, X.FlagHigh, X.PlausibleLow, X.PlausibleHigh, " & _
              "D.LongName, D.ShortName, D.DoDelta, D.DeltaValue, " & _
              "PrintPriority, COALESCE(PrintFormat, 1) PrintFormat, " & _
              "R.SampleID, R.Code, R.Result, " & _
              "D.ResultGroup, " & _
              "R.RunDateTime, R.UserName, R.Flags, R.Units, " & _
              "R.SampleType, R.Analyser, " & _
              "R.Comment AS Comment " & _
              " " & _
              "FROM HaemResults50 R, HaemTestDefinitions D, " & _
              "HaemDefIndex X " & _
              "WHERE " & _
              "R.DefIndex = X.Defindex " & _
              "AND SampleID = '" & SampleID & "' " & _
              "AND R.Code = D.Code " & _
              "AND AgeFromDays <= " & DaysOld & " " & _
              "AND AgeToDays >= " & DaysOld & " " & _
              "ORDER BY PrintPriority ASC"
18000     Set tb = New Recordset
18010     RecOpenServer 0, tb, sql
18020     Do While Not tb.EOF
18030         Set BR = New BIEResult
18040         With BR
18050             .SampleID = tb!SampleID & ""
18060             .Code = tb!Code & ""
18070             .Result = tb!Result & ""
18080             .Operator = tb!UserName & ""
18090             .Rundate = Format$(tb!RunDateTime, "dd/mm/yyyy")
18100             .RunTime = Format$(tb!RunDateTime, "dd/mm/yyyy hh:mm:ss")
18110             .Units = tb!Units & ""
18120             .ResultGroup = tb!ResultGroup & ""
                  '630       .Valid = IIf(tb!Valid, True, False)
                  '640       If Not IsNull(tb!Authorised) Then
                  '650         .Authorised = IIf(tb!Authorised, True, False)
                  '660       Else
                  '670         .Authorised = False
                  '680       End If
                  '    .Faxed = tb!Faxed
                  '    If TB!h & "" <> "" Then .h = TB!h Else .h = False
                  '    If TB!g & "" <> "" Then .g = TB!g Else .g = False
                  '    If TB!o & "" <> "" Then .o = TB!o Else .o = False
                  '    If TB!J & "" <> "" Then .J = TB!h Else .J = False
                  '    If TB!l & "" <> "" Then .l = TB!h Else .l = False
                  '    If TB!s & "" <> "" Then .s = TB!h Else .s = False
18130             .Flags = tb!Flags & ""
18140             .SampleType = tb!SampleType & ""
18150             .Printformat = tb!Printformat
18160             .ShortName = tb!ShortName & ""
18170             .LongName = tb!LongName & ""
18180             If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
18190             .DeltaLimit = IIf(IsNull(tb!DeltaValue), 9999, tb!DeltaValue)
18200             .Analyser = tb!Analyser & ""
18210             .Comment = tb!Comment & ""
18220             .Low = tb!Low
18230             .High = tb!High
18240             .FlagLow = tb!FlagLow
18250             .FlagHigh = tb!FlagHigh
18260             .PlausibleLow = tb!PlausibleLow
18270             .PlausibleHigh = tb!PlausibleHigh
                  '840       .PrintRefRange = tb!PrntRefRange
18280             .DefIndex = tb!DefIndex
          
18290             If .DefIndex = 0 Then
18300                 If DoB <> "" And Sex <> "" Then
18310                     sql = "SELECT " & _
                              SelectNormalRange & " COALESCE(FlagLow, 0) FlagLow, COALESCE(FlagHigh, 9999) FlagHigh, " & _
                              "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                              "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                              "FROM HaemTestDefinitions  " & _
                              "WHERE Code = '" & .Code & "' " & _
                              "AND AgeFromDays <= " & DaysOld & " " & _
                              "AND AgeToDays >= " & DaysOld & " "
18320                     Set tbRange = New Recordset
18330                     RecOpenServer 0, tbRange, sql
18340                     If Not tbRange.EOF Then
18350                         .Low = tbRange!Low
18360                         .High = tbRange!High
18370                         .FlagLow = tbRange!FlagLow
18380                         .FlagHigh = tbRange!FlagHigh
18390                         .PlausibleLow = tbRange!PlausibleLow
18400                         .PlausibleHigh = tbRange!PlausibleHigh
          
18410                         sql = "SELECT * FROM HaemDefIndex " & _
                                  "WHERE NormalLow = '" & .Low & "' " & _
                                  "AND NormalHigh = '" & .High & "' " & _
                                  "AND FlagLow = '" & .FlagLow & "' " & _
                                  "AND FlagHigh = '" & .FlagHigh & "' " & _
                                  "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                                  "AND PlausibleHigh = '" & .PlausibleHigh & "' "
          
18420                         Set tbNewIDX = New Recordset
18430                         RecOpenClient 0, tbNewIDX, sql
18440                         If Not tbNewIDX.EOF Then
18450                             .DefIndex = tbNewIDX!DefIndex
18460                         Else
          
18470                             sql = "INSERT INTO HaemDefIndex " & _
                                      "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                      "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                      "VALUES ( " & _
                                      .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                                      .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
18480                             Cnxn(0).Execute sql
          
18490                             sql = "SELECT MAX(DefIndex) NewIndex FROM HaemDefIndex"
18500                             Set tbNewIDX = New Recordset
18510                             RecOpenClient 0, tbNewIDX, sql
18520                             .DefIndex = tbNewIDX!NewIndex
          
18530                         End If
          
18540                         sql = "UPDATE HaemResults50 " & _
                                  "SET DefIndex = '" & .DefIndex & "' " & _
                                  "WHERE SampleID = '" & .SampleID & "' " & _
                                  "AND Code = '" & .Code & "'"
18550                         Cnxn(0).Execute sql
          
18560                     End If
18570                 End If
18580             End If
          
        
18590             BRs.Add BR
18600         End With
18610         tb.MoveNext
18620     Loop

18630     If BRs.Count <> 0 Then
18640         Set Load = BRs
18650     Else
18660         Set Load = Nothing
18670     End If
18680     Set BR = Nothing
18690     Set BRs = Nothing

End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
18700     Set NewEnum = mcolHaemResults.[_NewEnum]

End Function



Public Sub RemoveItem(ByVal X As Long)

18710     mcolHaemResults.Remove X

End Sub



Public Sub Save(ByVal Discipline As String, _
          ByVal BRs As BIEResults)
          'Discipline is either "Bio", "Imm" or "End"

          Dim tb As Recordset
          Dim sql As String
          Dim BR As BIEResult
          Dim lngSampleID As Long

18720     On Error GoTo Save_Error

18730     For Each BR In BRs
18740         With BR
18750             lngSampleID = Format$(Val(.SampleID))
18760             sql = "Select * from " & Discipline & "Results where " & _
                      "SampleID = '" & lngSampleID & "' " & _
                      "and Code = '" & BR.Code & "'"
18770             Set tb = New Recordset
18780             RecOpenClient 0, tb, sql
18790             If Not tb.EOF Then
18800                 sql = "Select * from " & Discipline & "Repeats where " & _
                          "SampleID = '" & lngSampleID & "'"
18810                 Set tb = New Recordset
18820                 RecOpenClient 0, tb, sql
18830             End If
18840             tb.AddNew
18850             tb!SampleID = lngSampleID
18860             tb!Rundate = Format$(.Rundate, "dd/mmm/yyyy")
18870             If Trim$(.RunTime) <> "" Then
18880                 tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
18890             End If
18900             tb!Code = .Code
18910             tb!Result = .Result
18920             tb!Units = .Units
18930             tb!Printed = .Printed
18940             tb!Valid = .Valid
18950             tb!FAXed = 0
18960             tb!Analyser = .Analyser
18970             tb!SampleType = .SampleType
18980             tb.Update

18990             sql = "Delete from " & Discipline & "Requests where " & _
                      "SampleID = '" & lngSampleID & "' " & _
                      "and Code = '" & BR.Code & "'"
19000             Cnxn(0).Execute sql
          
19010         End With
19020     Next

19030     Set BR = Nothing

19040     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

19050     intEL = Erl
19060     strES = Err.Description
19070     LogError "BIEResults", "Save", intEL, strES, sql

End Sub







