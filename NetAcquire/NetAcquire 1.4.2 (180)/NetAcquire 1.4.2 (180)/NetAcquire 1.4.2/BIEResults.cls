VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBIEResults As New Collection
Private mBIEResults As New BIEResults

Public Function Add(BR As BIEResult) As BIEResult

          Dim Test As BIEResult
          Dim Found As Boolean

3590      On Error GoTo Add_Error

3600      Found = False
3610      For Each Test In mcolBIEResults
3620          If Test.Code = BR.Code Then
3630              Found = True
3640              Exit For
3650          End If
3660      Next

3670      If Not Found Then
3680          mcolBIEResults.Add BR, BR.Code
3690          Set Add = BR
3700      End If

3710      Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer

3720      intEL = Erl
3730      strES = Err.Description
3740      LogError "BIEResults", "Add", intEL, strES

End Function

Public Property Get BIEResults() As BIEResults

3750      On Error GoTo BIEResults_Error

3760      Set BIEResults = mBIEResults

3770      Exit Property

BIEResults_Error:

          Dim strES As String
          Dim intEL As Integer

3780      intEL = Erl
3790      strES = Err.Description
3800      LogError "BIEResults", "BIEResults", intEL, strES


End Property

Public Sub Clear()

          Dim n As Integer

3810      On Error GoTo Clear_Error

3820      For n = 1 To mcolBIEResults.Count
3830          mcolBIEResults.Remove 1
3840      Next

3850      Exit Sub

Clear_Error:

          Dim strES As String
          Dim intEL As Integer

3860      intEL = Erl
3870      strES = Err.Description
3880      LogError "BIEResults", "Clear", intEL, strES


End Sub

Public Function Count() As Long

3890      On Error GoTo Count_Error

3900      Count = mcolBIEResults.Count

3910      Exit Function

Count_Error:

          Dim strES As String
          Dim intEL As Integer

3920      intEL = Erl
3930      strES = Err.Description
3940      LogError "BIEResults", "Count", intEL, strES


End Function

Public Function Item(ByVal Code As String) As BIEResult
Attribute Item.VB_UserMemId = 0


          Dim i As Integer
3950      On Error GoTo Item_Error

3960      Set Item = Nothing
3970      For i = 1 To mcolBIEResults.Count
3980          If mcolBIEResults.Item(i).Code = Code Then
3990              Set Item = mcolBIEResults(i)
4000          End If
4010      Next i



4020      Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

4030      intEL = Erl
4040      strES = Err.Description
4050      LogError "BIEResults", "Item", intEL, strES


End Function


Public Function Load(ByVal Discipline As String, _
          ByVal SampleID As String, _
          ByVal ResultOrRepeat As String, _
          ByVal v As Integer, _
          ByVal p As Integer, _
          Optional ByVal Cat As String, _
          Optional ByVal Connection As Integer, _
          Optional ByVal Sex As String = "", _
          Optional ByVal DoB As String = "") _
          As BIEResults
          'Discipline is either "Bio", "Imm" or "End"

          Dim BRs As New BIEResults
          Dim BR As BIEResult
          Dim tb As Recordset
          '+++ Junaid 25-08-2023
          Dim tbA As Recordset
          Dim l_Analyser As String
          '--- Junaid
          Dim tbRange As Recordset
          Dim tbNewIDX As Recordset
          Dim sql As String
          Dim DaysOld As Long
          Dim SelectNormalRange As String
          Dim SelectFlagRange As String
          Dim TableName As String
          Dim SelectSexLow As String
          Dim SelectSexHigh As String

4060      On Error GoTo Load_Error

4070      If Cat = "" Then Cat = "Human"

4080      If Val(SampleID) = 0 Then
4090          Set BRs = Nothing
4100          Exit Function
4110      End If

4120      TableName = Discipline & ResultOrRepeat

4130      If UCase$(Discipline) = "BIO" Then
4140          sql = "UPDATE BioResults " & _
                  "SET SampleType = 'S' " & _
                  "WHERE SampleID = '" & SampleID & "' " & _
                  "AND COALESCE(SampleType, '') = ''"
4150          Cnxn(0).Execute sql
4160      End If

4170      If Sex = "" And DoB = "" Then
4180          sql = "SELECT DoB, Sex FROM Demographics WHERE " & _
                  "SampleID = '" & SampleID & "'"
4190          Set tb = Cnxn(Connection).Execute(sql)
4200          If Not tb.EOF Then
4210              DoB = tb!DoB & ""
4220              Sex = tb!Sex & ""
4230          End If
4240      End If

4250      If IsDate(DoB) Then
4260          DoB = Format$(DoB, "dd/mmm/yyyy")
4270          DaysOld = DateDiff("d", DoB, Now)
4280      End If

4290      Select Case Left$(UCase$(Trim$(Sex)), 1)
              Case "M": SelectSexLow = "Male": SelectSexHigh = "Male"
4300          Case "F": SelectSexLow = "Female": SelectSexHigh = "Female"
4310          Case Else: SelectSexLow = "Female": SelectSexHigh = "Male"
4320      End Select

4330      SelectNormalRange = " COALESCE(" & SelectSexLow & "Low, 0) Low, " & _
              " COALESCE(" & SelectSexHigh & "High, 9999) High, "
4340      SelectFlagRange = " COALESCE(Flag" & SelectSexLow & "Low, 0) FlagLow, " & _
              " COALESCE(Flag" & SelectSexHigh & "High, 9999) FlagHigh, "
          '+++Junaid 25-08-2023
4350      sql = "Select Analyser from " & TableName & " " & "Where SampleID = '" & SampleID & "' "
4360      Set tbA = New Recordset
4370      RecOpenServer Connection, tbA, sql
4380      If Not tbA Is Nothing Then
4390          If Not tbA.EOF Then
4400              l_Analyser = ConvertNull(tbA!Analyser, "")
4410          End If
4420      End If
          
4430      If UCase(Trim(l_Analyser)) <> "BIOMNIS" And UCase(Trim(l_Analyser)) <> "MATER" And UCase(Trim(l_Analyser)) <> "NVRL" And UCase(Trim(l_Analyser)) <> "BEAUMONT" Then
4440          sql = "UPDATE " & TableName & " " & _
                  "SET Result = 'XXXXX' " & _
                  "WHERE SampleID = '" & SampleID & "' " & _
                  "AND ( Result LIKE 'I_' " & _
                  "      OR Result LIKE 'L_' " & _
                  "      OR Result LIKE 'H_' )"
4450          Cnxn(0).Execute sql
4460      End If
          '--- Junaid
4470      sql = "SELECT R.DefIndex, COALESCE(X.NormalLow, 0) Low, " & _
              "COALESCE(X.NormalHigh, 9999) High, COALESCE(X.FlagLow, 0) FlagLow, COALESCE(X.FlagHigh, 9999) FlagHigh, " & _
              "COALESCE(X.PlausibleLow, 0) PlausibleLow, COALESCE(X.PlausibleHigh, 9999) PlausibleHigh, " & _
              "LongName, ShortName, DoDelta, DeltaLimit, " & _
              "DP, PrintPriority, " & _
              "R.SampleID, R.Code, R.Result, " & _
              "COALESCE(R.Valid, 0) AS Valid, COALESCE(R.Printed, 0) Printed, " & _
              "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
              "R.SampleType, R.Analyser, R.Faxed, R.Authorised, " & _
              "R.Comment AS Comment, " & _
              "COALESCE(PrintRefRange, 1) PrntRefRange , D.DeltaDaysBackLimit " & _
              "FROM " & TableName & " R JOIN " & Discipline & "TestDefinitions D ON R.Code = D.Code " & _
              "LEFT JOIN " & Discipline & "DefIndex X ON R.DefIndex = X.DefIndex " & _
              "WHERE " & _
              "SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
              "AND R.Code = D.Code AND D.Inuse = 1 " & _
              "AND AgeFromDays <= " & DaysOld & " " & _
              "AND AgeToDays >= " & DaysOld & " "

          '290   If UCase(HospName(0)) <> "CAVAN" Then
4480      sql = sql & "AND R.SampleType = D.SampleType "
          '310   End If
4490      If p = gNOTPRINTED And v = gNOTVALID Then
4500          sql = sql & "AND Printed = 0 AND Valid = 0 "
4510      ElseIf p = gNOTPRINTED And v = gVALID Then
4520          sql = sql & "AND Printed = 0 AND Valid = 1 "
4530      ElseIf p = gNOTPRINTED And v = gDONTCARE Then
4540          sql = sql & "AND Printed = 0 "
4550      ElseIf p = gPRINTED And v = gNOTVALID Then
4560          sql = sql & "AND Printed = 1 AND Valid = 0 "
4570      ElseIf p = gPRINTED And v = gVALID Then
4580          sql = sql & "AND Printed = 1 AND Valid = 1 "
4590      ElseIf p = gPRINTED And v = gDONTCARE Then
4600          sql = sql & "AND Printed = 1 "
4610      ElseIf p = gDONTCARE And v = gNOTVALID Then
4620          sql = sql & "AND Valid = 0 "
4630      ElseIf p = gDONTCARE And v = gVALID Then
4640          sql = sql & "AND Valid = 1 "
4650      End If

4660      sql = sql & "ORDER BY PrintPriority ASC"
4670      Set tb = New Recordset
          'MsgBox Sql
4680      RecOpenServer Connection, tb, sql    '  RecOpenClient 0,tb, Sql
          '+++ Junaid 25-085-2023
4690      If Not tb Is Nothing Then
4700          If Not tb.EOF Then
4710              If UCase(Trim(l_Analyser)) = "BIOMNIS" Or UCase(Trim(l_Analyser)) = "MATER" Or UCase(Trim(l_Analyser)) = "NVRL" Or UCase(Trim(l_Analyser)) = "BEAUMONT" Then
4720                  Exit Function
4730              End If
4740          End If
4750      End If
          '--- Junaid
4760      Do While Not tb.EOF
4770          Set BR = New BIEResult
4780          With BR
4790              .SampleID = tb!SampleID & ""
4800              .Code = tb!Code & ""
4810              .Result = tb!Result & ""
4820              .Operator = tb!Operator & ""
4830              .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
4840              .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm:ss")
4850              .Units = tb!Units & ""
4860              .Printed = tb!Printed
4870              .Valid = IIf(tb!Valid, True, False)
4880              If Not IsNull(tb!Authorised) Then
4890                  .Authorised = IIf(tb!Authorised, True, False)
4900              Else
4910                  .Authorised = False
4920              End If
                  '    .Faxed = tb!Faxed
                  '    If TB!h & "" <> "" Then .h = TB!h Else .h = False
                  '    If TB!g & "" <> "" Then .g = TB!g Else .g = False
                  '    If TB!o & "" <> "" Then .o = TB!o Else .o = False
                  '    If TB!J & "" <> "" Then .J = TB!h Else .J = False
                  '    If TB!l & "" <> "" Then .l = TB!h Else .l = False
                  '    If TB!s & "" <> "" Then .s = TB!h Else .s = False
4930              .Flags = tb!Flags & ""
4940              .SampleType = tb!SampleType & ""
4950              .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
4960              .ShortName = tb!ShortName & ""
4970              .LongName = tb!LongName & ""
4980              If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
4990              .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
5000              .Analyser = tb!Analyser & ""
5010              .Comment = tb!Comment & ""
5020              .Low = tb!Low
5030              .High = tb!High
5040              .FlagLow = tb!FlagLow
5050              .FlagHigh = tb!FlagHigh
5060              .PlausibleLow = tb!PlausibleLow
5070              .PlausibleHigh = tb!PlausibleHigh
5080              .PrintRefRange = tb!PrntRefRange
5090              .DefIndex = tb!DefIndex
5100              .DeltaDaysBackLimit = IIf(IsNull(tb!DeltaDaysBackLimit), 0, tb!DeltaDaysBackLimit)
5110              If .DefIndex = 0 Then
5120                  If DoB <> "" And Sex <> "" Then
5130                      sql = "SELECT " & _
                              SelectNormalRange & SelectFlagRange & _
                              "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                              "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                              "FROM " & Discipline & "TestDefinitions  " & _
                              "WHERE category = '" & Cat & "' " & _
                              "AND Code = '" & .Code & "' " & _
                              "AND AgeFromDays <= " & DaysOld & " " & _
                              "AND AgeToDays >= " & DaysOld & " "
5140                      Set tbRange = New Recordset
5150                      RecOpenServer 0, tbRange, sql
5160                      If Not tbRange.EOF Then
5170                          .Low = tbRange!Low
5180                          .High = tbRange!High
5190                          .FlagLow = tbRange!FlagLow
5200                          .FlagHigh = tbRange!FlagHigh
5210                          .PlausibleLow = tbRange!PlausibleLow
5220                          .PlausibleHigh = tbRange!PlausibleHigh

5230                          sql = "SELECT * FROM " & Discipline & "DefIndex " & _
                                  "WHERE NormalLow = '" & .Low & "' " & _
                                  "AND NormalHigh = '" & .High & "' " & _
                                  "AND FlagLow = '" & .FlagLow & "' " & _
                                  "AND FlagHigh = '" & .FlagHigh & "' " & _
                                  "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                                  "AND PlausibleHigh = '" & .PlausibleHigh & "' "

5240                          Set tbNewIDX = New Recordset
5250                          RecOpenClient 0, tbNewIDX, sql
5260                          If Not tbNewIDX.EOF Then
5270                              .DefIndex = tbNewIDX!DefIndex
5280                          Else

5290                              sql = "INSERT INTO " & Discipline & "DefIndex " & _
                                      "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                      "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                      "VALUES ( " & _
                                      .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                                      .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
5300                              Cnxn(0).Execute sql

5310                              sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
5320                              Set tbNewIDX = New Recordset
5330                              RecOpenClient 0, tbNewIDX, sql
5340                              .DefIndex = tbNewIDX!NewIndex

5350                          End If

5360                          sql = "UPDATE " & TableName & " " & _
                                  "SET DefIndex = '" & .DefIndex & "' " & _
                                  "WHERE SampleID = '" & .SampleID & "' " & _
                                  "AND Code = '" & .Code & "'"
5370                          Cnxn(0).Execute sql

5380                      End If
5390                  End If
5400              End If


5410              BRs.Add BR
5420          End With
5430          tb.MoveNext
5440      Loop

5450      If BRs.Count <> 0 Then
5460          Set Load = BRs
5470      Else
5480          Set Load = Nothing
5490      End If
5500      Set BR = Nothing
5510      Set BRs = Nothing

5520      Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

5530      intEL = Erl
5540      strES = Err.Description
5550      LogError "BIEResults", "Load", intEL, strES, sql

End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

5560      On Error GoTo NewEnum_Error

5570      Set NewEnum = mcolBIEResults.[_NewEnum]

5580      Exit Function

NewEnum_Error:

          Dim strES As String
          Dim intEL As Integer

5590      intEL = Erl
5600      strES = Err.Description
5610      LogError "BIEResults", "NewEnum", intEL, strES


End Function



Public Sub RemoveItem(ByVal X As Long)

5620      On Error GoTo RemoveItem_Error

5630      mcolBIEResults.Remove X

5640      Exit Sub

RemoveItem_Error:

          Dim strES As String
          Dim intEL As Integer

5650      intEL = Erl
5660      strES = Err.Description
5670      LogError "BIEResults", "RemoveItem", intEL, strES


End Sub


Public Sub Save(ByVal Discipline As String, _
          ByVal BRs As BIEResults)
          'Discipline is either "Bio", "Imm" or "End"

          Dim tb As Recordset
          Dim sql As String
          Dim BR As BIEResult
          Dim lngSampleID As Long

5680      On Error GoTo Save_Error

5690      For Each BR In BRs
5700          With BR
5710              lngSampleID = Format$(Val(.SampleID))
5720              sql = "Select * from " & Discipline & "Results where " & _
                      "SampleID = '" & lngSampleID & "' " & _
                      "and Code = '" & BR.Code & "'"
5730              Set tb = New Recordset
5740              RecOpenClient 0, tb, sql
5750              If Not tb.EOF Then
5760                  sql = "Select * from " & Discipline & "Repeats where " & _
                          "SampleID = '" & lngSampleID & "'"
5770                  Set tb = New Recordset
5780                  RecOpenClient 0, tb, sql
5790              End If
5800              tb.AddNew
5810              tb!SampleID = lngSampleID
5820              tb!Rundate = Format$(.Rundate, "dd/mmm/yyyy")
5830              If Trim$(.RunTime) <> "" Then
5840                  tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
5850              End If
5860              tb!Code = .Code
5870              tb!Result = .Result
5880              tb!Units = .Units
5890              tb!Printed = .Printed
5900              tb!Valid = .Valid
5910              tb!FAXed = 0
5920              tb!Analyser = .Analyser
5930              tb!SampleType = .SampleType
5940              tb.Update

5950              sql = "Delete from " & Discipline & "Requests where " & _
                      "SampleID = '" & lngSampleID & "' " & _
                      "and Code = '" & BR.Code & "'"
5960              Cnxn(0).Execute sql

5970          End With
5980      Next

5990      Set BR = Nothing

6000      Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

6010      intEL = Erl
6020      strES = Err.Description
6030      LogError "BIEResults", "Save", intEL, strES, sql

End Sub





