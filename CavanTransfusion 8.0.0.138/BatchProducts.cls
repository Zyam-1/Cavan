VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBatchProducts As New Collection

Private mBatchProducts As New BatchProducts

Public Function CountProductBatchInStock(ByVal Product As String, ByVal BatchNumber As String) As Integer

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CountProductBatchInStock_Error

20    sql = "SELECT COUNT(*) Tot FROM BatchProducts " & _
            "WHERE Product = '" & Product & "' " & _
            "AND BatchNumber = '" & BatchNumber & "' " & _
            "AND (EventCode = 'C' OR EventCode = 'R')"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    CountProductBatchInStock = tb!Tot

60    Exit Function

CountProductBatchInStock_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BatchProducts", "CountProductBatchInStock", intEL, strES, sql

End Function

Public Function CountProductBatchInStockNotExpired(ByVal Product As String, ByVal BatchNumber As String) As Integer

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CountProductBatchInStockNotExpired_Error

20    sql = "SELECT COUNT(*) Tot FROM BatchProducts " & _
            "WHERE Product = '" & Product & "' " & _
            "AND BatchNumber = '" & BatchNumber & "' " & _
            "AND DATEDIFF(day, getdate(), DateExpiry) >= 0 " & _
            "AND (EventCode = 'C' OR EventCode = 'R')"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    CountProductBatchInStockNotExpired = tb!Tot

60    Exit Function

CountProductBatchInStockNotExpired_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BatchProducts", "CountProductBatchInStockNotExpired", intEL, strES, sql

End Function


Private Function Exists(ByVal BatchNumber As String, ByVal Product As String) As Boolean

      Dim BP As BatchProduct

10    On Error GoTo Exists_Error

20    Exists = False

30    For Each BP In mcolBatchProducts
40      If BP.BatchNumber = BatchNumber And BP.Product = Product Then
50        Exists = True
60        Exit For
70      End If
80    Next

90    Exit Function

Exists_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "Exists", intEL, strES

End Function

Public Sub LoadStockNotExpired()

Dim sql As String
Dim tb As Recordset

10    On Error GoTo LoadStockNotExpired_Error

20      sql = "SELECT * FROM BatchProducts " & _
              "WHERE (EventCode = 'C' OR EventCode = 'R') " & _
              "AND DATEDIFF(day, getdate(), DateExpiry) >= 0 " & _
              "Order by Product, DateReceived desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadStockNotExpired_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadStockNotExpired", intEL, strES, sql

End Sub

Public Sub LoadStockAllNotExpired()

Dim sql As String
Dim tb As Recordset

10    On Error GoTo LoadStockAllNotExpired_Error

20      sql = "SELECT * FROM BatchProducts " & _
              "WHERE DATEDIFF(day, getdate(), DateExpiry) >= 0 " & _
              "Order by Product, DateReceived desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadStockAllNotExpired_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadStockAllNotExpired", intEL, strES, sql

End Sub
Public Sub LoadStockAll()

Dim sql As String
Dim tb As Recordset

10    On Error GoTo LoadStockAll_Error

20      sql = "SELECT * FROM BatchProducts " & _
              "Order by Product, DateReceived desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadStockAll_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadStockAll", intEL, strES, sql

End Sub

Public Sub LoadStockWithExpired()

Dim sql As String
Dim tb As Recordset

10    On Error GoTo LoadStockWithExpired_Error

20      sql = "SELECT * FROM BatchProducts " & _
              "WHERE (EventCode = 'C' OR EventCode = 'R' )" & _
              "Order by Product, DateReceived desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadStockWithExpired_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadStockWithExpired", intEL, strES, sql

End Sub

Private Sub Populate(ByVal tb As Recordset)

      Dim BP As New BatchProduct
10    On Error GoTo Populate_Error

20    With BP
30      .BatchNumber = tb!BatchNumber & ""
40      .Product = tb!Product & ""
50      .Identifier = tb!Identifier & ""
60      .UnitVolume = tb!UnitVolume & ""
70      .DateExpiry = tb!DateExpiry & ""
80      .DateReceived = tb!DateReceived & ""
90      .UnitGroup = tb!UnitGroup & ""
100     .Concentration = tb!Concentration & ""
  
110     .Chart = tb!Chart & ""
120     .PatName = tb!PatName & ""
130     .DoB = tb!DoB & ""
140     .Age = tb!Age & ""
150     .Sex = tb!Sex & ""
160     .Addr0 = tb!Addr0 & ""
170     .Addr1 = tb!Addr1 & ""
180     .Addr2 = tb!Addr2 & ""
190     .Ward = tb!Ward & ""
200     .Clinician = tb!Clinician & ""
  
210     .PatientGroup = tb!PatientGroup & ""
220     .SampleID = tb!SampleID & ""
230     .Typenex = tb!Typenex & ""
240     .AandE = tb!AandE & ""
250     .EventCode = tb!EventCode & ""
260     .Comment = tb!Comment & ""
270     .EventStart = tb!EventStart & ""
280     .EventEnd = tb!EventEnd & ""
290     .UserName = tb!UserName & ""
300     .RecordDateTime = tb!RecordDateTime & ""
310     .LabelPrinted = tb!LabelPrinted

320     mcolBatchProducts.Add BP
330   End With

340   Exit Sub

Populate_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "BatchProducts", "Populate", intEL, strES

End Sub

Public Sub Update(ByVal BP As BatchProduct)
  
      Dim sql As String

10    On Error GoTo Update_Error

20    sql = "IF EXISTS (SELECT * FROM BatchProducts WHERE " & _
            "           Identifier = '" & BP.Identifier & "') " & _
            "  UPDATE BatchProducts " & _
            "  SET BatchNumber = '" & BP.BatchNumber & "', " & _
            "  Product = '" & BP.Product & "', " & _
            "  Identifier = '" & BP.Identifier & "', " & _
            "  UnitVolume = '" & BP.UnitVolume & "', " & _
            "  DateExpiry = '" & Format$(BP.DateExpiry, "dd/MMM/yyyy") & "', " & _
            "  DateReceived = '" & Format$(BP.DateReceived, "dd/MMM/yyyy") & "', " & _
            "  UnitGroup = '" & BP.UnitGroup & "', " & _
            "  Concentration = '" & BP.Concentration & "', " & _
            "  Chart = '" & BP.Chart & "', " & _
            "  PatName = '" & AddTicks(BP.PatName) & "', " & _
            "  DoB = ' " & Format$(BP.DoB, "dd/MMM/yyyy") & "', " & _
            "  Age = '" & BP.Age & "', " & _
            "  Sex = '" & BP.Sex & "', " & _
            "  Addr0 = '" & AddTicks(BP.Addr0) & "', " & _
            "  Addr1 = '" & AddTicks(BP.Addr1) & "', " & _
            "  Addr2 = '" & AddTicks(BP.Addr2) & "', " & _
            "  Ward = '" & AddTicks(BP.Ward) & "', " & _
            "  Clinician = '" & AddTicks(BP.Clinician) & "', " & _
            "  PatientGroup = '" & BP.PatientGroup & "', " & _
            "  SampleID = '" & BP.SampleID & "', " & _
            "  Typenex = '" & BP.Typenex & "', " & _
      "  AandE = '" & BP.AandE & "', "
30    sql = sql & "  EventCode = '" & BP.EventCode & "', " & _
            "  Comment = '" & AddTicks(BP.Comment) & "', " & _
            "  EventStart = '" & BP.EventStart & "', " & _
            "  EventEnd = '" & BP.EventEnd & "', " & _
            "  UserName = '" & AddTicks(BP.UserName) & "', " & _
            "  RecordDateTime = getdate(), LabelPrinted = '" & BP.LabelPrinted & "' " & _
            "  WHERE Identifier = '" & BP.Identifier & "' " & _
            "ELSE " & _
            "  INSERT INTO BatchProducts " & _
            "( BatchNumber, Product, Identifier, UnitVolume, DateExpiry, DateReceived, UnitGroup, Concentration, " & _
            "  Chart, PatName, DoB, Age, Sex, Addr0, Addr1, Addr2, Ward, Clinician, PatientGroup, SampleID, Typenex, " & _
            "  AandE, EventCode, Comment, EventStart, EventEnd, UserName, RecordDateTime, LabelPrinted) " & _
            "  VALUES " & _
            "( '" & BP.BatchNumber & "', " & _
            "  '" & BP.Product & "', " & _
            "  '" & BP.Identifier & "', " & _
            "  '" & BP.UnitVolume & "', " & _
            "  '" & Format$(BP.DateExpiry, "dd/MMM/yyyy") & "', " & _
            "  '" & Format$(BP.DateReceived, "dd/MMM/yyyy") & "', " & _
            "  '" & BP.UnitGroup & "', " & _
            "  '" & BP.Concentration & "', " & _
            "  '" & BP.Chart & "', " & _
            "  '" & AddTicks(BP.PatName) & "', " & _
            "  '" & Format$(BP.DoB, "dd/MMM/yyyy") & "', "
40    sql = sql & "  '" & BP.Age & "', " & _
            "  '" & BP.Sex & "', " & _
            "  '" & AddTicks(BP.Addr0) & "', " & _
            "  '" & AddTicks(BP.Addr1) & "', " & _
            "  '" & AddTicks(BP.Addr2) & "', " & _
            "  '" & AddTicks(BP.Ward) & "', " & _
            "  '" & AddTicks(BP.Clinician) & "', " & _
            "  '" & BP.PatientGroup & "', " & _
            "  '" & BP.SampleID & "', " & _
            "  '" & BP.Typenex & "', " & _
            "  '" & BP.AandE & "', " & _
            "  '" & BP.EventCode & "', " & _
            "  '" & AddTicks(BP.Comment) & "', " & _
            "  '" & BP.EventStart & "', " & _
            "  '" & BP.EventEnd & "', " & _
            "  '" & AddTicks(BP.UserName) & "', " & _
            "  getdate(), '" & BP.LabelPrinted & "' )"

50    CnxnBB(0).Execute sql

60    Exit Sub

Update_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BatchProducts", "Update", intEL, strES, sql

End Sub


Public Function Count() As Long
  
10    On Error Resume Next
20    Count = mcolBatchProducts.Count

End Function



Public Sub LoadSpecificIdentifier(ByVal Identifier As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSpecificIdentifier_Error

20    sql = "SELECT * FROM BatchProducts " & _
            "WHERE Identifier = '" & Identifier & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    If Not tb.EOF Then
60     Populate tb
70    End If

80    sql = "SELECT * FROM BatchProductsAudit " & _
            "WHERE Identifier = '" & Identifier & "' " & _
            "ORDER BY RecordDateTime desc"
90    Set tb = New Recordset

100   RecOpenClientBB 0, tb, sql
110   Do While Not tb.EOF
120     Populate tb
130     tb.MoveNext
140   Loop

150   Exit Sub

LoadSpecificIdentifier_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "BatchProducts", "LoadSpecificIdentifier", intEL, strES, sql

End Sub


Public Sub LoadSampleID(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSampleID_Error

20    sql = "SELECT * FROM BatchProducts " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60     Populate tb
70     tb.MoveNext
80    Loop

90    sql = "SELECT * FROM BatchProductsAudit " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "ORDER BY RecordDateTime desc"
100   Set tb = New Recordset

110   RecOpenClientBB 0, tb, sql
120   Do While Not tb.EOF
130     Populate tb
140     tb.MoveNext
150   Loop

160   Exit Sub

LoadSampleID_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "BatchProducts", "LoadSampleID", intEL, strES, sql

End Sub

Public Sub LoadBatchNumbersFromProduct(ByVal Product As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadBatchNumbersFromProduct_Error

20    sql = "SELECT DISTINCT BatchNumber, Product, " & _
            "'' Identifier, '' UnitVolume, DateExpiry, " & _
            "'01/01/1900' DateReceived, '' UnitGroup, " & _
            "'' Concentration, '' Chart, '' PatName, " & _
            "'01/01/1900' DoB, '' Age, '' Sex, " & _
            "'' Addr0, '' Addr1, '' Addr2, " & _
            "'' Ward, '' Clinician, " & _
            "'' PatientGroup, '' SampleID, '' Typenex, '' AandE, " & _
            "'' EventCode, '' Comment, " & _
            "'01/01/1900' EventStart, '01/01/1900' EventEnd, " & _
            "'' UserName, '01/01/1900' RecordDateTime, " & _
            "LabelPrinted " & _
            "FROM BatchProducts " & _
            "WHERE Product = '" & Product & "' " & _
            "AND (EventCode = 'C' OR EventCode = 'R') " & _
            "AND DateExpiry > '" & Format$(Now, "dd/mmm/yyyy") & "' "

30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60     Populate tb
70     tb.MoveNext
80    Loop

90    Exit Sub

LoadBatchNumbersFromProduct_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadBatchNumbersFromProduct", intEL, strES, sql

End Sub

Public Sub LoadSampleIDNoAudit(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSampleIDNoAudit_Error

20    sql = "SELECT * FROM BatchProducts " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60     Populate tb
70     tb.MoveNext
80    Loop

90    Exit Sub

LoadSampleIDNoAudit_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadSampleIDNoAudit", intEL, strES, sql

End Sub

Public Sub LoadSpecificBatch(ByVal BatchNumber As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSpecificBatch_Error

20    sql = "SELECT * FROM BatchProducts " & _
            "WHERE BatchNumber = '" & BatchNumber & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadSpecificBatch_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BatchProducts", "LoadSpecificBatch", intEL, strES, sql

End Sub
Public Sub LoadBetweenDates(ByVal FromDate As String, ByVal ToDate As String, ByVal EventCode As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadBetweenDates_Error

20    FromDate = Format$(FromDate, "dd/MMM/yyyy")
30    ToDate = Format$(ToDate, "dd/MMM/yyyy") & " 23:59"

40    sql = "SELECT * FROM BatchProducts WHERE " & _
            "(RecordDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "') "
50    If EventCode <> "" Then
60      sql = sql & "AND EventCode = '" & EventCode & "'"
70    End If

80    Set tb = New Recordset
90    RecOpenClientBB 0, tb, sql
100   Do While Not tb.EOF
110     Populate tb
120     tb.MoveNext
130   Loop

140   Exit Sub

LoadBetweenDates_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "BatchProducts", "LoadBetweenDates", intEL, strES, sql

End Sub
Public Function Item(ByVal Index As Integer) As BatchProduct
  
10    On Error Resume Next

20    Set Item = mcolBatchProducts.Item(Index)

End Function

Public Sub LoadSpecificIdentifierLatest(ByVal Identifier As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSpecificIdentifierLatest_Error

20    sql = "SELECT * FROM BatchProducts " & _
            "WHERE Identifier = '" & Identifier & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    If Not tb.EOF Then
60      Populate tb
70    End If

80    Exit Sub

LoadSpecificIdentifierLatest_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "BatchProducts", "LoadSpecificIdentifierLatest", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolBatchProducts.[_NewEnum]

End Function

Public Property Get BatchProducts() As BatchProducts
  
10    Set BatchProducts = mBatchProducts

End Property


