VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolProducts As New Collection

Private mProducts As New Products



Public Property Get Products() As Products
  
10    Set Products = mProducts

End Property
'Zyam 27-2-24
Public Function Count() As Long

10    On Error Resume Next
20    Count = mcolProducts.Count

End Function
'Zyam 27-2-24


Public Function Item(ByVal Index As Integer) As Product
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolProducts.Item(Index)

End Function
Public Sub LoadSampleID(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSampleID_Error

20    sql = "SELECT * FROM Product " & _
            "WHERE LabNumber = '" & SampleID & "' " & _
            "ORDER BY DateTime desc"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
        
        
        
80    Loop

90    Exit Sub

LoadSampleID_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "LoadSampleID", intEL, strES, sql

End Sub
Public Sub LoadLatestBySampleID(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadLatestBySampleID_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE LabNumber = '" & SampleID & "' " & _
            "ORDER BY DateTime desc"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
      
70      tb.MoveNext
        
        
        
80    Loop

90    Exit Sub

LoadLatestBySampleID_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "LoadLatestBySampleID", intEL, strES, sql

End Sub
Public Sub Load(ByVal PackNumber As String, _
                ByVal Barcode As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Load_Error

20    sql = "SELECT * FROM Product " & _
            "WHERE ISBT128 = '" & PackNumber & "' " & _
            "AND BarCode = '" & Barcode & "' "
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

90    Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "Load", intEL, strES, sql

End Sub

Public Sub LoadLatest(ByVal PackNumber As String, _
                      ByVal Barcode As String, _
                      ByVal DateExpiry As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Load_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE Number = '" & PackNumber & "' " & _
            "AND BarCode = '" & Barcode & "' " & _
            "AND DateExpiry = '" & Format$(DateExpiry, "dd/MMM/yyyy HH:mm") & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

90    Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "Load", intEL, strES, sql

End Sub


Public Sub LoadLatestISBT128(ByVal PackNumber As String, _
                      ByVal Barcode As String)

      Dim tb As Recordset
      Dim sql As String

   On Error GoTo LoadLatestISBT128_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE ISBT128 = '" & PackNumber & "' " & _
            "AND BarCode = '" & Barcode & "' "
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

   Exit Sub

LoadLatestISBT128_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "Products", "LoadLatestISBT128", intEL, strES, sql

End Sub

Public Sub LoadLatestBetweenDates(ByVal FromDate As String, _
                                  ByVal ToDate As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadLatestBetweenDates_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE DateTime BETWEEN '" & Format$(FromDate, "dd/MMM/yyyy") & "' " & _
            "AND '" & Format$(ToDate, "dd/MMM/yyyy") & " 23:59' " & _
            "ORDER BY DateTime desc "
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadLatestBetweenDates_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "LoadLatestBetweenDates", intEL, strES, sql

End Sub

Public Sub LoadLatestBetweenExpiryDates(ByVal FromDate As String, _
                                        ByVal ToDate As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadLatestBetweenExpiryDates_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE DateExpiry BETWEEN '" & Format$(FromDate, "dd/MMM/yyyy") & "' " & _
            "AND '" & Format$(ToDate, "dd/MMM/yyyy") & " 23:59' " & _
            "ORDER BY DateTime desc "
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

90    Exit Sub

LoadLatestBetweenExpiryDates_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "LoadLatestBetweenExpiryDates", intEL, strES, sql

End Sub


Public Sub LoadNoAudit(ByVal PackNumber As String, _
                       ByVal Barcode As String, _
                       ByVal DateExpiry As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadNoAudit_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE Number = '" & PackNumber & "' " & _
            "AND BarCode = '" & Barcode & "' " & _
            "AND DateExpiry = '" & Format$(DateExpiry, "dd/MMM/yyyy HH:mm") & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60     Populate tb
        
70     tb.MoveNext
80    Loop

90    Exit Sub

LoadNoAudit_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "LoadNoAudit", intEL, strES, sql

End Sub

Public Sub LoadByUnitNumber(ByVal UnitNumber As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadByUnitNumber_Error

20    sql = "SELECT * FROM Product " & _
            "WHERE Number = '" & UnitNumber & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60     Populate tb
        
70     tb.MoveNext
80    Loop

90    Exit Sub

LoadByUnitNumber_Error:

Dim strES As String
Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Products", "LoadByUnitNumber", intEL, strES, sql

End Sub


Public Sub LoadLatestByUnitNumber(ByVal UnitNumber As String)

      Dim tb As Recordset
      Dim sql As String

   On Error GoTo LoadLatestByUnitNumber_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE Number = '" & UnitNumber & "'"
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

   Exit Sub

LoadLatestByUnitNumber_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "Products", "LoadLatestByUnitNumber", intEL, strES, sql

End Sub


Public Sub LoadLatestByUnitNumberISBT128(ByVal UnitNumber As String)

      Dim tb As Recordset
      Dim sql As String

   On Error GoTo LoadLatestByUnitNumberISBT128_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE ISBT128 = '" & UnitNumber & "'"
      frmXM.txtSQL = sql
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF
60      Populate tb
        
70      tb.MoveNext
80    Loop

   Exit Sub

LoadLatestByUnitNumberISBT128_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "Products", "LoadLatestByUnitNumberISBT128", intEL, strES, sql

End Sub
Public Sub LoadSampleIDNoAudit(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadSampleIDNoAudit_Error

20    sql = "SELECT * FROM Latest " & _
            "WHERE LabNumber = '" & SampleID & "'"
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
120   LogError "Products", "LoadSampleIDNoAudit", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolProducts.[_NewEnum]

End Function


Private Sub Populate(ByVal tb As Recordset)

      Dim p As New Product

10    On Error GoTo Populate_Error

20    With p
30      .PackNumber = tb!Number & ""
40      .Barcode = tb!Barcode & ""
50      .DateExpiry = Format(tb!DateExpiry & "", "dd/mmm/yyyy HH:mm")
60      .ISBT128 = tb!ISBT128 & ""
70      .PackEvent = tb!Event & ""
80      .Chart = tb!Patid & ""
90      .PatName = tb!PatName & ""
100     .UserName = tb!Operator & ""
110     .RecordDateTime = tb!DateTime & ""
120     .GroupRh = tb!GroupRh & ""
130     .Supplier = tb!Supplier & ""
140     .SampleID = tb!LabNumber & ""
150     .crt = tb!crt
160     .cco = tb!cco
170     .cen = tb!cen
180     .crtr = tb!crtr
190     .ccor = tb!ccor
200     .cenr = tb!cenr
210     .Checked = tb!Checked
220     .Notes = tb!Notes & ""
230     .EventStart = tb!EventStart & ""
240     .EventEnd = tb!EventEnd & ""
250     .OrderNumber = tb!OrderNumber & ""
260     .Screen = tb!Screen & ""
270     .Reason = tb!Reason & ""

280     mcolProducts.Add p
        
290   End With

300   Exit Sub

Populate_Error:

Dim strES As String
Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "Products", "Populate", intEL, strES

End Sub

