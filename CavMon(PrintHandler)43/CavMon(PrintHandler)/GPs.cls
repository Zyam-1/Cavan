VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolGPs As New Collection

Private mGPs As New GPs

Public Function Count() As Long
  
10    Count = mcolGPs.Count

End Function


Public Property Get GPs() As GPs
  
10    Set GPs = mGPs

End Property

Public Function Item(ByVal GPName As String) As GP
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolGPs.Item(GPName)

End Function

Public Sub LoadByMcNumber(ByVal HospitalCode As String, ByVal McNumber As String, ByVal OnlyInUse As Boolean)

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As GP

10    On Error GoTo LoadByMcNumber_Error

20    sql = "SELECT G.Text, G.Code, G.Addr0, G.Addr1, " & _
            "G.Title, G.ForeName, G.SurName, G.Phone, G.FAX, G.Practice, " & _
            "G.HospitalCode, COALESCE(G.ListOrder, 999) ListOrder, G.MCNumber, G.PracticeNumber, " & _
            "COALESCE(G.InUse, 1) InUse, " & _
            "COALESCE(G.Compiled, 1) Compiled, " & _
            "COALESCE(G.HealthLink, 1) Healthlink, " & _
            "COALESCE(G.Interim, 1) Interim, " & _
            "COALESCE(G.PrintReport, 1) PrintReport, " & _
            "COALESCE(E.Include, 0) EGFR, AutoCC " & _
            "FROM GPs G JOIN IncludeEGFR E " & _
            "ON G.Text = E.SourceName " & _
            "WHERE HospitalCode = '" & HospitalCode & "' " & _
            "AND E.SourceType = 'GP' AND MCNumber = '" & McNumber & "' "
30    If OnlyInUse Then
40        sql = sql & "AND InUse = '1' "
50    End If
60    sql = sql & "ORDER BY ListOrder"
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    With tb
100       Do While Not .EOF
110           Set Gx = New GP
120           Gx.Text = !Text & ""
130           Gx.Code = !Code & ""
140           Gx.Addr0 = !Addr0 & ""
150           Gx.Addr1 = !Addr1 & ""
160           Gx.InUse = !InUse
170           Gx.Title = !Title & ""
180           Gx.ForeName = !ForeName & ""
190           Gx.SurName = !SurName & ""
200           Gx.Phone = !Phone & ""
210           Gx.FAX = !FAX & ""
220           Gx.Practice = !Practice & ""
230           Gx.Compiled = !Compiled
240           Gx.HospitalCode = !HospitalCode & ""
250           Gx.ListOrder = !ListOrder
260           Gx.HealthLink = !HealthLink
270           Gx.McNumber = !McNumber & ""
280           Gx.PracticeNumber = !PracticeNumber & ""
290           Gx.PrintReport = !PrintReport
300           Gx.EGFR = !EGFR
310           Gx.AutoCC = !AutoCC
320           mcolGPs.Add Gx, Gx.Text
330           .MoveNext
340       Loop
350   End With

360   Exit Sub

LoadByMcNumber_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "GPs", "LoadByMcNumber", intEL, strES, sql

End Sub

Public Sub Load(ByVal HospitalCode As String, ByVal OnlyInUse As Boolean)

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As GP

10    On Error GoTo Load_Error

20    sql = "SELECT G.Text, G.Code, G.Addr0, G.Addr1, " & _
            "G.Title, G.ForeName, G.SurName, G.Phone, G.FAX, G.Practice, " & _
            "G.HospitalCode, COALESCE(G.ListOrder, 999) ListOrder, G.MCNumber, G.PracticeNumber, " & _
            "COALESCE(G.InUse, 1) InUse, " & _
            "COALESCE(G.Compiled, 1) Compiled, " & _
            "COALESCE(G.HealthLink, 1) Healthlink, " & _
            "COALESCE(G.PrintReport, 1) PrintReport, " & _
            "COALESCE(E.Include, 0) EGFR, AutoCC, " & _
            "COALESCE(G.Interim, 1) Interim " & _
            "FROM GPs G JOIN IncludeEGFR E " & _
            "ON G.Text = E.SourceName " & _
            "WHERE HospitalCode = '" & HospitalCode & "' " & _
            "AND E.SourceType = 'GP' "
30    If OnlyInUse Then
40        sql = sql & "AND InUse = '1' "
50    End If
60    sql = sql & "ORDER BY ListOrder"
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    With tb
100       Do While Not .EOF
110           Set Gx = New GP
120           Gx.Text = !Text & ""
130           Gx.Code = !Code & ""
140           Gx.Addr0 = !Addr0 & ""
150           Gx.Addr1 = !Addr1 & ""
160           Gx.InUse = !InUse
170           Gx.Title = !Title & ""
180           Gx.ForeName = !ForeName & ""
190           Gx.SurName = !SurName & ""
200           Gx.Phone = !Phone & ""
210           Gx.FAX = !FAX & ""
220           Gx.Practice = !Practice & ""
230           Gx.Compiled = !Compiled
240           Gx.HospitalCode = !HospitalCode & ""
250           Gx.ListOrder = !ListOrder
260           Gx.HealthLink = !HealthLink
270           Gx.McNumber = !McNumber & ""
280           Gx.PracticeNumber = !PracticeNumber & ""
290           Gx.PrintReport = !PrintReport
300           Gx.EGFR = !EGFR
310           Gx.AutoCC = !AutoCC
320           Gx.Interim = !Interim
330           mcolGPs.Add Gx, Gx.Text
340           .MoveNext
350       Loop
360   End With

370   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "GPs", "Load", intEL, strES, sql

End Sub
Public Sub LoadListFaxNumbers()

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As GP

10    On Error GoTo LoadListFaxNumbers_Error

20    sql = "SELECT DISTINCT FAX FROM GPs"

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Gx = New GP
70      Gx.FAX = tb!FAX & ""
80      mcolGPs.Add Gx
90      tb.MoveNext
100   Loop

110   Exit Sub

LoadListFaxNumbers_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "GPs", "LoadListFaxNumbers", intEL, strES, sql

End Sub
Public Sub LoadListo(ByVal HospitalCode As String, ByVal OnlyInUse As Boolean)

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As GP

10    On Error GoTo LoadList_Error

20    sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM GPs WHERE " & _
            "HospitalCode = '" & HospitalCode & "' "
30    If OnlyInUse Then
40      sql = sql & "AND InUse = 1 "
50    End If
60    sql = sql & "GROUP BY Text " & _
            "ORDER BY L"

70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    Do While Not tb.EOF
100     Set Gx = New GP
110     Gx.Text = tb!Text & ""
120     mcolGPs.Add Gx, Gx.Text
130     tb.MoveNext
140   Loop

150   Exit Sub

LoadList_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "GPs", "LoadList", intEL, strES, sql

End Sub


Public Function GpCCed(ByVal HospitalCode As String, ByVal strGP As String) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As GP

10    On Error GoTo GpCCed_Error

20    GpCCed = False

30    sql = "SELECT AutoCC FROM GPs WHERE " & _
            "HospitalCode = '" & HospitalCode & "' "
40    sql = sql & "AND Text = '" & strGP & "' and AutoCC = 1 "

50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If Not tb.EOF Then
80      GpCCed = True
90    Else
100     GpCCed = False
110   End If

120   Exit Function

GpCCed_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "GPs", "GpCCed", intEL, strES, sql

End Function




Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolGPs.[_NewEnum]

End Function




Public Sub Save(ByVal Gx As GP)

      Dim sql As String

10    On Error GoTo Save_Error

20    sql = "IF EXISTS (SELECT * FROM GPs WHERE " & _
            "         Code = '" & Gx.Code & "') " & _
            "  UPDATE GPs " & _
            "  SET Text = '" & Gx.Text & "', " & _
            "  Addr0 = '" & AddTicks(Gx.Addr0) & "', " & _
            "  Addr1 = '" & AddTicks(Gx.Addr1) & "', " & _
            "  InUse = '" & Gx.InUse & "', " & _
            "  Title = '" & Gx.Title & "', " & _
            "  ForeName = '" & AddTicks(Gx.ForeName) & "', " & _
            "  SurName = '" & AddTicks(Gx.SurName) & "', " & _
            "  Phone = '" & Gx.Phone & "', " & _
            "  Fax = '" & Gx.FAX & "', " & _
            "  Practice = '" & Gx.Practice & "', " & _
            "  Compiled = '" & Gx.Compiled & "', " & _
            "  HospitalCode = '" & Gx.HospitalCode & "', " & _
            "  ListOrder = '" & Gx.ListOrder & "', " & _
            "  HealthLink = '" & Gx.HealthLink & "', " & _
            "  MCNumber = '" & Gx.McNumber & "', " & _
            "  PrintReport = '" & Gx.PrintReport & "' " & _
      "  WHERE Code = '" & Gx.Code & "' "
30    sql = sql & "ELSE " & _
            "  INSERT INTO GPs(Text, Code, Addr0, Addr1, InUse, " & _
            "                  Title, ForeName, SurName, Phone, FAX, " & _
            "                  Practice, Compiled, HospitalCode, ListOrder, " & _
            "                  Healthlink, MCNumber, PrintReport) " & _
            "  VALUES " & _
            "  ('" & Gx.Text & "', " & _
            "  '" & Gx.Code & "', " & _
            "  '" & AddTicks(Gx.Addr0) & "', " & _
            "  '" & AddTicks(Gx.Addr1) & "', " & _
            "  '" & Gx.InUse & "', " & _
            "  '" & Gx.Title & "', " & _
            "  '" & AddTicks(Gx.ForeName) & "', " & _
            "  '" & AddTicks(Gx.SurName) & "', " & _
            "  '" & Gx.Phone & "', " & _
            "  '" & Gx.FAX & "', " & _
            "  '" & Gx.Practice & "', " & _
            "  '" & Gx.Compiled & "', " & _
            "  '" & Gx.HospitalCode & "', " & _
            "  '" & Gx.ListOrder & "', " & _
            "  '" & Gx.HealthLink & "', " & _
            "  '" & Gx.McNumber & "', " & _
            "  '" & Gx.PrintReport & "')"
40    Cnxn(0).Execute sql

50    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "GPs", "Save", intEL, strES, sql

End Sub



