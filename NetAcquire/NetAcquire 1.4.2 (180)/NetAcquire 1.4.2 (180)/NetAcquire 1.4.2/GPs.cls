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
        
16320     Count = mcolGPs.Count

End Function


Public Property Get GPs() As GPs
        
16330     Set GPs = mGPs

End Property

Public Function Item(ByVal GPName As String) As GP
Attribute Item.VB_UserMemId = 0
        
16340     On Error Resume Next

16350     Set Item = mcolGPs.Item(GPName)

End Function

Public Sub LoadByMcNumber(ByVal HospitalCode As String, ByVal McNumber As String, ByVal OnlyInUse As Boolean)

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As GP

16360     On Error GoTo LoadByMcNumber_Error

16370     sql = "SELECT G.Text, G.Code, G.Addr0, G.Addr1, " & _
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
16380     If OnlyInUse Then
16390         sql = sql & "AND InUse = '1' "
16400     End If
16410     sql = sql & "ORDER BY ListOrder"
16420     Set tb = New Recordset
16430     RecOpenClient 0, tb, sql
16440     With tb
16450         Do While Not .EOF
16460             Set Gx = New GP
16470             Gx.Text = !Text & ""
16480             Gx.Code = !Code & ""
16490             Gx.Addr0 = !Addr0 & ""
16500             Gx.Addr1 = !Addr1 & ""
16510             Gx.InUse = !InUse
16520             Gx.Title = !Title & ""
16530             Gx.ForeName = !ForeName & ""
16540             Gx.SurName = !SurName & ""
16550             Gx.Phone = !Phone & ""
16560             Gx.FAX = !FAX & ""
16570             Gx.Practice = !Practice & ""
16580             Gx.Compiled = !Compiled
16590             Gx.HospitalCode = !HospitalCode & ""
16600             Gx.ListOrder = !ListOrder
16610             Gx.HealthLink = !HealthLink
16620             Gx.McNumber = !McNumber & ""
16630             Gx.PracticeNumber = !PracticeNumber & ""
16640             Gx.PrintReport = !PrintReport
16650             Gx.EGFR = !EGFR
16660             Gx.AutoCC = !AutoCC
16670             mcolGPs.Add Gx, Gx.Text
16680             .MoveNext
16690         Loop
16700     End With

16710     Exit Sub

LoadByMcNumber_Error:

          Dim strES As String
          Dim intEL As Integer

16720     intEL = Erl
16730     strES = Err.Description
16740     LogError "GPs", "LoadByMcNumber", intEL, strES, sql

End Sub

Public Sub Load(ByVal HospitalCode As String, ByVal OnlyInUse As Boolean)

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As GP

16750     On Error GoTo Load_Error

16760     sql = "SELECT G.Text, G.Code, G.Addr0, G.Addr1, " & _
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
16770     If OnlyInUse Then
16780         sql = sql & "AND InUse = '1' "
16790     End If
16800     sql = sql & "ORDER BY ListOrder"
16810     Set tb = New Recordset
16820     RecOpenClient 0, tb, sql
16830     With tb
16840         Do While Not .EOF
16850             Set Gx = New GP
16860             Gx.Text = !Text & ""
16870             Gx.Code = !Code & ""
16880             Gx.Addr0 = !Addr0 & ""
16890             Gx.Addr1 = !Addr1 & ""
16900             Gx.InUse = !InUse
16910             Gx.Title = !Title & ""
16920             Gx.ForeName = !ForeName & ""
16930             Gx.SurName = !SurName & ""
16940             Gx.Phone = !Phone & ""
16950             Gx.FAX = !FAX & ""
16960             Gx.Practice = !Practice & ""
16970             Gx.Compiled = !Compiled
16980             Gx.HospitalCode = !HospitalCode & ""
16990             Gx.ListOrder = !ListOrder
17000             Gx.HealthLink = !HealthLink
17010             Gx.McNumber = !McNumber & ""
17020             Gx.PracticeNumber = !PracticeNumber & ""
17030             Gx.PrintReport = !PrintReport
17040             Gx.EGFR = !EGFR
17050             Gx.AutoCC = !AutoCC
17060             Gx.Interim = !Interim
                  '330           mcolGPs.Add Gx, Gx.Text
                  'Zyam added exception handling in case of duplication
                  ' Only add if the key does not exist
17070             mcolGPs.Add Gx, Gx.Text
                  'Zyam 20-05-24
17080             .MoveNext
17090         Loop
17100     End With

17110     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

17120     intEL = Erl
17130     strES = Err.Description
          'LogError "GPs", "Load", intEL, strES, Sql

End Sub
Public Sub LoadListFaxNumbers()

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As GP

17140     On Error GoTo LoadListFaxNumbers_Error

17150     sql = "SELECT DISTINCT FAX FROM GPs"

17160     Set tb = New Recordset
17170     RecOpenClient 0, tb, sql
17180     Do While Not tb.EOF
17190         Set Gx = New GP
17200         Gx.FAX = tb!FAX & ""
17210         mcolGPs.Add Gx
17220         tb.MoveNext
17230     Loop

17240     Exit Sub

LoadListFaxNumbers_Error:

          Dim strES As String
          Dim intEL As Integer

17250     intEL = Erl
17260     strES = Err.Description
17270     LogError "GPs", "LoadListFaxNumbers", intEL, strES, sql

End Sub
Public Sub LoadListo(ByVal HospitalCode As String, ByVal OnlyInUse As Boolean)

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As GP

17280     On Error GoTo LoadList_Error

17290     sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM GPs WHERE " & _
              "HospitalCode = '" & HospitalCode & "' "
17300     If OnlyInUse Then
17310         sql = sql & "AND InUse = 1 "
17320     End If
17330     sql = sql & "GROUP BY Text " & _
              "ORDER BY L"

17340     Set tb = New Recordset
17350     RecOpenClient 0, tb, sql
17360     Do While Not tb.EOF
17370         Set Gx = New GP
17380         Gx.Text = tb!Text & ""
17390         mcolGPs.Add Gx, Gx.Text
17400         tb.MoveNext
17410     Loop

17420     Exit Sub

LoadList_Error:

          Dim strES As String
          Dim intEL As Integer

17430     intEL = Erl
17440     strES = Err.Description
17450     LogError "GPs", "LoadList", intEL, strES, sql

End Sub


Public Function GpCCed(ByVal HospitalCode As String, ByVal strGP As String) As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As GP

17460     On Error GoTo GpCCed_Error

17470     GpCCed = False

17480     sql = "SELECT AutoCC FROM GPs WHERE " & _
              "HospitalCode = '" & HospitalCode & "' "
17490     sql = sql & "AND Text = '" & AddTicks(strGP) & "' and AutoCC = 1 "

17500     Set tb = New Recordset
17510     RecOpenClient 0, tb, sql
17520     If Not tb.EOF Then
17530         GpCCed = True
17540     Else
17550         GpCCed = False
17560     End If

17570     Exit Function

GpCCed_Error:

          Dim strES As String
          Dim intEL As Integer

17580     intEL = Erl
17590     strES = Err.Description
17600     LogError "GPs", "GpCCed", intEL, strES, sql

End Function




Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
17610     Set NewEnum = mcolGPs.[_NewEnum]

End Function




Public Sub Save(ByVal Gx As GP)

          Dim sql As String

17620     On Error GoTo Save_Error

17630     sql = "IF EXISTS (SELECT * FROM GPs WHERE " & _
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
17640     sql = sql & "ELSE " & _
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
17650     Cnxn(0).Execute sql

17660     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

17670     intEL = Erl
17680     strES = Err.Description
17690     LogError "GPs", "Save", intEL, strES, sql

End Sub




