VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BGAResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Private mcolBGAResults As New Collection
Private mBGAResults As New BGAResults

Public Function Add(ByVal BG As BGAResult) As BGAResult

1170      mcolBGAResults.Add BG

1180      Set Add = BG

End Function


Public Property Get BGAResults() As BGAResults
        
1190      Set BGAResults = mBGAResults

End Property

Public Sub Clear()

          Dim n As Integer

1200      For n = 1 To mcolBGAResults.Count
1210          mcolBGAResults.Remove 1
1220      Next

End Sub




Public Function Count() As Long
        
1230      Count = mcolBGAResults.Count

End Function



Public Function Item(ByVal Index As Long) As BGAResult
Attribute Item.VB_UserMemId = 0

1240      Set Item = mcolBGAResults.Item(Index)

End Function


Public Function LoadResults(ByVal SampleID As String) _
          As BGAResult

          Dim sn As Recordset
          Dim CR As BGAResult
          Dim sql As String

1250      On Error GoTo LoadResults_Error

1260      sql = "select * from BGAresults where " & _
              "sampleid = '" & SampleID & "'"
1270      Set sn = New Recordset
1280      RecOpenServer 0, sn, sql

1290      If Not sn.EOF Then
1300          Set CR = New BGAResult
1310          With CR
1320              .Rundate = sn!Rundate & ""
1330              .RunDateTime = sn!RunDateTime & ""
1340              .pH = Trim$(sn!pH & "")
1350              .PCO2 = Trim$(sn!PCO2 & "")
1360              .PO2 = Trim$(sn!PO2 & "")
1370              .HCO3 = Trim$(sn!HCO3 & "")
1380              .BE = Trim$(sn!BE & "")
1390              .O2SAT = Trim$(sn!O2SAT & "")
1400              .TotCO2 = Trim$(sn!TotCO2 & "")
1410              If Not IsNull(sn!Valid) Then
1420                  .Valid = IIf(sn!Valid, True, False)
1430              Else
1440                  .Valid = False
1450              End If
1460              If Not IsNull(sn!Printed) Then
1470                  .Printed = IIf(sn!Printed, True, False)
1480              Else
1490                  .Printed = False
1500              End If
1510          End With
1520          Set LoadResults = CR
1530      Else
1540          Set LoadResults = Nothing
1550      End If

1560      Exit Function

LoadResults_Error:

          Dim strES As String
          Dim intEL As Integer

1570      intEL = Erl
1580      strES = Err.Description
1590      LogError "BGAResults", "LoadResults", intEL, strES, sql

End Function
Public Sub LogAsPrinted(ByVal SampleID As String)

          Dim sql As String

1600      On Error GoTo LogAsPrinted_Error

1610      sql = "Update BGAresults " & _
              "set Printed = 1 where " & _
              "sampleid = '" & SampleID & "'"
1620      Cnxn(0).Execute sql

1630      Exit Sub

LogAsPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

1640      intEL = Erl
1650      strES = Err.Description
1660      LogError "BGAResults", "LogAsPrinted", intEL, strES, sql


End Sub

Public Sub SaveResults(ByVal CR As BGAResult)

          Dim tb As Recordset
          Dim sql As String

1670      On Error GoTo SaveResults_Error

1680      sql = "Select * from BGAResults where " & _
              "sampleid = '" & Format$(Val(CR.SampleID)) & "'"
1690      Set tb = New Recordset
1700      RecOpenServer 0, tb, sql
1710      tb.AddNew
1720      tb!SampleID = Format$(Val(CR.SampleID))
1730      tb!Rundate = Format(CR.Rundate, "dd/mmm/yyyy")
1740      tb!RunDateTime = Format(CR.RunDateTime, "dd/mmm/yyyy hh:mm:ss")
1750      tb!pH = CR.pH
1760      tb!PCO2 = CR.PCO2
1770      tb!PO2 = CR.PO2
1780      tb!HCO3 = CR.HCO3
1790      tb!BE = CR.BE
1800      tb!O2SAT = CR.O2SAT
1810      tb!TotCO2 = CR.TotCO2
1820      tb!Valid = CR.Valid
1830      tb!Printed = CR.Printed
1840      tb.Update

1850      Exit Sub

SaveResults_Error:

          Dim strES As String
          Dim intEL As Integer

1860      intEL = Erl
1870      strES = Err.Description
1880      LogError "BGAResults", "SaveResults", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
1890      Set NewEnum = mcolBGAResults.[_NewEnum]

End Function

Public Sub Remove(ByVal Index As Long)

1900      mcolBGAResults.Remove (Index)

End Sub





