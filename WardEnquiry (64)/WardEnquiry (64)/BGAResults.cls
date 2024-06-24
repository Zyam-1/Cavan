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

10    mcolBGAResults.Add BG

20    Set Add = BG

End Function


Public Property Get BGAResults() As BGAResults
  
10    Set BGAResults = mBGAResults

End Property

Public Sub Clear()

      Dim n As Integer

10    For n = 1 To mcolBGAResults.Count
20      mcolBGAResults.Remove 1
30    Next

End Sub




Public Function Count() As Long
  
10    Count = mcolBGAResults.Count

End Function



Public Function Item(ByVal Index As Long) As BGAResult
Attribute Item.VB_UserMemId = 0

10    Set Item = mcolBGAResults.Item(Index)

End Function


Public Function LoadResults(ByVal Cn As Integer, ByVal SampleID As String) _
                            As BGAResult

      Dim sn As Recordset
      Dim CR As BGAResult
      Dim sql As String



10    sql = "select * from BGAresults where " & _
            "sampleid = '" & SampleID & "'"
20    Set sn = New Recordset
30    RecOpenServer Cn, sn, sql

40    If Not sn.EOF Then
50      Set CR = New BGAResult
60      With CR
70        .RunDate = sn!RunDate & ""
80        .RunDateTime = sn!RunDateTime & ""
90        .pH = sn!pH & ""
100       .PCO2 = sn!PCO2 & ""
110       .PO2 = sn!PO2 & ""
120       .HCO3 = sn!HCO3 & ""
130       .BE = sn!BE & ""
140       .O2SAT = sn!O2SAT & ""
150       .TotCO2 = sn!TotCO2 & ""
160       If Not IsNull(sn!Valid) Then
170         .Valid = IIf(sn!Valid, True, False)
180       Else
190         .Valid = False
200       End If
210       If Not IsNull(sn!Printed) Then
220         .Printed = IIf(sn!Printed, True, False)
230       Else
240         .Printed = False
250       End If
260     End With
270     Set LoadResults = CR
280   Else
290     Set LoadResults = Nothing
300   End If




End Function
Public Sub LogAsPrinted(ByVal SampleID As String)

      Dim sql As String


10    sql = "Update BGAresults " & _
            "set Printed = 1 where " & _
            "sampleid = '" & SampleID & "'"
20    Cnxn(0).Execute sql

End Sub

Public Sub SaveResults(ByVal Cn As Integer, ByVal CR As BGAResult)

      Dim tb As Recordset
      Dim sql As String
  


10    sql = "Select * from BGAResults where " & _
            "sampleid = '" & Format$(Val(CR.SampleID)) & "'"
20    Set tb = New Recordset
30    RecOpenServer Cn, tb, sql
40    tb.AddNew
50    tb!SampleID = Format$(Val(CR.SampleID))
60    tb!RunDate = Format(CR.RunDate, "dd/mmm/yyyy")
70    tb!RunDateTime = Format(CR.RunDateTime, "dd/mmm/yyyy hh:mm:ss")
80    tb!pH = CR.pH
90    tb!PCO2 = CR.PCO2
100   tb!PO2 = CR.PO2
110   tb!HCO3 = CR.HCO3
120   tb!BE = CR.BE
130   tb!O2SAT = CR.O2SAT
140   tb!TotCO2 = CR.TotCO2
150   tb!Valid = CR.Valid
160   tb!Printed = CR.Printed
170   tb.Update



End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolBGAResults.[_NewEnum]

End Function

Public Sub Remove(ByVal Index As Long)

10    mcolBGAResults.Remove (Index)

End Sub





