VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fastings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolFastings As New Collection

Private mFastings As New Fastings

Public Function Add(ByVal Fx As Fasting) _
          As Fasting

          Dim sql As String
          Dim tb As Recordset
          Dim f As Fasting
          Dim Found As Boolean

13920     On Error GoTo Add_Error

13930     sql = "Select * from Fastings where TestName = '" & Fx.TestName & "'"
13940     Set tb = New Recordset
13950     RecOpenServer 0, tb, sql
13960     With Fx
13970         If tb.EOF Then
13980             tb.AddNew
13990         End If

14000         tb!TestName = UCase$(.TestName)
14010         tb!FastingLow = .FastingLow
14020         tb!FastingHigh = .FastingHigh
14030         tb!FastingText = .FastingText
14040         tb.Update

14050         For Each f In mcolFastings
14060             Found = False
14070             If f.TestName = .TestName Then
14080                 f.FastingHigh = .FastingHigh
14090                 f.FastingLow = .FastingLow
14100                 f.FastingText = .FastingText
14110                 Found = True
14120                 Exit For
14130             End If
14140         Next
14150         If Not Found Then
14160             mcolFastings.Add Fx, .TestName
14170         End If
14180     End With

14190     Set Add = Fx

14200     Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer

14210     intEL = Erl
14220     strES = Err.Description
14230     LogError "Fastings", "Add", intEL, strES, sql

End Function
Public Function Count() As Long
        
14240     Count = mcolFastings.Count

End Function




Public Sub Delete(ByVal TestName As String)
        
          Dim sql As String

14250     On Error GoTo Delete_Error

14260     sql = "Delete from Fastings where " & _
              "TestName = '" & TestName & "'"
14270     Cnxn(0).Execute sql

14280     Refresh

14290     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

14300     intEL = Erl
14310     strES = Err.Description
14320     LogError "Fastings", "Delete", intEL, strES, sql

End Sub

Public Function Item(ByVal TestName As String) As Fasting
Attribute Item.VB_UserMemId = 0
        
14330     On Error Resume Next

14340     Set Item = mcolFastings.Item(TestName)

End Function



Private Sub Load()

          Dim tb As Recordset
          Dim sql As String
          Dim Fx As Fasting

14350     On Error GoTo ehlc

14360     sql = "Select * from Fastings"
14370     Set tb = New Recordset
14380     RecOpenClient 0, tb, sql
14390     Do While Not tb.EOF
14400         Set Fx = New Fasting
14410         With Fx
14420             .TestName = UCase$(tb!TestName & "")
14430             .FastingLow = tb!FastingLow & ""
14440             .FastingHigh = tb!FastingHigh & ""
14450             .FastingText = tb!FastingText & ""
14460             mcolFastings.Add Fx, .TestName
14470         End With
14480         tb.MoveNext
14490     Loop

14500     Exit Sub

ehlc:
          Dim er As Long
          Dim ers As String

14510     er = Err.Number
14520     ers = Err.Description

14530     If er = 457 Then Resume Next 'Duplicate key
14540     Screen.MousePointer = 0
14550     MsgBox "Fastings/Load:" & Format$(er) & ":" & ers
14560     Exit Sub

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
14570     Set NewEnum = mcolFastings.[_NewEnum]

End Function


Public Property Get Fastings() As Fastings
        
14580     Set Fastings = mFastings

End Property

Public Sub Refresh()

14590     Set mcolFastings = Nothing

14600     Load

End Sub

Private Sub Class_Initialize()

14610     Load

End Sub


