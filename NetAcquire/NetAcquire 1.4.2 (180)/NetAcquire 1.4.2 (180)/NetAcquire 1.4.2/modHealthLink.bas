Attribute VB_Name = "modHealthLink"
Option Explicit

Public Sub ReleaseMicro(ByVal SampleID As String, ByVal Release As Boolean)

      Dim sql As String

1190  On Error GoTo ReleaseMicro_Error

1200  If Release Then
1210    sql = "UPDATE Demographics " & _
              "SET ForMicro = '1', " & _
              "MicroHealthLinkReleaseTime = getdate() " & _
              "WHERE SampleID = '" & SampleID & "'"
1220  Else
1230    sql = "UPDATE Demographics " & _
              "SET ForMicro = '0', " & _
              "MicroHealthLinkReleaseTime = NULL " & _
              "WHERE SampleID = '" & SampleID & "'"
1240  End If

1250  Cnxn(0).Execute sql

1260  Exit Sub

ReleaseMicro_Error:

      Dim strES As String
      Dim intEL As Integer

1270  intEL = Erl
1280  strES = Err.Description
1290  LogError "modHealthLink", "ReleaseMicro", intEL, strES, sql

End Sub


Public Function IsMicroReleased(ByVal SampleID As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

1300  On Error GoTo IsMicroReleased_Error

1310  IsMicroReleased = False

1320  sql = "SELECT COUNT(*) Tot FROM Demographics " & _
            "WHERE COALESCE(ForMicro, 0) <> 0 " & _
            "AND SampleID = '" & SampleID & "'"
1330  Set tb = Cnxn(0).Execute(sql)
1340  If tb!Tot > 0 Then
1350    IsMicroReleased = True
1360  End If

1370  Exit Function

IsMicroReleased_Error:

      Dim strES As String
      Dim intEL As Integer

1380  intEL = Erl
1390  strES = Err.Description
1400  LogError "modHealthLink", "IsMicroReleased", intEL, strES, sql

End Function


