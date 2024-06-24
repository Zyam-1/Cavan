Attribute VB_Name = "modRecOpen"
Option Explicit

Public Sub RecClose(ByVal RS As Recordset)

10    RS.Close
20    Set RS = Nothing

End Sub


Public Sub RecOpenClient(ByVal Cn As Integer, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20      .CursorLocation = adUseClient
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = Cnxn(Cn)
60      .Source = sql
70      .Open
80    End With


End Sub


Public Sub RecOpenServer(ByVal Cn As Integer, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20      .CursorLocation = adUseServer
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = Cnxn(Cn)
60      .Source = sql
70      .Open
80    End With

End Sub



Public Sub RecOpenClientBB(ByVal Cn As Integer, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20      .CursorLocation = adUseClient
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = CnxnBB(Cn)
60      .Source = sql
70      .Open
80    End With

End Sub




Public Sub RecOpenServerBB(ByVal Cn As Integer, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20      .CursorLocation = adUseServer
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = CnxnBB(Cn)
60      .Source = sql
70      .Open
80    End With

End Sub


