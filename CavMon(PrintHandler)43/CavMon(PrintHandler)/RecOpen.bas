Attribute VB_Name = "RecOpen"
Option Explicit

Public Sub RecOpenServer(ByVal n As Long, _
                         ByRef tb As Recordset, _
                         ByVal sql As String)
    
10    With tb
20      .CursorLocation = adUseServer
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = Cnxn(n)
60      .Source = sql
70      .Open
80    End With

End Sub


Public Sub RecOpenServerBB(ByVal n As Long, _
                           ByRef tb As Recordset, _
                           ByVal sql As String)
    
10    With tb
20      .CursorLocation = adUseServer
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = CnxnBB(n)
60      .Source = sql
70      .Open
80    End With

End Sub

Public Sub RecOpenClientBB(ByVal n As Long, _
                           ByRef tb As Recordset, _
                           ByVal sql As String)
    
10    With tb
20      .CursorLocation = adUseClient
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = CnxnBB(n)
60      .Source = sql
70      .Open
80    End With

End Sub


Public Sub RecOpenClient(ByVal n As Long, _
                         ByRef tb As Recordset, _
                         ByVal sql As String)
    
10    With tb
20      .CursorLocation = adUseClient
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = Cnxn(n)
60      .Source = sql
70      .Open
80    End With

End Sub


