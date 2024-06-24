Attribute VB_Name = "RecOpen"
Option Explicit

Public Sub RecOpenServer(ByVal n As Long, _
                         ByRef tb As Recordset, _
                         ByVal sql As String)
          
27290 With tb
27300   .CursorLocation = adUseServer
27310   .CursorType = adOpenDynamic
27320   .LockType = adLockOptimistic
27330   .ActiveConnection = Cnxn(n)
27340   .Source = sql
27350   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27360 End With

End Sub


Public Sub RecOpenServerRemote(ByRef tb As Recordset, _
                               ByVal sql As String)
          
27370 With tb
27380   .CursorLocation = adUseServer
27390   .CursorType = adOpenDynamic
27400   .LockType = adLockOptimistic
27410   .ActiveConnection = CnxnRemote
27420   .Source = sql
27430   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27440 End With

End Sub



Public Sub RecOpenServerRemoteBB(ByRef tb As Recordset, _
                                 ByVal sql As String)
          
27450 With tb
27460   .CursorLocation = adUseServer
27470   .CursorType = adOpenDynamic
27480   .LockType = adLockOptimistic
27490   .ActiveConnection = CnxnRemoteBB(0)
27500   .Source = sql
27510   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27520 End With

End Sub

Public Sub RecOpenServerBB(ByVal n As Long, _
                           ByRef tb As Recordset, _
                           ByVal sql As String)
          
27530 With tb
27540   .CursorLocation = adUseServer
27550   .CursorType = adOpenDynamic
27560   .LockType = adLockOptimistic
27570   .ActiveConnection = CnxnBB(n)
27580   .Source = sql
27590   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27600 End With

End Sub

Public Sub RecOpenClientBB(ByVal n As Long, _
                           ByRef tb As Recordset, _
                           ByVal sql As String)
          
27610 With tb
27620   .CursorLocation = adUseClient
27630   .CursorType = adOpenDynamic
27640   .LockType = adLockOptimistic
27650   .ActiveConnection = CnxnBB(n)
27660   .Source = sql
27670   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27680 End With

End Sub


Public Sub RecOpenClienty(ByVal n As Long, _
                         ByRef tb As Recordset, _
                         ByVal sql As String)

      Dim cmd As New ADODB.Command

27690 cmd.CommandType = adCmdText
27700 cmd.CommandText = sql
27710 Set cmd.ActiveConnection = Cnxn(0)
27720 Set tb = cmd.Execute


      '10    With tb
      '20      .CursorLocation = adUseClient
      '30      .CursorType = adOpenDynamic
      '40      .LockType = adLockOptimistic
      '50      .ActiveConnection = Cnxn(n)
      '60      .Source = sql
      '70      .Open
      '80    End With

End Sub


Public Sub RecOpenClient(ByVal n As Long, _
                         ByRef tb As Recordset, _
                         ByVal sql As String)
          
27730 With tb
27740   .CursorLocation = adUseClient
27750   .CursorType = adOpenDynamic
27760   .LockType = adLockOptimistic
27770   .ActiveConnection = Cnxn(n)
27780   .Source = sql
27790   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27800 End With

End Sub

Public Sub RecOpenClientRemote(ByRef tb As Recordset, _
                               ByVal sql As String)
          
27810 With tb
27820   .CursorLocation = adUseClient
27830   .CursorType = adOpenDynamic
27840   .LockType = adLockOptimistic
27850   .ActiveConnection = CnxnRemote(0)
27860   .Source = sql
27870   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27880 End With

End Sub


Public Sub RecOpenClientRemoteBB(ByRef tb As Recordset, _
                                 ByVal sql As String)
          
27890 With tb
27900   .CursorLocation = adUseClient
27910   .CursorType = adOpenDynamic
27920   .LockType = adLockOptimistic
27930   .ActiveConnection = CnxnRemoteBB(0)
27940   .Source = sql
27950   .Open
      '+++ Junaid 12-12-2023
      '
      '--- Junaid
27960 End With

End Sub

