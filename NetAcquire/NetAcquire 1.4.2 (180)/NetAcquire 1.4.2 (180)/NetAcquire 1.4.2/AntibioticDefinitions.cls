VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ABDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolABDefinitions As New Collection

Private mABDefinitions As New ABDefinitions

Public Function Count() As Long
        
490       Count = mcolABDefinitions.Count

End Function




Public Sub Delete(ByVal AB As ABDefinition)
        
          Dim sql As String

500       On Error GoTo Delete_Error

510       sql = "DELETE FROM ABDefinitions WHERE " & _
              "AntibioticName = '" & AB.AntibioticName & "' " & _
              "AND OrganismGroup = '" & AB.OrganismGroup & "' " & _
              "AND Site = '" & AB.Site & "'"
520       Cnxn(0).Execute sql

530       Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "ABDefinitions", "Delete", intEL, strES, sql

End Sub

Public Property Get ABDefinitions() As ABDefinitions
        
570       Set ABDefinitions = mABDefinitions

End Property


Public Function Item(ByVal Site As String, ByVal OrganismGroup As String, ByVal AntibioticName As String) As ABDefinition
        

          Dim i As Integer

580       On Error GoTo Item_Error

590       Set Item = Nothing
600       For i = 1 To mcolABDefinitions.Count
610           If mcolABDefinitions.Item(i).Site = Site And mcolABDefinitions.Item(i).OrganismGroup = OrganismGroup _
                  And mcolABDefinitions.Item(i).AntibioticName = AntibioticName Then
620               Set Item = mcolABDefinitions(i)
630           End If
640       Next i


650       Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

660       intEL = Erl
670       strES = Err.Description
680       LogError "ABDefinitions", "Item", intEL, strES
          
End Function


Public Sub Load(ByVal Site As String, ByVal OrganismGroup As String)

          Dim tb As Recordset
          Dim sql As String
          Dim AB As ABDefinition
          Dim sx As Sensitivity

690       On Error GoTo Load_Error

700       sql = "SELECT A.Code, D.AntibioticName, D.OrganismGroup, D.Site, D.ListOrder, D.PriSec, " & _
              "COALESCE(A.AllowIfPregnant, 0) AllowIfPregnant, " & _
              "COALESCE(A.AllowIfOutPatient, 0) AllowIfOutPatient, " & _
              "COALESCE(A.AllowIfChild, 0) AllowIfChild, " & _
              "COALESCE(A.AllowIfPenAll, 0) AllowIfPenAll, " & _
              "COALESCE(A.ViewInGrid, 0) ViewInGrid " & _
              "FROM ABDefinitions D JOIN Antibiotics A " & _
              "ON D.AntibioticName = A.AntibioticName " & _
              "WHERE D.OrganismGroup = '" & OrganismGroup & "' " & _
              "AND D.Site = '" & Site & "' " & _
              "ORDER BY D.ListOrder"
710       Set tb = New Recordset
720       RecOpenClient 0, tb, sql
730       If tb.EOF Then
740           sql = "SELECT A.Code, D.AntibioticName, D.OrganismGroup, D.Site, D.ListOrder, D.PriSec, " & _
                  "COALESCE(A.AllowIfPregnant, 0) AllowIfPregnant, " & _
                  "COALESCE(A.AllowIfOutPatient, 0) AllowIfOutPatient, " & _
                  "COALESCE(A.AllowIfChild, 0) AllowIfChild, " & _
                  "COALESCE(A.AllowIfPenAll, 0) AllowIfPenAll, " & _
                  "COALESCE(A.ViewInGrid, 0) ViewInGrid " & _
                  "FROM ABDefinitions D JOIN Antibiotics A " & _
                  "ON D.AntibioticName = A.AntibioticName " & _
                  "WHERE D.OrganismGroup = '" & OrganismGroup & "' " & _
                  "AND D.Site = 'Generic' " & _
                  "ORDER BY D.ListOrder"
750           Set tb = New Recordset
760           RecOpenClient 0, tb, sql
770       End If
780       Do While Not tb.EOF
790           Set AB = New ABDefinition
800           With AB
810               .Code = tb!Code & ""
820               .AntibioticName = tb!AntibioticName & ""
830               .OrganismGroup = tb!OrganismGroup & ""
840               .Site = tb!Site & ""

850               .ListOrder = tb!ListOrder
860               .PriSec = tb!PriSec & ""
870               .AllowIfPregnant = tb!AllowIfPregnant
880               .AllowIfOutPatient = tb!AllowIfOutPatient
890               .AllowIfChild = tb!AllowIfChild
900               .AllowIfPenicillinAllergy = tb!AllowIfPenAll
910               .ViewInGrid = tb!ViewInGrid
920               mcolABDefinitions.Add AB, .Site & .OrganismGroup & .AntibioticName
930           End With
940           tb.MoveNext
950       Loop

960       Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

970       intEL = Erl
980       strES = Err.Description
990       LogError "ABDefinitions", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
1000      Set NewEnum = mcolABDefinitions.[_NewEnum]

End Function

Public Sub Save(ByVal AB As ABDefinition)

          Dim sql As String

1010      On Error GoTo Save_Error

1020      sql = "IF EXISTS (SELECT * FROM ABDefinitions WHERE " & _
              "           AntibioticName = '" & AB.AntibioticName & "' " & _
              "           AND OrganismGroup = '" & AB.OrganismGroup & "' " & _
              "           AND Site = '" & AB.Site & "') " & _
              "  UPDATE ABDefinitions " & _
              "  SET ListOrder = '" & AB.ListOrder & "', " & _
              "  PriSec = '" & AB.PriSec & "' " & _
              "  WHERE AntibioticName = '" & AB.AntibioticName & "' " & _
              "  AND OrganismGroup = '" & AB.OrganismGroup & "' " & _
              "  AND Site = '" & AB.Site & "' "
1030      sql = sql & "ELSE " & _
              "  INSERT INTO ABDefinitions (AntibioticName, OrganismGroup, Site, ListOrder, PriSec ) " & _
              "  VALUES " & _
              "  ('" & AB.AntibioticName & "', " & _
              "  '" & AB.OrganismGroup & "', " & _
              "  '" & AB.Site & "', " & _
              "  '" & AB.ListOrder & "', " & _
              "  '" & AB.PriSec & "')"
1040      Cnxn(0).Execute sql

1050      Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

1060      intEL = Erl
1070      strES = Err.Description
1080      LogError "ABDefinitions", "Save", intEL, strES, sql

End Sub


Public Sub SaveAllowable(ByVal AB As ABDefinition)

          Dim sql As String

1090      On Error GoTo SaveAllowable_Error

1100      sql = "IF EXISTS (SELECT * FROM Antibiotics WHERE " & _
              "           AntibioticName = '" & AB.AntibioticName & "') " & _
              "  UPDATE Antibiotics " & _
              "  SET ListOrder = '" & AB.ListOrder & "', " & _
              "  AllowIfPregnant = '" & AB.AllowIfPregnant & "', " & _
              "  AllowIfOutPatient = '" & AB.AllowIfOutPatient & "', " & _
              "  AllowIfChild = '" & AB.AllowIfChild & "', " & _
              "  AllowIfPenAll = '" & AB.AllowIfPenicillinAllergy & "' " & _
              "  WHERE AntibioticName = '" & AB.AntibioticName & "' "
1110      sql = sql & "ELSE " & _
              "  INSERT INTO Antibiotics (Code, AntibioticName, ListOrder, AllowIfPregnant, AllowIfOutPatient, AllowIfChild, AllowIfPenAll)" & _
              "  VALUES " & _
              "  ('" & AB.Code & "', " & _
              "  '" & AB.AntibioticName & "', " & _
              "  '" & AB.ListOrder & "', " & _
              "  '" & AB.AllowIfPregnant & "', " & _
              "  '" & AB.AllowIfOutPatient & "', " & _
              "  '" & AB.AllowIfChild & "', " & _
              "  '" & AB.AllowIfPenicillinAllergy & "')"
1120      Cnxn(0).Execute sql

1130      Exit Sub

SaveAllowable_Error:

          Dim strES As String
          Dim intEL As Integer

1140      intEL = Erl
1150      strES = Err.Description
1160      LogError "ABDefinitions", "SaveAllowable", intEL, strES, sql

End Sub



