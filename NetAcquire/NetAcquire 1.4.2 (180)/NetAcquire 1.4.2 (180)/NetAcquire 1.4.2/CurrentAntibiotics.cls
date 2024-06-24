VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrentAntibiotics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolCurrentAntibiotics As New Collection

Private mCurrentAntibiotics As New CurrentAntibiotics

Public Function Count() As Long
        
8370      Count = mcolCurrentAntibiotics.Count

End Function

Public Sub Delete(ByVal SampleID As String, ByVal AB As String)
        
          Dim sql As String
          Dim CURS As New CurrentAntibiotics

8380      On Error GoTo Delete_Error

8390      CURS.Load SampleID

8400      sql = "DELETE FROM CurrentAntibiotics50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND Antibiotic = '" & AB & "'"
8410      Cnxn(0).Execute sql

8420      Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

8430      intEL = Erl
8440      strES = Err.Description
8450      LogError "CurrentAntibiotics", "Delete", intEL, strES, sql

End Sub
Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ix As CurrentAntibiotic

8460      On Error GoTo Load_Error

8470      sql = "SELECT SampleID, COALESCE(Entry, 0) Entry, Antibiotic, " & _
              "UserName, DateTimeOfRecord " & _
              "FROM CurrentAntibiotics50 " & _
              "WHERE SampleID = '" & SampleID & "'"
8480      Set tb = New Recordset
8490      RecOpenClient 0, tb, sql
8500      Do While Not tb.EOF
8510          Set Ix = New CurrentAntibiotic
8520          With Ix
8530              .SampleID = tb!SampleID & ""
8540              .Entry = tb!Entry
8550              .Antibiotic = tb!Antibiotic & ""
8560              .UserName = tb!UserName & ""
8570              .DateTimeOfRecord = tb!DateTimeOfRecord & ""
8580              mcolCurrentAntibiotics.Add Ix, CStr(Ix.Entry)
8590          End With
8600          tb.MoveNext
8610      Loop

8620      Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

8630      intEL = Erl
8640      strES = Err.Description
8650      LogError "CurrentAntibiotics", "Load", intEL, strES, sql

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
8660      Set NewEnum = mcolCurrentAntibiotics.[_NewEnum]

End Function


Public Function Item(ByVal Entry As String) As CurrentAntibiotic
Attribute Item.VB_UserMemId = 0
        
          Dim i As Integer

8670      On Error GoTo Item_Error

8680      Set Item = Nothing
8690      For i = 1 To mcolCurrentAntibiotics.Count
8700          If mcolCurrentAntibiotics.Item(i).Entry = Entry Then
8710              Set Item = mcolCurrentAntibiotics(i)
8720          End If
8730      Next i


8740      Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

8750      intEL = Erl
8760      strES = Err.Description
8770      LogError "CurrentAntibiotics", "Item", intEL, strES
          
End Function








Public Sub Add(ByVal c As CurrentAntibiotic)

8780      On Error Resume Next

8790      mcolCurrentAntibiotics.Add c, CStr(c.Entry)

End Sub


Public Sub Save()

          Dim sql As String
          Dim Ix As CurrentAntibiotic
          Dim UN As String
          Dim Sxs As New Sensitivities
          Dim sx As Sensitivity
          Dim NewSx As Sensitivity
          Dim Iso As Integer
          Dim SID As String
          Dim AB As ABDefinition

8800      On Error GoTo Save_Error

8810      SID = ""

8820      For Each Ix In mcolCurrentAntibiotics
8830          SID = Ix.SampleID
8840          UN = AddTicks(Ix.UserName)
8850          sql = "IF EXISTS (SELECT * FROM CurrentAntibiotics50 WHERE " & _
                  "           SampleID = '" & Ix.SampleID & "' " & _
                  "           AND Entry = '" & Ix.Entry & "') " & _
                  "  UPDATE CurrentAntibiotics50 " & _
                  "  SET Antibiotic = '" & Ix.Antibiotic & "', " & _
                  "  UserName = '" & UN & "' " & _
                  "  WHERE SampleID = '" & Ix.SampleID & "' " & _
                  "  AND Entry = '" & Ix.Entry & "' " & _
                  "ELSE " & _
                  "  INSERT INTO CurrentAntibiotics50 (SampleID, Entry, Antibiotic, UserName) " & _
                  "  VALUES " & _
                  "  ('" & Ix.SampleID & "', " & _
                  "  '" & Ix.Entry & "', " & _
                  "  '" & Ix.Antibiotic & "', " & _
                  "  '" & UN & "')"
8860          Cnxn(0).Execute sql

8870      Next

          'Ensure this is in Sensitivities
8880      If SID <> "" Then
8890          Sxs.Load SID
8900          For Iso = 1 To 4
8910              For Each sx In Sxs
8920                  If sx.IsolateNumber = Iso Then
8930                      For Each Ix In mcolCurrentAntibiotics
8940                          Set AB = New ABDefinition
8950                          Set AB = AB.Load(Ix.Antibiotic)
8960                          If Not AB Is Nothing Then
8970                              If AB.ViewInGrid Then
8980                                  Set NewSx = Sxs.Item(Iso, AntibioticCodeFor(Ix.Antibiotic))
8990                                  If NewSx Is Nothing Then
9000                                      Set NewSx = New Sensitivity
9010                                      NewSx.AntibioticCode = AntibioticCodeFor(Ix.Antibiotic)
9020                                      NewSx.AntibioticName = Ix.Antibiotic
9030                                      NewSx.Forced = 1
9040                                      NewSx.IsolateNumber = Iso
9050                                      NewSx.Report = 1
9060                                      NewSx.Rundate = Format$(Now, "dd/MMM/yyyy")
9070                                      NewSx.RunDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
9080                                      NewSx.SampleID = SID
9090                                      NewSx.UserCode = UserCode
9100                                      NewSx.Save
9110                                  Else
9120                                      Sxs.ForceUnForce NewSx, 1
9130                                  End If
9140                              End If
9150                          End If
9160                      Next
9170                  End If
9180              Next
9190          Next
9200      End If

9210      Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

9220      intEL = Erl
9230      strES = Err.Description
9240      LogError "CurrentAntibiotics", "Save", intEL, strES, sql

End Sub

Public Property Get CurrentAntibiotics() As CurrentAntibiotics
        
9250      Set CurrentAntibiotics = mCurrentAntibiotics

End Property

