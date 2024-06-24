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
  
10    Count = mcolCurrentAntibiotics.Count

End Function

Public Sub Delete(ByVal SampleID As String, ByVal AB As String)
  
      Dim sql As String
      Dim CURS As New CurrentAntibiotics

10    On Error GoTo Delete_Error

20    CURS.Load SampleID

30    sql = "DELETE FROM CurrentAntibiotics50 WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Antibiotic = '" & AB & "'"
40    Cnxn(0).Execute sql

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "CurrentAntibiotics", "Delete", intEL, strES, sql

End Sub
Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Ix As CurrentAntibiotic

10    On Error GoTo Load_Error

20    sql = "SELECT SampleID, COALESCE(Entry, 0) Entry, Antibiotic, " & _
            "UserName, DateTimeOfRecord " & _
            "FROM CurrentAntibiotics50 " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Ix = New CurrentAntibiotic
70      With Ix
80        .SampleID = tb!SampleID & ""
90        .Entry = tb!Entry
100       .Antibiotic = tb!Antibiotic & ""
110       .UserName = tb!UserName & ""
120       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
130       mcolCurrentAntibiotics.Add Ix, CStr(Ix.Entry)
140     End With
150     tb.MoveNext
160   Loop

170   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "CurrentAntibiotics", "Load", intEL, strES, sql

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolCurrentAntibiotics.[_NewEnum]

End Function


Public Function Item(ByVal Entry As String) As CurrentAntibiotic
Attribute Item.VB_UserMemId = 0
        
      Dim I As Integer

10    On Error GoTo Item_Error

20    Set Item = Nothing
30    For I = 1 To mcolCurrentAntibiotics.Count
40        If mcolCurrentAntibiotics.Item(I).Entry = Entry Then
50            Set Item = mcolCurrentAntibiotics(I)
60        End If
70    Next I


80    Exit Function

Item_Error:

       Dim strES As String
       Dim intEL As Integer

90     intEL = Erl
100    strES = Err.Description
110    LogError "CurrentAntibiotics", "Item", intEL, strES
          
End Function








Public Sub Add(ByVal c As CurrentAntibiotic)

10    On Error Resume Next

20    mcolCurrentAntibiotics.Add c, CStr(c.Entry)

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

10    On Error GoTo Save_Error

20    SID = ""

30    For Each Ix In mcolCurrentAntibiotics
40      SID = Ix.SampleID
50      UN = AddTicks(Ix.UserName)
60      sql = "IF EXISTS (SELECT * FROM CurrentAntibiotics50 WHERE " & _
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
70      Cnxn(0).Execute sql

80    Next

      'Ensure this is in Sensitivities
90    If SID <> "" Then
100     Sxs.Load SID
110     For Iso = 1 To 4
120       For Each sx In Sxs
130         If sx.IsolateNumber = Iso Then
140           For Each Ix In mcolCurrentAntibiotics
150             Set AB = New ABDefinition
160             Set AB = AB.Load(Ix.Antibiotic)
170             If Not AB Is Nothing Then
180               If AB.ViewInGrid Then
190                 Set NewSx = Sxs.Item(Iso, AntibioticCodeFor(Ix.Antibiotic))
200                 If NewSx Is Nothing Then
210                   Set NewSx = New Sensitivity
220                   NewSx.AntibioticCode = AntibioticCodeFor(Ix.Antibiotic)
230                   NewSx.AntibioticName = Ix.Antibiotic
240                   NewSx.Forced = 1
250                   NewSx.IsolateNumber = Iso
260                   NewSx.Report = 1
270                   NewSx.Rundate = Format$(Now, "dd/MMM/yyyy")
280                   NewSx.RunDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
290                   NewSx.SampleID = SID
300                   NewSx.UserCode = UserCode
310                   NewSx.Save
320                 Else
330                   Sxs.ForceUnForce NewSx, 1
340                 End If
350               End If
360             End If
370           Next
380         End If
390       Next
400     Next
410   End If

420   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "CurrentAntibiotics", "Save", intEL, strES, sql

End Sub

Public Property Get CurrentAntibiotics() As CurrentAntibiotics
  
10    Set CurrentAntibiotics = mCurrentAntibiotics

End Property

