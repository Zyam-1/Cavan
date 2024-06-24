VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DemogValidations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolDemogValidations As New Collection

Private mDemogValidations As New DemogValidations

Public Sub Add(ByVal DV As DemogValidation)

11320     On Error Resume Next

11330     If IsNumeric(DV.SampleID) Then
11340         mcolDemogValidations.Add DV, DV.SampleID
11350     End If

End Sub

Public Function Count() As Long
        
11360     Count = mcolDemogValidations.Count

End Function
Public Property Get DemogValidations() As DemogValidations
        
11370     Set DemogValidations = mDemogValidations

End Property

Public Function Item(ByVal SampleID As String) As DemogValidation
Attribute Item.VB_UserMemId = 0
        
          Dim D As DemogValidation
          Dim RetVal As DemogValidation

11380     Set RetVal = Nothing
11390     For Each D In mcolDemogValidations
11400         If D.SampleID = SampleID Then
11410             Set RetVal = D
11420             Exit For
11430         End If
11440     Next
11450     Set Item = RetVal
        
        
        
          'On Error Resume Next
          '
          'Set Item = mcolDemogValidations.Item(SampleID)

End Function
Public Sub LoadSingle(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String

11460     On Error GoTo LoadSingle_Error

11470     sql = "SELECT SampleID, EnteredBy, ValidatedBy, " & _
              "EnteredDateTime, ValidatedDateTime " & _
              "FROM DemogValidation " & _
              "WHERE SampleID = '" & SampleID & "'"
11480     Set tb = New Recordset
11490     RecOpenClient 0, tb, sql
11500     If Not tb.EOF Then
11510         Populate tb
11520     End If

11530     Exit Sub

LoadSingle_Error:

          Dim strES As String
          Dim intEL As Integer

11540     intEL = Erl
11550     strES = Err.Description
11560     LogError "DemogValidations", "LoadSingle", intEL, strES, sql

End Sub

Public Sub LoadByDate(ByVal DateOfInterest As String)

          Dim tb As Recordset
          Dim sql As String

11570     On Error GoTo LoadByDate_Error

11580     sql = "SELECT DV.SampleID, DV.EnteredBy, DV.ValidatedBy, " & _
              "D.DateTimeDemographics EnteredDateTime, DV.ValidatedDateTime " & _
              "FROM DemogValidation DV JOIN Demographics D " & _
              "ON DV.SampleID = D.SampleID " & _
              "WHERE D.RunDate = '" & Format$(DateOfInterest, "dd/MMM/yyyy") & "'"
11590     Set tb = New Recordset
11600     RecOpenClient 0, tb, sql
11610     Do While Not tb.EOF
11620         Populate tb
11630         tb.MoveNext
11640     Loop

11650     Exit Sub

LoadByDate_Error:

          Dim strES As String
          Dim intEL As Integer

11660     intEL = Erl
11670     strES = Err.Description
11680     LogError "DemogValidations", "LoadByDate", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
11690     Set NewEnum = mcolDemogValidations.[_NewEnum]

End Function

Private Sub Populate(ByVal tb As Recordset)

          Dim DV As New DemogValidation

11700     On Error GoTo Populate_Error

11710     With DV
11720         .SampleID = tb!SampleID & ""
11730         .EnteredBy = tb!EnteredBy & ""
11740         .ValidatedBy = tb!ValidatedBy & ""
11750         .EnteredDateTime = tb!EnteredDateTime
11760         .ValidatedDateTime = tb!ValidatedDateTime
11770         mcolDemogValidations.Add DV, DV.SampleID
11780     End With

11790     Exit Sub

Populate_Error:

          Dim strES As String
          Dim intEL As Integer

11800     intEL = Erl
11810     strES = Err.Description
11820     LogError "DemogValidations", "Populate", intEL, strES

End Sub

Public Sub Save(ByVal DV As DemogValidation)

          Dim sql As String

11830     On Error GoTo Save_Error

11840     sql = "IF NOT EXISTS(SELECT * FROM DemogValidation " & _
              "              WHERE SampleID = '" & DV.SampleID & "') " & _
              "  INSERT INTO DemogValidation " & _
              "  (SampleID, EnteredBy, ValidatedBy, EnteredDateTime) " & _
              "  VALUES " & _
              "  ('" & DV.SampleID & "', " & _
              "   '" & AddTicks(DV.EnteredBy) & "', " & _
              "   '" & AddTicks(DV.ValidatedBy) & "', " & _
              "   '" & Format$(DV.EnteredDateTime, "dd/MMM/yyyy HH:nn:ss") & "') "
11850     Cnxn(0).Execute sql

11860     sql = "UPDATE Demographics " & _
              "SET Valid = '1' " & _
              "WHERE SampleID = '" & DV.SampleID & "'"
11870     Cnxn(0).Execute sql

11880     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

11890     intEL = Erl
11900     strES = Err.Description
11910     LogError "DemogValidations", "Save", intEL, strES, sql

End Sub


Public Sub Delete(ByVal SampleID As String)

          Dim sql As String

11920     On Error GoTo Delete_Error

11930     sql = "DELETE FROM DemogValidation " & _
              "WHERE SampleID = '" & SampleID & "'"
11940     Cnxn(0).Execute sql

11950     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

11960     intEL = Erl
11970     strES = Err.Description
11980     LogError "DemogValidations", "Delete", intEL, strES, sql

End Sub


Public Sub SaveAll()

          Dim DV As DemogValidation

11990     On Error GoTo SaveAll_Error

12000     For Each DV In mcolDemogValidations
12010         Save DV
12020     Next

12030     Exit Sub

SaveAll_Error:

          Dim strES As String
          Dim intEL As Integer

12040     intEL = Erl
12050     strES = Err.Description
12060     LogError "DemogValidations", "SaveAll", intEL, strES

End Sub


