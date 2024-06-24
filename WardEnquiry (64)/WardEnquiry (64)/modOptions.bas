Attribute VB_Name = "modOptions"
Option Explicit

Public Type udtOptionList
  Description As String
  Value As String
  DefinedAs As String 'Boolean/String/Single/Long/Integer etc
End Type

Public sysOptBioCodeForChol() As String
Public sysOptBioCodeForCholHDLRatio() As String
Public sysOptBioCodeForGlucose() As String
Public sysOptBioCodeForLDL() As String
Public sysOptBioCodeForTrig() As String
Public sysOptBioCodeForGentamicin() As String
Public sysOptBioCodeForTobramicin() As String

Public sysOptBioPhone() As String

Public sysOptCheckCholHDLRatio() As Boolean

Public sysOptDefaultABs() As Integer
Public sysOptDeptBga() As Boolean
Public sysOptDeptBio() As Boolean
Public sysOptDeptCoag() As Boolean
Public sysOptDeptEnd() As Boolean
Public sysOptDeptExt() As Boolean
Public sysOptDeptHaem() As Boolean
Public sysOptDeptImm() As Boolean
Public sysOptDeptMicro() As Boolean
Public sysOptDeptSemen() As Boolean
Public sysOptDisableWardOrdering() As Boolean
Public sysOptDisableWardPrinting() As Boolean
Public sysOptDoAssGlucose() As Boolean

Public sysOptHaemAn1() As String
Public sysOptHaemPhone() As String
Public sysOptHighBack() As Long
Public sysOptHighFore() As Long

Public sysOptLowBack() As Long
Public sysOptLowFore() As Long  'Text Color of Low

Public sysOptMicroOffset() As Long '0
Public sysOptMicroOffsetOLD() As Long    '20,000,000
Public sysOptPOCTOffset() As Long    '30,000,000

Public sysOptOrderComms() As Boolean

Public sysOptPlasBack() As String
Public sysOptPlasFore() As String

Public sysOptRemote() As Boolean

Public sysOptSemenOffset() As Long '10,000,000
Public sysOptSoundCritical() As String
Public sysOptSoundInformation() As String
Public sysOptSoundQuestion() As String
Public sysOptSoundSevere() As String

Public sysOptWardSearchDoB() As Boolean
Public sysOptWardSearchName() As Boolean
Public sysOptViewUnsignedSamples() As Boolean
Public sysOptWBCDC() As Boolean  '
Public sysOptWardChartLocation() As Boolean


Public Sub LoadOptions()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long

10    On Error GoTo ehlc

20    ReDimOptions

30    n = 0
40    sql = "Select * from Options " & _
            "order by ListOrder"

50    sysOptLowBack(n) = &HFFFF80
60    sysOptLowFore(n) = &HC00000
70    sysOptHighBack(n) = &HFFFF&
80    sysOptHighFore(n) = &HFF&

90    Set tb = New Recordset
100   RecOpenClient n, tb, sql
110   Do While Not tb.EOF
120   If tb!Description & "" = "ViewUnsignedSamples" Then
130   Debug.Print tb!Description & ""
140   End If
150       Select Case UCase$(Trim$(tb!Description & ""))

          Case "BIOCODEFORCHOL": sysOptBioCodeForChol(n) = Trim$(tb!Contents & "")
160       Case "BIOCODEFORCHOLHDLRATIO": sysOptBioCodeForCholHDLRatio(n) = Trim$(tb!Contents & "")
170       Case "BIOCODEFORGLUCOSE": sysOptBioCodeForGlucose(n) = Trim$(tb!Contents & "")
180       Case "BIOCODEFORTRIG": sysOptBioCodeForTrig(n) = Trim$(tb!Contents & "")
190       Case "BIOCODEFORGENTAMICIN": sysOptBioCodeForGentamicin(n) = Trim$(tb!Contents & "")
200       Case "BIOCODEFORTOBRAMICIN": sysOptBioCodeForTobramicin(n) = Trim$(tb!Contents & "")
210       Case "BIOPHONE": sysOptBioPhone(n) = Trim$(tb!Contents & "")
220       Case "CHECKCHOLHDLRATIO": sysOptCheckCholHDLRatio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
230       Case "DEFAULTABS": sysOptDefaultABs(n) = Val(tb!Contents & "")
240       Case "DEPTBGA": sysOptDeptBga(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
250       Case "DEPTBIO": sysOptDeptBio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
260       Case "DEPTCOAG": sysOptDeptCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
270       Case "DEPTEND": sysOptDeptEnd(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
280       Case "DEPTEXT": sysOptDeptExt(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
290       Case "DEPTHAEM": sysOptDeptHaem(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
300       Case "DEPTIMM": sysOptDeptImm(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
310       Case "DEPTMICRO": sysOptDeptMicro(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
320       Case "DEPTSEMEN": sysOptDeptSemen(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
330       Case "DISABLEWARDORDERING": sysOptDisableWardOrdering(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
340       Case "DISABLEWARDPRINTING": sysOptDisableWardPrinting(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
350       Case "DOASSGLUCOSE": sysOptDoAssGlucose(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
360       Case "HAEMAN1": sysOptHaemAn1(n) = Trim$(tb!Contents & "")
370       Case "HAEMPHONE": sysOptHaemPhone(n) = Trim$(tb!Contents & "")
380       Case "HIGHBACK": sysOptHighBack(n) = Val(Trim$(tb!Contents & ""))
390       Case "HIGHFORE": sysOptHighFore(n) = Val(Trim$(tb!Contents & ""))
400       Case "LOWBACK": sysOptLowBack(n) = Val(Trim$(tb!Contents & ""))
410       Case "LOWFORE": sysOptLowFore(n) = Val(Trim$(tb!Contents & ""))
420       Case "MICROOFFSET": sysOptMicroOffset(n) = Val(Trim$(tb!Contents & ""))
430       Case "MICROOFFSETOLD": sysOptMicroOffsetOLD(n) = Val(Trim$(tb!Contents & ""))
440       Case "POCTOFFSET": sysOptPOCTOffset(n) = Val(Trim$(tb!Contents & ""))
450       Case "ORDERCOMMS": sysOptOrderComms(0) = IIf(Trim$(tb!Contents & "") = "1", True, False)
460       Case "PLASBACK": sysOptPlasBack(n) = Trim$(tb!Contents & "")
470       Case "PLASFORE": sysOptPlasFore(n) = Trim$(tb!Contents & "")
480       Case "REMOTE": sysOptRemote(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
490       Case "SEMENOFFSET": sysOptSemenOffset(n) = Val(Trim$(tb!Contents & ""))
500       Case "SOUNDCRITICAL": sysOptSoundCritical(n) = Trim$(tb!Contents & "")
510       Case "SOUNDINFORMATION": sysOptSoundInformation(n) = Trim$(tb!Contents & "")
520       Case "SOUNDQUESTION": sysOptSoundQuestion(n) = Trim$(tb!Contents & "")
530       Case "SOUNDSEVERE": sysOptSoundSevere(n) = Trim$(tb!Contents & "")
540       Case "WARDSEARCHDOB": sysOptWardSearchDoB(n) = Val(Trim$(tb!Contents & ""))
550       Case UCase("ViewUnsignedSamples"): sysOptViewUnsignedSamples(n) = Val(Trim$(tb!Contents & ""))
560       Case "WARDSEARCHNAME": sysOptWardSearchName(n) = Val(Trim$(tb!Contents & ""))
570       Case "WARDCHARTLOCATION": sysOptWardChartLocation(n) = Trim$(tb!Contents & "")
580       Case "WBCDC": sysOptWBCDC(n) = Trim$(tb!Contents & "")

590       End Select
600       tb.MoveNext
610   Loop
      'Next

620   Exit Sub

ehlc:
      Dim er As Long
      Dim ers As String

630   er = Err.Number
640   ers = Err.Description

650   Screen.MousePointer = 0
      'LogError "modOptions/LoadOptions:" & Str(er) & ":" & ers
660   Exit Sub

End Sub


Public Sub LoadFormOptions(ByRef Opts() As udtOptionList)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo LoadFormOptions_Error

20    For n = 0 To UBound(Opts)
30      sql = "Select * from Options where " & _
              "Description = '" & Opts(n).Description & "'"
40      Set tb = New Recordset
50      RecOpenClient 0, tb, sql
60      If Not tb.EOF Then
70        Opts(n).Value = Trim$(tb!Contents & "")
80      End If
90    Next

100   Exit Sub

LoadFormOptions_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modOptions", "LoadFormOptions", intEL, strES, sql


End Sub

Private Sub ReDimOptions()

10    ReDim sysOptBioCodeForChol(0 To intOtherHospitalsInGroup) As String
20    ReDim sysOptBioCodeForCholHDLRatio(0 To intOtherHospitalsInGroup) As String
30    ReDim sysOptBioCodeForGlucose(0 To intOtherHospitalsInGroup) As String
40    ReDim sysOptBioCodeForTrig(0 To intOtherHospitalsInGroup) As String
50    ReDim sysOptBioCodeForGentamicin(0 To intOtherHospitalsInGroup) As String
60    ReDim sysOptBioCodeForTobramicin(0 To intOtherHospitalsInGroup) As String

70    ReDim sysOptBioPhone(0 To intOtherHospitalsInGroup) As String

80    ReDim sysOptCheckCholHDLRatio(0 To intOtherHospitalsInGroup) As Boolean

90    ReDim sysOptDefaultABs(0 To intOtherHospitalsInGroup) As Integer
100   ReDim sysOptDeptBga(0 To intOtherHospitalsInGroup) As Boolean
110   ReDim sysOptDeptBio(0 To intOtherHospitalsInGroup) As Boolean
120   ReDim sysOptDeptCoag(0 To intOtherHospitalsInGroup) As Boolean
130   ReDim sysOptDeptEnd(0 To intOtherHospitalsInGroup) As Boolean
140   ReDim sysOptDeptExt(0 To intOtherHospitalsInGroup) As Boolean
150   ReDim sysOptDeptHaem(0 To intOtherHospitalsInGroup) As Boolean
160   ReDim sysOptDeptImm(0 To intOtherHospitalsInGroup) As Boolean
170   ReDim sysOptDeptMicro(0 To intOtherHospitalsInGroup) As Boolean
180   ReDim sysOptDeptSemen(0 To intOtherHospitalsInGroup) As Boolean
190   ReDim sysOptDisableWardOrdering(0 To intOtherHospitalsInGroup) As Boolean
200   ReDim sysOptDisableWardPrinting(0 To intOtherHospitalsInGroup)
210   ReDim sysOptDoAssGlucose(0 To intOtherHospitalsInGroup) As Boolean

220   ReDim sysOptHaemAn1(0 To intOtherHospitalsInGroup) As String
230   ReDim sysOptHaemPhone(0 To intOtherHospitalsInGroup) As String
240   ReDim sysOptHighBack(0 To intOtherHospitalsInGroup) As Long
250   ReDim sysOptHighFore(0 To intOtherHospitalsInGroup) As Long

260   ReDim sysOptLowBack(0 To intOtherHospitalsInGroup) As Long
270   ReDim sysOptLowFore(0 To intOtherHospitalsInGroup) As Long  'Text Color of Low

280   ReDim sysOptMicroOffset(0 To intOtherHospitalsInGroup) As Long '20,000,000
290   ReDim sysOptMicroOffsetOLD(0 To intOtherHospitalsInGroup) As Long    '20,000,000
300   ReDim sysOptPOCTOffset(0 To intOtherHospitalsInGroup) As Long '30,000,000

310   ReDim sysOptOrderComms(0 To intOtherHospitalsInGroup) As Boolean

320   ReDim sysOptPlasBack(0 To intOtherHospitalsInGroup) As String
330   ReDim sysOptPlasFore(0 To intOtherHospitalsInGroup) As String

340   ReDim sysOptRemote(0 To intOtherHospitalsInGroup) As Boolean

350   ReDim sysOptSemenOffset(0 To intOtherHospitalsInGroup) As Long '10,000,000
360   ReDim sysOptSoundCritical(0 To intOtherHospitalsInGroup) As String
370   ReDim sysOptSoundInformation(0 To intOtherHospitalsInGroup) As String
380   ReDim sysOptSoundQuestion(0 To intOtherHospitalsInGroup) As String
390   ReDim sysOptSoundSevere(0 To intOtherHospitalsInGroup) As String

400   ReDim sysOptWardSearchDoB(0 To intOtherHospitalsInGroup)
410   ReDim sysOptViewUnsignedSamples(0 To intOtherHospitalsInGroup)
420   ReDim sysOptWardSearchName(0 To intOtherHospitalsInGroup)
430   ReDim sysOptWBCDC(0 To intOtherHospitalsInGroup) As Boolean  '
440   ReDim sysOptWardChartLocation(0 To intOtherHospitalsInGroup) As Boolean  '

End Sub


Public Function GetOptionSetting(ByVal Description As String, _
                                 ByVal Default As String, _
                                 ByVal UserName As String) As String
   
      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetOptionSetting_Error

20    sql = "SELECT Contents FROM Options WHERE " & _
            "Description = '" & Description & "' " & _
            "AND COALESCE(Username, '') = '" & AddTicks(UserName) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      RetVal = Default
70    ElseIf Trim$(tb!Contents & "") = "" Then
80      RetVal = Default
90    Else
100     RetVal = tb!Contents
110   End If

120   GetOptionSetting = RetVal

130   Exit Function

GetOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "basOptions", "GetOptionSetting", intEL, strES, sql

End Function

Public Sub SaveOptionSetting(ByVal Description As String, _
                             ByVal Contents As String, _
                             ByVal UserName As String)
   
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveOptionSetting_Error

20    UserName = AddTicks(UserName)
30    Contents = AddTicks(Contents)

40    sql = "IF EXISTS (SELECT * FROM Options WHERE " & _
            "           Description = '" & Description & "' AND " & _
            "           Username = '" & UserName & "') " & _
            "  UPDATE Options SET Contents = '" & Contents & "' " & _
            "  WHERE Description = '" & Description & "' AND " & _
            "  Username = '" & UserName & "' " & _
            "ELSE " & _
            "  INSERT INTO Options " & _
            "  (Description, Contents, UserName) VALUES ( " & _
            "  '" & Description & "', " & _
            "  '" & Contents & "', " & _
            "  '" & UserName & "')"
50    Cnxn(0).Execute sql

60    Exit Sub

SaveOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "basOptions", "SaveOptionSetting", intEL, strES, sql

End Sub

