Attribute VB_Name = "modOptions"
Option Explicit

Public Type udtOptionList
    Description As String
    Value As String
    DefinedAs As String    'Boolean/String/Single/Long/Integer etc
End Type

Public sysOptAllowCopyDemographics() As Boolean
Public sysOptAllowGPFreeText() As Boolean
Public sysOptAllowWardFreeText() As Boolean
Public sysOptAlphaOrderTechnicians() As Boolean
Public sysOptAlwaysRequestFBC() As Boolean
Public sysOptAutoScrollAfterOrder() As Boolean

Public sysOptBioPhone() As String
Public sysOptBlankSID() As Boolean
Public sysOptBloodBank() As Boolean
Public sysOptBloodPhone() As String

Public sysOptCheckCholHDLRatio() As Boolean
Public sysOptCoagPhone() As String
Public sysOptCytoOffset() As Long    '40,000,000

Public sysOptDefaultABs() As Integer
Public sysOptDefaultTab() As String  'Default Tab
Public sysOptDeptBga() As Boolean
Public sysOptDeptBio() As Boolean
Public sysOptDeptCoag() As Boolean
Public sysOptDeptCyto() As Boolean
Public sysOptDeptEnd() As Boolean
Public sysOptDeptExt() As Boolean
Public sysOptDeptHaem() As Boolean
Public sysOptDeptHisto() As Boolean
Public sysOptDeptImm() As Boolean
Public sysOptDeptMicro() As Boolean
Public sysOptDeptSemen() As Boolean
Public sysOptDipStick() As Boolean
Public sysOptDoAssGlucose() As Boolean
Public sysOptDontPrintAllCoag() As Boolean
Public sysOptDontShowPrevCoag() As Boolean

Public sysOptExp() As Boolean
Public sysOptExtDefault() As Boolean

Public sysOptGpClin() As Boolean    'Allow Update of Gp/Clin/Ward

Public sysOptHaemAn1() As String
Public sysOptHaemAn2() As String
Public sysOptHaemPhone() As String
Public sysOptHistoOffset() As Long    '30,000,000
Public sysOptHospital() As Boolean  'Hospital Name

Public sysOptMicroOffset() As Long    '20,000,000
Public sysOptMicroOffsetOLD() As Long    '20,000,000

Public sysOptOrderComms() As Boolean

Public sysOptRemote() As Boolean

Public sysOptSemenOffset() As Long    '10,000,000
Public sysOptShortFaeces() As Boolean
Public sysOptShortUrine() As Boolean
Public sysOptSoundCritical() As String
Public sysOptSoundInformation() As String
Public sysOptSoundQuestion() As String
Public sysOptSoundSevere() As String

Public sysOptUrgent() As Boolean    'Urgent Request Look Up
Public sysOptUrgentRef() As Single  'Refresh rate of Urgent
Public sysOptUseFullID() As Boolean


Public Sub LoadOptions()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long

3170  On Error GoTo LoadOptions_Error

3180  ReDimOptions

      'For n = 0 To intOtherHospitalsInGroup
3190  n = 0
3200  sql = "Select * from Options " & _
            "order by ListOrder"

3210  sysOptUrgentRef(n) = 0.063

3220  Set tb = New Recordset
3230  RecOpenClient n, tb, sql
3240  Do While Not tb.EOF
3250      Select Case UCase$(Trim$(tb!Description & ""))
          Case "ALLOWCOPYDEMOGRAPHICS": sysOptAllowCopyDemographics(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3260      Case "ALLOWGPFREETEXT": sysOptAllowGPFreeText(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3270      Case "ALLOWWARDFREETEXT": sysOptAllowWardFreeText(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3280      Case "ALPHAORDERTECHNICIANS": sysOptAlphaOrderTechnicians(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3290      Case "ALWAYSREQUESTFBC": sysOptAlwaysRequestFBC(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3300      Case "AUTOSCROLLAFTERORDER": sysOptAutoScrollAfterOrder(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3310      Case "BIOPHONE": sysOptBioPhone(n) = Trim$(tb!Contents & "")
3320      Case "BLANKSID": sysOptBlankSID(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3330      Case "BLOODBANK": sysOptBloodBank(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3340      Case "BLOODPHONE": sysOptBloodPhone(n) = Trim$(tb!Contents & "")
3350      Case "CHECKCHOLHDLRATIO": sysOptCheckCholHDLRatio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3360      Case "COAGPHONE": sysOptCoagPhone(n) = Trim$(tb!Contents & "")
3370      Case "CYTOOFFSET": sysOptCytoOffset(n) = Val(tb!Contents & "")
3380      Case "DEFAULTABS": sysOptDefaultABs(n) = Val(tb!Contents & "")
3390      Case "DEFAULTTAB": sysOptDefaultTab(n) = Val(tb!Contents & "")
3400      Case "DEPTBGA": sysOptDeptBga(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3410      Case "DEPTBIO": sysOptDeptBio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3420      Case "DEPTCOAG": sysOptDeptCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3430      Case "DEPTCYTO": sysOptDeptCyto(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3440      Case "DEPTEND": sysOptDeptEnd(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3450      Case "DEPTEXT": sysOptDeptExt(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3460      Case "DEPTHAEM": sysOptDeptHaem(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3470      Case "DEPTHISTO": sysOptDeptHisto(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3480      Case "DEPTIMM": sysOptDeptImm(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3490      Case "DEPTMICRO": sysOptDeptMicro(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3500      Case "DEPTSEMEN": sysOptDeptSemen(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3510      Case "DIPSTICK": sysOptDipStick(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3520      Case "DOASSGLUCOSE": sysOptDoAssGlucose(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3530      Case "DONTPRINTALLCOAG": sysOptDontPrintAllCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3540      Case "DONTSHOWPREVCOAG": sysOptDontShowPrevCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3550      Case "EXP": sysOptExp(n) = Trim$(tb!Contents & "")
3560      Case "EXTDEFAULT": sysOptExtDefault(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3570      Case "GPCLIN": sysOptGpClin(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3580      Case "HAEMAN1": sysOptHaemAn1(n) = Trim$(tb!Contents & "")
3590      Case "HAEMAN2": sysOptHaemAn2(n) = Trim$(tb!Contents & "")
3600      Case "HAEMPHONE": sysOptHaemPhone(n) = Trim$(tb!Contents & "")
3610      Case "HISTOOFFSET": sysOptHistoOffset(n) = Trim$(tb!Contents & "")
3620      Case "HOSPITAL": sysOptHospital(n) = Trim$(tb!Contents & "")
3630      Case "MICROOFFSET": sysOptMicroOffset(n) = Val(Trim$(tb!Contents & ""))
3640      Case "MICROOFFSETOLD": sysOptMicroOffsetOLD(n) = Val(Trim$(tb!Contents & ""))
3650      Case "ORDERCOMMS": sysOptOrderComms(0) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3660      Case "REMOTE": sysOptRemote(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3670      Case "SEMENOFFSET": sysOptSemenOffset(n) = Val(Trim$(tb!Contents & ""))
3680      Case "SHORTFAECES": sysOptShortFaeces(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3690      Case "SHORTURINE": sysOptShortUrine(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3700      Case "SOUNDCRITICAL": sysOptSoundCritical(n) = Trim$(tb!Contents & "")
3710      Case "SOUNDINFORMATION": sysOptSoundInformation(n) = Trim$(tb!Contents & "")
3720      Case "SOUNDQUESTION": sysOptSoundQuestion(n) = Trim$(tb!Contents & "")
3730      Case "SOUNDSEVERE": sysOptSoundSevere(n) = Trim$(tb!Contents & "")
3740      Case "URGENT": sysOptUrgent(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
3750      Case "URGENTREF": sysOptUrgentRef(n) = Trim$(tb!Contents & "")
3760      Case "USEFULLID": sysOptUseFullID(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)

3770      End Select
3780      tb.MoveNext
3790  Loop
      'Next

3800  Exit Sub

3810  Exit Sub

LoadOptions_Error:

      Dim strES As String
      Dim intEL As Integer

3820  intEL = Erl
3830  strES = Err.Description
3840  LogError "modOptions", "LoadOptions", intEL, strES, sql

End Sub


Public Sub LoadFormOptions(ByRef Opts() As udtOptionList)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

3850  On Error GoTo LoadFormOptions_Error

3860  For n = 0 To UBound(Opts)
3870      sql = "Select * from Options where " & _
                "Description = '" & Opts(n).Description & "'"
3880      Set tb = New Recordset
3890      RecOpenClient 0, tb, sql
3900      If Not tb.EOF Then
3910          Opts(n).Value = Trim$(tb!Contents & "")
3920      End If
3930  Next

3940  Exit Sub

LoadFormOptions_Error:

      Dim strES As String
      Dim intEL As Integer

3950  intEL = Erl
3960  strES = Err.Description
3970  LogError "modOptions", "LoadFormOptions", intEL, strES, sql

End Sub

Private Sub ReDimOptions()

3980  ReDim sysOptAllowCopyDemographics(0 To intOtherHospitalsInGroup) As Boolean
3990  ReDim sysOptAllowGPFreeText(0 To intOtherHospitalsInGroup) As Boolean
4000  ReDim sysOptAllowWardFreeText(0 To intOtherHospitalsInGroup) As Boolean
4010  ReDim sysOptAlphaOrderTechnicians(0 To intOtherHospitalsInGroup)
4020  ReDim sysOptAlwaysRequestFBC(0 To intOtherHospitalsInGroup) As Boolean
4030  ReDim sysOptAutoScrollAfterOrder(0 To intOtherHospitalsInGroup) As Boolean

4040  ReDim sysOptBioPhone(0 To intOtherHospitalsInGroup) As String
4050  ReDim sysOptBlankSID(0 To intOtherHospitalsInGroup) As Boolean
4060  ReDim sysOptBloodBank(0 To intOtherHospitalsInGroup) As Boolean
4070  ReDim sysOptBloodPhone(0 To intOtherHospitalsInGroup) As String

4080  ReDim sysOptCheckCholHDLRatio(0 To intOtherHospitalsInGroup) As Boolean
4090  ReDim sysOptCoagPhone(0 To intOtherHospitalsInGroup) As String
4100  ReDim sysOptCytoOffset(0 To intOtherHospitalsInGroup) As Long    '40,000,000

4110  ReDim sysOptDefaultABs(0 To intOtherHospitalsInGroup) As Integer
4120  ReDim sysOptDefaultTab(0 To intOtherHospitalsInGroup) As String  'Default Tab
4130  ReDim sysOptDeptBga(0 To intOtherHospitalsInGroup) As Boolean
4140  ReDim sysOptDeptBio(0 To intOtherHospitalsInGroup) As Boolean
4150  ReDim sysOptDeptCoag(0 To intOtherHospitalsInGroup) As Boolean
4160  ReDim sysOptDeptCyto(0 To intOtherHospitalsInGroup) As Boolean
4170  ReDim sysOptDeptEnd(0 To intOtherHospitalsInGroup) As Boolean
4180  ReDim sysOptDeptExt(0 To intOtherHospitalsInGroup) As Boolean
4190  ReDim sysOptDeptHaem(0 To intOtherHospitalsInGroup) As Boolean
4200  ReDim sysOptDeptHisto(0 To intOtherHospitalsInGroup) As Boolean
4210  ReDim sysOptDeptImm(0 To intOtherHospitalsInGroup) As Boolean
4220  ReDim sysOptDeptMicro(0 To intOtherHospitalsInGroup) As Boolean
4230  ReDim sysOptDeptSemen(0 To intOtherHospitalsInGroup) As Boolean
4240  ReDim sysOptDipStick(0 To intOtherHospitalsInGroup) As Boolean
4250  ReDim sysOptDoAssGlucose(0 To intOtherHospitalsInGroup) As Boolean
4260  ReDim sysOptDontPrintAllCoag(0 To intOtherHospitalsInGroup) As Boolean
4270  ReDim sysOptDontShowPrevCoag(0 To intOtherHospitalsInGroup) As Boolean

4280  ReDim sysOptExp(0 To intOtherHospitalsInGroup) As Boolean
4290  ReDim sysOptExtDefault(0 To intOtherHospitalsInGroup) As Boolean

4300  ReDim sysOptGpClin(0 To intOtherHospitalsInGroup) As Boolean    'Allow Update of Gp/Clin/Ward

4310  ReDim sysOptHaemAn1(0 To intOtherHospitalsInGroup) As String
4320  ReDim sysOptHaemAn2(0 To intOtherHospitalsInGroup) As String
4330  ReDim sysOptHaemPhone(0 To intOtherHospitalsInGroup) As String
4340  ReDim sysOptHistoOffset(0 To intOtherHospitalsInGroup) As Long    '30,000,000
4350  ReDim sysOptHospital(0 To intOtherHospitalsInGroup) As Boolean  'Hospital Name

4360  ReDim sysOptMicroOffset(0 To intOtherHospitalsInGroup) As Long    '20,000,000
4370  ReDim sysOptMicroOffsetOLD(0 To intOtherHospitalsInGroup) As Long    '20,000,000
4380  ReDim sysOptOrderComms(0 To intOtherHospitalsInGroup) As Boolean

4390  ReDim sysOptRemote(0 To intOtherHospitalsInGroup) As Boolean

4400  ReDim sysOptSemenOffset(0 To intOtherHospitalsInGroup) As Long    '10,000,000
4410  ReDim sysOptShortFaeces(0 To intOtherHospitalsInGroup) As Boolean
4420  ReDim sysOptShortUrine(0 To intOtherHospitalsInGroup) As Boolean
4430  ReDim sysOptSoundCritical(0 To intOtherHospitalsInGroup) As String
4440  ReDim sysOptSoundInformation(0 To intOtherHospitalsInGroup) As String
4450  ReDim sysOptSoundQuestion(0 To intOtherHospitalsInGroup) As String
4460  ReDim sysOptSoundSevere(0 To intOtherHospitalsInGroup) As String

4470  ReDim sysOptUrgent(0 To intOtherHospitalsInGroup) As Boolean
4480  ReDim sysOptUrgentRef(0 To intOtherHospitalsInGroup) As Single
4490  ReDim sysOptUseFullID(0 To intOtherHospitalsInGroup) As Boolean
End Sub

Public Function CheckDemo(p_SampleID As String) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

4500  On Error GoTo CheckDemo_Error

4510  CheckDemo = False
4520  sql = "Select SampleID from Demographics Where SampleID = '" & p_SampleID & "'"
4530  Set tb = New Recordset
4540  RecOpenClient 0, tb, sql
4550  If Not tb Is Nothing Then
4560      If Not tb.EOF Then
4570           If p_SampleID = ConvertNull(tb!SampleID, "") Then
4580              CheckDemo = True
4590           End If
4600      End If
4610  End If

4620  Exit Function

CheckDemo_Error:

      Dim strES As String
      Dim intEL As Integer

4630  intEL = Erl
4640  strES = Err.Description
4650  LogError "modOptions", "CheckDemo", intEL, strES, sql

End Function
