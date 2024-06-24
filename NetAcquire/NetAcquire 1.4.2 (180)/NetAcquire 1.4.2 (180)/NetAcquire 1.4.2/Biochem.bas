Attribute VB_Name = "Biochem"
Option Explicit

Public fMainCounter As Long
Public fmainImageCounter As Integer
Public pbCounter As Long


Public Function AnalyserFor(ByVal Discipline As String, _
                            ByVal Code As String) As String

      Dim sql As String
      Dim tb As Recordset

54110 On Error GoTo AnalyserFor_Error

54120 AnalyserFor = ""
54130 sql = "SELECT Analyser " & _
            "FROM " & Discipline & "TestDefinitions " & _
            "WHERE Code = '" & Code & "'"
54140 Set tb = New Recordset
54150 RecOpenServer 0, tb, sql
54160 If Not tb.EOF Then
54170     AnalyserFor = tb!Analyser & ""
54180 End If

54190 Exit Function

AnalyserFor_Error:

      Dim strES As String
      Dim intEL As Integer

54200 intEL = Erl
54210 strES = Err.Description
54220 LogError "Biochem", "AnalyserFor", intEL, strES, sql

End Function


Public Function EntriesOK(ByVal SampleID As String, _
                          ByVal SurName As String, _
                          ByVal Sex As String, _
                          ByVal Ward As String, _
                          ByVal GP As String, _
                          ByVal Clin As String) _
                          As Boolean

54230 EntriesOK = False

54240 If Trim$(SampleID) = "" Then
54250     iMsg "Must have Lab Number.", vbCritical
54260     Exit Function
54270 End If

54280 If Trim$(Sex) = "" Then
54290     iMsg "Sex not entered." & vbCrLf & "Must have to enter sex.", vbCritical
54300         Exit Function
54310 End If

54320 If Trim$(SurName) <> "" Then
54330     If Trim$(Ward) = "" Then
54340         iMsg "Must have Ward entry.", vbCritical
54350         Exit Function
54360     End If

54370     If UCase$(Trim$(Ward)) = "GP" Then
54380         If Trim$(GP) = "" Then
54390             iMsg "Must have GP entry.", vbCritical
54400             Exit Function
54410         End If
54420         If Trim$(Clin) <> "" Then
54430             iMsg "Can't have Clinician entry if Ward is GP.", vbCritical
54440             Exit Function
54450         End If
              
54460     Else
              'Ward is provided

54470         If Trim$(GP) <> "" Then
54480             iMsg "Can't have GP entry if Ward is not GP.", vbCritical
54490             Exit Function
54500         End If
54510         If InStr(1, Ward, "nursing") = 0 And Trim$(Clin) = "" Then
54520             iMsg "Must have Clinician entry", vbCritical
54530             Exit Function
54540         End If
54550     End If
          
          
          
54560 End If




54570 EntriesOK = True

End Function

Public Function SetFormPrinter() As Boolean

54580 SetFormPrinter = True

End Function


Public Sub SetOptions()

54590 If sysOptDeptSemen(0) Then
54600     frmMain.mnuEditSemen.Visible = True
54610     frmMain.mnuCommentList(4).Visible = True
54620 Else
54630     frmMain.mnuEditSemen.Visible = False
54640     frmMain.mnuCommentList(4).Visible = False
54650 End If

54660 If sysOptDeptMicro(0) Then
54670     frmMain.mnuMicroReports.Visible = True
54680     frmMain.mnuEditMicrobiology.Visible = True
54690     frmMain.mnuDefaultsMicro.Visible = True
54700 Else
54710     frmMain.mnuMicroReports.Visible = False
54720     frmMain.mnuEditMicrobiology.Visible = False
54730     frmMain.mnuDefaultsMicro.Visible = False
54740 End If

End Sub


Public Function FlagMessage(ByVal strType As String, _
                            ByVal Historical As String, _
                            ByVal Current As String, _
                            Optional ByVal SampleID As String = "") _
                            As Boolean
      'Returns True to reject

      Dim s As String
      Dim RetVal As Boolean

54750 If Trim$(Historical) = "" Then Historical = "<Blank>"
54760 If Trim$(Current) = "" Then Current = "<Blank>"

54770 s = "Patients " & strType & " has changed!" & vbCrLf & _
          "Was '" & Historical & "'" & vbCrLf & _
          "Now '" & Current & "'" & vbCrLf & _
          "To accept this change, Press 'OK'"

54780 RetVal = iMsg(s, vbCritical + vbOKCancel, "Critical Warning") = vbCancel

54790 If Not RetVal Then
54800     StrEvent = Trim$(SampleID & " Name Change accepted. (") & Replace(s, vbCrLf, " ") & ")"
54810     LogEvent StrEvent, "Biochem", "FlagMessage"
54820 End If

54830 FlagMessage = RetVal

End Function

Public Sub LogEvent(ByVal e As String, _
                    ByVal ModuleName As String, _
                    ByVal ProcedureName As String)

      Dim tb As Recordset
      Dim sql As String
      Dim MyMachineName As String

54840 On Error GoTo LogEvent_Error

54850 MyMachineName = vbGetComputerName()

54860 sql = "Insert Into CustomEventLog " & _
            "( MSG, DateTime, ModuleName, ProcedureName,  " & _
            "  UserName, MachineName) VALUES " & _
            "('" & AddTicks(e) & "', " & _
            "'" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
            "'" & ModuleName & "', " & _
            "'" & ProcedureName & "', " & _
            "'" & AddTicks(UserName) & "', " & _
            "'" & AddTicks(MyMachineName) & "')"

54870 Set tb = New Recordset
54880 RecOpenClient 0, tb, sql

54890 Exit Sub

LogEvent_Error:

      Dim strES As String
      Dim intEL As Integer

54900 intEL = Erl
54910 strES = Err.Description
54920 LogError "Biochem", "LogEvent", intEL, strES, sql


End Sub

