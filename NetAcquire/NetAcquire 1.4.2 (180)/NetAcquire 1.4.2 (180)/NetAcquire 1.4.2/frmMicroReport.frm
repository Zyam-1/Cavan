VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMicroReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11040
   ControlBox      =   0   'False
   HelpContextID   =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   11040
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   705
      Left            =   9630
      Picture         =   "frmMicroReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Log as Phoned"
      Top             =   2250
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   13080
      TabIndex        =   22
      Top             =   1260
      Width           =   1245
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View/Edit"
      Height          =   615
      Left            =   9630
      Picture         =   "frmMicroReport.frx":0351
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1590
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   585
      Left            =   9630
      Picture         =   "frmMicroReport.frx":0C1B
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   300
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   9630
      Picture         =   "frmMicroReport.frx":0F25
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "bprint"
      Top             =   930
      Width           =   1275
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   5835
      Left            =   150
      TabIndex        =   18
      Top             =   3690
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   10292
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMicroReport.frx":158F
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   1935
      Left            =   150
      TabIndex        =   16
      Top             =   1680
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   7
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmMicroReport.frx":1611
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9180
      Top             =   -60
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   630
      Left            =   9630
      Picture         =   "frmMicroReport.frx":16CD
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   9345
      Begin VB.OptionButton optNew 
         Caption         =   "Ward View"
         Height          =   285
         Left            =   8055
         TabIndex        =   25
         Top             =   1035
         Width           =   1185
      End
      Begin VB.OptionButton optOld 
         Caption         =   "Lab view"
         Height          =   285
         Left            =   8055
         TabIndex        =   24
         Top             =   810
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Label lblClDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1080
         Width           =   7785
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1065
         TabIndex        =   13
         Top             =   210
         Width           =   3540
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   12
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6045
         TabIndex        =   11
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4515
         TabIndex        =   10
         Top             =   510
         Width           =   4140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   615
         TabIndex        =   9
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   8
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   5130
         TabIndex        =   7
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   3915
         TabIndex        =   6
         Top             =   540
         Width           =   570
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7935
         TabIndex        =   5
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2775
         TabIndex        =   4
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   7620
         TabIndex        =   3
         Top             =   210
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   2445
         TabIndex        =   2
         Top             =   540
         Width           =   270
      End
      Begin VB.Label lblDemogComment 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   810
         Width           =   7785
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   150
      TabIndex        =   15
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmMicroReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pReturnedSampleID As String

Private Activated As Boolean

Private Type OrgGroup
    OrgGroup As String
    OrgName As String
    ShortName As String
    ReportName As String
    Qualifier As String
End Type

Dim WordResultPrinted As Boolean
Private SelectedIndex As Integer

Private pPrintToPrinter As String

Private pFromEdit As Boolean

Private Function AnyNotForReporting(ByVal Sxs As Sensitivities, ByVal IsolateNumber As Integer) As Boolean

      Dim RetVal As Boolean
      Dim sx As Sensitivity

43600 RetVal = False

43610 For Each sx In Sxs
43620   If sx.IsolateNumber = IsolateNumber And sx.Report = False And Trim$(sx.RSI) <> "" Then
43630     RetVal = True
43640     Exit For
43650   End If
43660 Next

43670 AnyNotForReporting = RetVal

End Function

Private Sub FillIsolates(ByVal SampleID As String)

      Dim Ix As Isolate
      Dim Ixs As New Isolates
      Dim n As Integer

43680 On Error GoTo FillIsolates_Error

43690 Ixs.Load Val(SampleID) ' + sysOptMicroOffset(0)
43700 If Ixs.Count = 0 Then Exit Sub

43710 For n = 1 To 4
43720   Set Ix = Ixs(n)
43730   If Not Ix Is Nothing Then
43740     rtb.SelFontName = "Courier New"
43750     rtb.SelColor = vbBlack
43760     rtb.SelFontSize = 10
43770     rtb.SelBold = False
43780     rtb.SelText = Space$(5) & Ix.OrganismName & " " & Ix.Qualifier & vbCrLf
43790   End If
43800 Next

43810 Exit Sub

FillIsolates_Error:

      Dim strES As String
      Dim intEL As Integer

43820 intEL = Erl
43830 strES = Err.Description
43840 LogError "frmMicroReport", "FillIsolates", intEL, strES

End Sub

Public Property Get PrintToPrinter() As String

43850 PrintToPrinter = pPrintToPrinter

End Property


Public Property Let FromEdit(ByVal blnNewValue As Boolean)

43860 pFromEdit = blnNewValue

End Property

Public Property Let PrintToPrinter(ByVal strNewValue As String)

43870 pPrintToPrinter = strNewValue

End Property


Private Sub FillCommentsRTB(ByVal SampleID As String, ByVal Cn As Integer)

      Dim OB As Observation
      Dim OBs As Observations
      Dim s As String
      Dim sql As String
      Dim CommentsFound As Boolean
      Dim DemogComment As String
      Dim MSCComment As String
      Dim ConsComment As String

43880 On Error GoTo FillCommentsRTB_Error

43890 CommentsFound = False
43900 Set OBs = New Observations
43910 Set OBs = OBs.Load(SampleID + sysOptMicroOffset(Cn), "Demographic", "MicroConsultant", "MicroCS")
43920 If Not OBs Is Nothing Then
43930   For Each OB In OBs
43940     Select Case UCase$(OB.Discipline)
            Case "DEMOGRAPHIC": DemogComment = OB.Comment
43950       Case "MICROCONSULTANT": ConsComment = OB.Comment
43960       Case "MICROCS": MSCComment = OB.Comment
43970     End Select
43980   Next
43990   CommentsFound = True
44000   rtb.SelColor = vbBlue
44010   rtb.SelFontSize = 12
44020   rtb.SelBold = True
44030   rtb.SelText = "Comments:" & vbCrLf
44040   rtb.SelColor = vbBlack
44050   rtb.SelFontSize = 10
44060   If Trim$(DemogComment) <> "" Then
44070     rtb.SelBold = True
44080     rtb.SelText = "Demographic Comment: "
44090     rtb.SelBold = False
44100     rtb.SelText = DemogComment & vbCrLf
44110   End If
44120   If Trim$(ConsComment) <> "" Then
44130     rtb.SelBold = True
44140     rtb.SelText = "Consultant Comment: "
44150     rtb.SelBold = False
44160     rtb.SelText = ConsComment & vbCrLf
44170   End If
44180   If Trim$(MSCComment) <> "" Then
44190     rtb.SelBold = True
44200     rtb.SelText = "Medical Scientist Comment: "
44210     rtb.SelBold = False
44220     rtb.SelText = MSCComment & vbCrLf
44230   End If
44240 End If

      Dim CURS As New CurrentAntibiotics
      Dim Cur As CurrentAntibiotic
44250 CURS.Load Val(SampleID) + sysOptMicroOffset(Cn)

44260 s = ""
44270 For Each Cur In CURS
44280   s = Cur.Antibiotic & " "
44290 Next
44300 If Trim$(s) <> "" Then
44310   If Not CommentsFound Then
44320     rtb.SelColor = vbBlue
44330     rtb.SelFontSize = 12
44340     rtb.SelBold = True
44350     rtb.SelText = "Comments:" & vbCrLf
44360     rtb.SelColor = vbBlack
44370     rtb.SelFontSize = 10
44380   End If
44390   rtb.SelBold = True
44400   rtb.SelText = "Current Antibiotics: "
44410   rtb.SelBold = False
44420   rtb.SelText = s & vbCrLf
44430   CommentsFound = True
44440 End If

44450 If CommentsFound Then
44460     rtb.SelText = vbCrLf
44470     rtb.SelText = String(80, "-") & vbCrLf
44480 End If

44490 Exit Sub

FillCommentsRTB_Error:

      Dim strES As String
      Dim intEL As Integer

44500 intEL = Erl
44510 strES = Err.Description
44520 LogError "frmMicroReport", "FillCommentsRTB", intEL, strES, sql

End Sub
Private Sub FillGenericResultsRTB(ByVal SampleID As String)

      Dim TestName As String
      Dim Gx As GenericResult
      Dim GXs As New GenericResults

44530 On Error GoTo FillGenericResultsRTB_Error

44540 GXs.Load Val(SampleID) ' + sysOptMicroOffset(0)
44550 If GXs.Count > 0 Then
44560   rtb.SelText = vbCrLf
44570   If Not WordResultPrinted Then
44580     rtb.SelColor = vbBlue
44590     rtb.SelBold = True
44600     rtb.SelFontSize = 12
44610     rtb.SelText = "Results:" & vbCrLf
44620     WordResultPrinted = True
44630   End If
44640   rtb.SelColor = vbBlack
44650   rtb.SelBold = False
44660   rtb.SelFontSize = 10
44670   For Each Gx In GXs
44680     rtb.SelColor = vbBlack
44690     rtb.SelBold = False
44700     rtb.SelFontSize = 10
44710     TestName = Gx.TestName
44720     If UCase$(TestName) = "REDSUB" Then
44730       TestName = "Reducing Substances"
44740     End If
44750     rtb.SelText = TestName & " : "
44760     rtb.SelBold = True
44770     rtb.SelText = IIf(Gx.Valid <> 0, Gx.Result, "Not yet available") & vbCrLf
44780   Next
44790   rtb.SelText = vbCrLf
44800 End If

44810 Exit Sub

FillGenericResultsRTB_Error:

        Dim strES As String
        Dim intEL As Integer

44820   intEL = Erl
44830   strES = Err.Description
44840   LogError "frmMicroReport", "FillGenericResultsRTB", intEL, strES

End Sub

Private Sub FillFaecesRTB(ByVal SampleID As String, _
                          ByVal Cn As Integer)

      Dim n As Integer
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

44850 On Error GoTo FillFaecesRTB_Error

44860 Fxs.Load Val(SampleID) + sysOptMicroOffset(Cn)

44870 If Fxs.Count > 0 Then
44880   rtb.SelText = vbCrLf
44890   If Not WordResultPrinted Then
44900       rtb.SelColor = vbBlue
44910       rtb.SelBold = True
44920       rtb.SelFontSize = 12
44930       rtb.SelText = "Results:" & vbCrLf
44940       WordResultPrinted = True
44950   End If
44960   rtb.SelColor = vbBlack

        'Occult Blood
44970   For n = 0 To 2
44980     Set Fx = Fxs.Item("OB" & Format$(n))
44990     If Not Fx Is Nothing Then
45000       rtb.SelBold = False
45010       rtb.SelText = "Occult Blood (" & n & ") : "
45020       rtb.SelBold = True
45030       If Fx.Valid = 0 Then
45040         rtb.SelText = "Not yet available." & vbCrLf
45050       Else
45060         Select Case Fx.Result
                Case "N": rtb.SelText = "Negative" & vbCrLf
45070           Case "P": rtb.SelText = "Positive" & vbCrLf
45080           Case Else: rtb.SelText = "Not available" & vbCrLf
45090         End Select
45100       End If
45110     End If
45120   Next

        'Rota and Adeno
45130   Set Fx = Fxs.Item("Rota")
45140   If Not Fx Is Nothing Then
45150     rtb.SelText = vbCrLf
45160     rtb.SelBold = False
45170     rtb.SelText = "Rota Virus : "
45180     rtb.SelBold = True
45190     If Fx.Valid = 0 Then
45200       rtb.SelText = "Not yet available." & vbCrLf
45210     Else
45220       Select Case Fx.Result
              Case "N": rtb.SelText = "Negative" & vbCrLf
45230         Case "P": rtb.SelText = "Positive" & vbCrLf
45240       End Select
45250     End If
45260   End If
45270   Set Fx = Fxs.Item("Adeno")
45280   If Not Fx Is Nothing Then
45290     rtb.SelBold = False
45300     rtb.SelText = vbCrLf
45310     rtb.SelBold = False
45320     rtb.SelText = "Adeno Virus : "
45330     rtb.SelBold = True
45340     If Fx.Valid = 0 Then
45350       rtb.SelText = "Not yet available." & vbCrLf
45360     Else
45370       Select Case Fx.Result
              Case "N": rtb.SelText = "Negative" & vbCrLf
45380         Case "P": rtb.SelText = "Positive" & vbCrLf
45390       End Select
45400     End If
45410   End If

        'C.diff
45420   Set Fx = Fxs.Item("ToxinAL")
45430   If Not Fx Is Nothing Then
45440     rtb.SelText = vbCrLf
45450     rtb.SelBold = False
45460     rtb.SelText = "C. difficile : "
45470     rtb.SelBold = True
45480     If Fx.Valid = 0 Then
45490       rtb.SelText = "Not yet available." & vbCrLf
45500     Else
45510       Select Case Fx.Result
              Case "N": rtb.SelText = "Not detected" & vbCrLf
45520         Case "P": rtb.SelText = "Positive" & vbCrLf
45530         Case "I": rtb.SelText = "Inconclusive" & vbCrLf
45540         Case "R": rtb.SelText = "Sample Rejected" & vbCrLf
45550       End Select
45560     End If
45570   End If

        'Cryptosporidium
45580   Set Fx = Fxs.Item("AUS")
45590   If Not Fx Is Nothing Then
45600     rtb.SelText = vbCrLf
45610     rtb.SelBold = False
45620     rtb.SelText = "Cryptosporidium : "
45630     rtb.SelBold = True
45640     If Fx.Valid = 0 Then
45650       rtb.SelText = "Not yet available." & vbCrLf
45660     Else
45670       Select Case Fx.Result
              Case "N": rtb.SelText = "Negative" & vbCrLf
45680         Case "P": rtb.SelText = "Positive" & vbCrLf
45690       End Select
45700     End If
45710   End If

        'Ova/Parasites
45720   Set Fx = Fxs.Item("OP0")
45730   If Not Fx Is Nothing Then
45740     rtb.SelText = vbCrLf
45750     rtb.SelBold = False
45760     rtb.SelText = "Ova/Parasites(1) : "
45770     rtb.SelBold = True
45780     If Fx.Valid = 0 Then
45790       rtb.SelText = "Not yet available." & vbCrLf
45800     Else
45810       rtb.SelText = Fx.Result & vbCrLf
45820     End If
45830   End If
45840   Set Fx = Fxs.Item("OP1")
45850   If Not Fx Is Nothing Then
45860     rtb.SelText = vbCrLf
45870     rtb.SelBold = False
45880     rtb.SelText = "Ova/Parasites(2) : "
45890     rtb.SelBold = True
45900     If Fx.Valid = 0 Then
45910       rtb.SelText = "Not yet available." & vbCrLf
45920     Else
45930       rtb.SelText = Fx.Result & vbCrLf
45940     End If
45950   End If
45960   Set Fx = Fxs.Item("OP2")
45970   If Not Fx Is Nothing Then
45980     rtb.SelText = vbCrLf
45990     rtb.SelBold = False
46000     rtb.SelText = "Ova/Parasites(3) : "
46010     rtb.SelBold = True
46020     If Fx.Valid = 0 Then
46030       rtb.SelText = "Not yet available." & vbCrLf
46040     Else
46050       rtb.SelText = Fx.Result & vbCrLf
46060     End If
46070   End If
        
46080 End If

46090 Exit Sub

FillFaecesRTB_Error:

      Dim strES As String
      Dim intEL As Integer

46100 intEL = Erl
46110 strES = Err.Description
46120 LogError "frmMicroReport", "FillFaecesRTB", intEL, strES

End Sub


Private Sub FillGrid()

          Dim sqlBase As String
          Dim sql As String
          Dim tb As Recordset
          Dim s As String
          Dim Cn As Integer

46130     On Error GoTo FillGrid_Error

46140     With grdSID
46150         .ColWidth(6) = 0    'Cn
46160         .Rows = 2
46170         .AddItem ""
46180         .RemoveItem 1
46190     End With

46200     sqlBase = "SELECT D.*, P.PrintedDateTime " & _
                    "FROM Demographics D LEFT JOIN PrintValidLog P " & _
                    "ON D.SampleID = P.SampleID " & _
                    "WHERE PatName = '" & AddTicks(lblName) & "' "
46210     If Trim$(lblChart) <> "" Then
46220         sqlBase = sqlBase & "AND Chart = '" & lblChart & "' "
46230     Else
46240         sqlBase = sqlBase & "AND ( Chart IS NULL OR Chart = '' ) "
46250     End If
46260     If IsDate(lblDoB) Then
46270         sqlBase = sqlBase & "AND DoB = '" & Format$(lblDoB, "dd/mmm/yyyy") & "' "
46280     Else
46290         sqlBase = sqlBase & "AND ( DoB IS NULL OR DoB = '' ) "
46300     End If
46310     sqlBase = sqlBase & "AND D.SampleID > "

46320     For Cn = 0 To intOtherHospitalsInGroup
46330         Set tb = New Recordset
46340         sql = sqlBase & sysOptMicroOffset(Cn) & " " & _
                    "--AND SampleDate > '01/Jan/2007'"
46350         RecOpenClient Cn, tb, sql

46360         Do While Not tb.EOF
46370             s = Format$(Val(tb!SampleID) - sysOptMicroOffset(Cn)) & vbTab & _
                      tb!Rundate & vbTab
46380             If IsDate(tb!SampleDate) Then
46390                 If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
46400                     s = s & Format(tb!SampleDate, "dd/MM/yy hh:mm")
46410                 Else
46420                     s = s & Format(tb!SampleDate, "dd/MM/yy")
46430                 End If
46440             Else
46450                 s = s & "Not Specified"
46460             End If
46470             s = s & vbTab
46480             If Not IsNull(tb!PrintedDateTime) Then
46490                 s = s & Format$(tb!PrintedDateTime, "dd/MM/yy HH:mm")
46500             End If
46510             s = s & vbTab
46520             s = s & LoadOutstandingMicro(tb!SampleID)
46530             s = s & vbTab & HospName(Cn) & vbTab & Format$(Cn)
46540             grdSID.AddItem s

46550             lblAge = tb!Age & ""
46560             Select Case Left$(UCase$(tb!Sex & ""), 1)
                  Case "M": lblSex = "Male"
46570             Case "F": lblSex = "Female"
46580             Case Else: lblSex = ""
46590             End Select
46600             lblAddress = tb!Addr0 & " " & tb!Addr1 & ""
46610             tb.MoveNext
46620         Loop
46630     Next

46640     If grdSID.Rows > 2 Then
46650         grdSID.RemoveItem 1
46660     End If
46670     grdSID.Col = 1
46680     grdSID.Sort = 9

46690     Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

46700     intEL = Erl
46710     strES = Err.Description
46720     LogError "frmMicroReport", "FillGrid", intEL, strES, sql


End Sub
Private Sub FillSensitivityResults(ByVal Sxs As Sensitivities, ByVal IsolateNumber As Integer, ByVal Report As Integer)

      Dim sx As Sensitivity

46730 For Each sx In Sxs
46740   If sx.IsolateNumber = IsolateNumber And sx.Report = Report Then
          'Only report if R/S/I not for X
46750     If Trim$(sx.RSI) <> "" And sx.RSI <> "X" Then
46760       rtb.SelFontName = "Courier New"
46770       rtb.SelColor = vbBlack
46780       rtb.SelText = Left$(sx.AntibioticName & Space$(20), 20)
46790       If sx.RSI = "R" Then
46800         rtb.SelColor = vbRed
46810         rtb.SelFontName = "Courier New"
46820         rtb.SelText = "Resistant" & vbCrLf
46830       ElseIf sx.RSI = "S" Then
46840         rtb.SelColor = vbGreen
46850         rtb.SelFontName = "Courier New"
46860         rtb.SelText = "Sensitive" & vbCrLf
46870       ElseIf sx.RSI = "I" Then
46880         rtb.SelColor = vbBlue
46890         rtb.SelFontName = "Courier New"
46900         rtb.SelText = "Intermediate" & vbCrLf
46910       End If
46920     End If
46930   End If
46940 Next

End Sub


Private Function LoadOutstandingMicro(ByVal SampleIDWithOffset As String) As String

      Dim sql As String
      Dim s As String
      Dim UR As UrineRequest
      Dim URS As New UrineRequests
      Dim fr As FaecesRequest
      Dim FRs As New FaecesRequests
      Dim SDS As New SiteDetails

46950 On Error GoTo LoadOutstandingMicro_Error

46960 SDS.Load SampleIDWithOffset
46970 If SDS.Count > 0 Then
46980   s = SDS(1).Site & " " & SDS(1).SiteDetails & " "
46990   If SDS(1).Site = "Faeces" Then
47000     FRs.Load SampleIDWithOffset
47010     For Each fr In FRs
47020       s = s & fr.Request & " "
47030     Next
47040   ElseIf SDS(1).Site = "Urine" Then
47050     URS.Load SampleIDWithOffset
47060     For Each UR In URS
47070       s = s & UR.Request & " "
47080     Next
47090   End If
47100 End If
47110 LoadOutstandingMicro = s

47120 Exit Function

LoadOutstandingMicro_Error:

      Dim strES As String
      Dim intEL As Integer

47130 intEL = Erl
47140 strES = Err.Description
47150 LogError "frmMicroReport", "LoadOutstandingMicro", intEL, strES, sql

End Function

Private Sub bPrint_Click()
      Dim i As Integer

47160 SelectedIndex = 0
47170 If UCase$(HospName(0)) = "CAVAN" Then
47180     For i = 1 To grdSID.Rows - 1
47190         grdSID.Col = 1
47200         grdSID.row = i
47210         If grdSID.CellBackColor = vbYellow Then
47220             SelectedIndex = i
47230         End If
47240     Next
47250     If SelectedIndex = 0 Then
47260         iMsg "Please select report first", vbInformation
47270         Exit Sub
47280     End If
          
47290     PrintThis

47300 End If
End Sub

Private Sub cmdCancel_Click()

47310 pReturnedSampleID = ""
47320 Me.Hide

End Sub

Private Sub cmdPhone_Click()

      Dim SID As String

47330 On Error GoTo cmdPhone_Click_Error

47340 SID = grdSID.TextMatrix(grdSID.row, 0)
47350 If Val(SID) > 0 Then
        
47360   With frmPhoneLog
47370     .SampleID = Val(SID) ' + sysOptMicroOffset(0)
47380     .Discipline = "Micro"
47390     .GP = ""
47400     .WardOrGP = "Ward"
47410     .Show 1
47420   End With
        
47430   CheckIfPhoned

47440 End If

47450 Exit Sub

cmdPhone_Click_Error:

      Dim strES As String
      Dim intEL As Integer

47460 intEL = Erl
47470 strES = Err.Description
47480 LogError "frmMicroReport", "cmdPhone_Click", intEL, strES

End Sub

Private Sub cmdSetPrinter_Click()

47490 frmForcePrinter.From = Me
47500 frmForcePrinter.Show 1

47510 If pPrintToPrinter = "Automatic Selection" Then
47520   pPrintToPrinter = ""
47530 End If

47540 If pPrintToPrinter <> "" Then
47550   cmdSetPrinter.BackColor = vbRed
47560   cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
47570 Else
47580   cmdSetPrinter.BackColor = vbButtonFace
47590   pPrintToPrinter = ""
47600   cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
47610 End If

End Sub

Private Sub cmdView_Click()

      Dim f As Form

47620 If pFromEdit Then
47630   Me.Hide
47640 Else
47650   Set f = New frmEditMicrobiology
47660   With f
47670     .FromViewReportSID = pReturnedSampleID
47680     .Show 1
47690   End With
47700   Unload f
47710   Set f = Nothing
47720 End If

End Sub

Private Sub Form_Activate()

47730     pBar.max = LogOffDelaySecs
47740     pBar = 0

47750     Timer1.Enabled = True

47760     If Activated Then Exit Sub
47770     Activated = True

47780     FillGrid

End Sub

Private Sub CheckIfPhoned()

      Dim s As String
      Dim PhLog As PhoneLog
      Dim sql As String
      Dim OBs As Observations
      Dim SID As String

47790 On Error GoTo CheckIfPhoned_Error

47800 SID = grdSID.TextMatrix(grdSID.row, 0)
47810 If Val(SID) > 0 Then

      '40      PhLog = CheckPhoneLog(Val(SID) + sysOptMicroOffset(0))
47820   PhLog = CheckPhoneLog(Val(SID))
47830   If PhLog.SampleID <> 0 Then
47840     cmdPhone.BackColor = vbYellow
47850     cmdPhone.Caption = "Results Phoned"
47860     cmdPhone.ToolTipText = "Results Phoned"
47870     If InStr(lblDemogComment, "Results Phoned") = 0 Then
47880       s = "Results Phoned to " & PhLog.PhonedTo & " at " & _
                Format$(PhLog.DateTime, "hh:mm") & " on " & Format$(PhLog.DateTime, "dd/MM/yyyy") & _
                " by " & PhLog.PhonedBy & "."
47890       If Trim$(lblDemogComment) = "" Then
47900         lblDemogComment = s
47910       Else
47920         lblDemogComment = lblDemogComment & ". " & s
47930       End If
47940       Set OBs = New Observations
47950       OBs.Save PhLog.SampleID, True, "Demographic", lblDemogComment
          
47960     End If
47970   Else
47980     cmdPhone.BackColor = &H8000000F
47990     cmdPhone.Caption = "Phone Results"
48000     cmdPhone.ToolTipText = "Phone Results"
48010   End If

48020 End If

48030 Exit Sub

CheckIfPhoned_Error:

      Dim strES As String
      Dim intEL As Integer

48040 intEL = Erl
48050 strES = Err.Description
48060 LogError "frmMicroReport", "CheckIfPhoned", intEL, strES, sql

End Sub

Private Sub Form_Deactivate()

48070     Timer1.Enabled = False

End Sub


Private Sub Form_Load()

48080     Activated = False

48090     pBar.max = LogOffDelaySecs
48100     pBar = 0

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

48110     pBar = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

48120 Activated = False
48130 pPrintToPrinter = ""

End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

48140     pBar = 0

End Sub


Private Sub grdSID_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim Cn As String
      '+++ Junaid
      Dim f As Form

48150 Set f = New frmReportViewer
      '--- Junaid
48160 On Error GoTo grdSID_Click_Error

48170 rtb.Text = ""

48180 lblDemogComment = ""

48190 If grdSID.MouseRow = 0 Then
48200     If SortOrder Then
48210         grdSID.Sort = flexSortGenericAscending
48220     Else
48230         grdSID.Sort = flexSortGenericDescending
48240     End If
48250     SortOrder = Not SortOrder
48260     Exit Sub
48270 End If

48280 If grdSID.Rows = 2 And grdSID.TextMatrix(1, 0) = "" Then Exit Sub

48290 For Y = 1 To grdSID.Rows - 1
48300     grdSID.row = Y
48310     For X = 1 To grdSID.Cols - 1
48320         grdSID.Col = X
48330         grdSID.CellBackColor = 0
48340     Next
48350 Next

48360 grdSID.row = grdSID.MouseRow
48370 For X = 1 To grdSID.Cols - 1
48380     grdSID.Col = X
48390     grdSID.CellBackColor = vbYellow
48400 Next

48410 Cn = grdSID.TextMatrix(grdSID.row, 6)

48420 pReturnedSampleID = grdSID.TextMatrix(grdSID.row, 0)
48430 cmdView.Caption = "View/Edit " & pReturnedSampleID

48440 cmdView.Visible = True
      '+++ Junaid


          

48450     f.Dept = "Microbiology"
48460     f.SampleID = pReturnedSampleID
48470     f.Show 1

48480     Set f = Nothing
      '--- Junaid

      'RP.sampleid = pReturnedSampleID & ""
      '300   RTFPrintResultMicro (pReturnedSampleID & "")
      '310   rtb.Text = ""
      '320   If optNew Then
      '330       rtb.SelText = frmMicroReport.rtb
      '340       FillCommentsRTB pReturnedSampleID, Cn
      '350   ElseIf optOld Then
      '360       FillCommentsRTB pReturnedSampleID, Cn
      '370       FillResultMicroRTB pReturnedSampleID, Cn
      '380   End If
      '      '330   FillCommentsRTB pReturnedSampleID, Cn


48490 CheckIfPhoned

48500 Exit Sub

grdSID_Click_Error:

      Dim strES As String
      Dim intEL As Integer

48510 intEL = Erl
48520 strES = Err.Description
48530 LogError "frmMicroReport", "grdSID_Click", intEL, strES

End Sub
Private Sub FillResultMicroRTB(ByVal SampleID As String, _
                               ByVal Cn As Integer)

          Dim sql As String
          Dim tb As Recordset

48540     On Error GoTo FillResultMicroRTB_Error

48550     lblClDetails = ""

48560     sql = "Select ClDetails from Demographics where " & _
                "SampleID = '" & Val(SampleID) + sysOptMicroOffset(Cn) & "'"
48570     Set tb = New Recordset
48580     RecOpenClient Cn, tb, sql
48590     If Not tb.EOF Then
48600         lblClDetails = tb!ClDetails & ""
48610     End If

48620     WordResultPrinted = False

48630     FillMicroscopyRTB SampleID, Cn

      'FillIsolates SampleID

48640     FillGramStainRTB SampleID
          
48650     FillGenericResultsRTB SampleID

48660     FillFaecesRTB SampleID, Cn

48670     FillMicroUrineComment SampleID, Cn

48680     FillCSF SampleID

48690     FillSensitivitiesRTB SampleID

48700     Exit Sub

FillResultMicroRTB_Error:

          Dim strES As String
          Dim intEL As Integer

48710     intEL = Erl
48720     strES = Err.Description
48730     LogError "frmMicroReport", "FillResultMicroRTB", intEL, strES, sql


End Sub

Private Sub FillMicroscopyRTB(ByVal SampleID As String, _
                              ByVal Cn As Integer)

      Dim Ux As UrineResult
      Dim Uxs As New UrineResults
      Dim R As String

48740 On Error GoTo FillMicroscopyRTB_Error

48750 Uxs.Load Val(SampleID) + sysOptMicroOffset(Cn)
48760 If Uxs.Count = 0 Then Exit Sub

48770 If Not WordResultPrinted Then
48780   rtb.SelColor = vbBlue
48790   rtb.SelBold = True
48800   rtb.SelFontSize = 12
48810   rtb.SelText = "Results:" & vbCrLf
48820   WordResultPrinted = True
48830 End If

48840 Set Ux = Uxs.Item("Pregnancy")
48850 If Not Ux Is Nothing Then
48860   rtb.SelColor = vbBlack
48870   rtb.SelBold = False
48880   rtb.SelFontSize = 10
48890   rtb.SelText = "Pregnancy: " & Ux.Result & vbCrLf & vbCrLf
48900   rtb.SelFontName = "Courier New"
48910 End If



48920    rtb.SelColor = vbBlack
48930    rtb.SelBold = False
48940    rtb.SelFontSize = 10
48950    rtb.SelText = "Microscopy:" & vbCrLf

48960    rtb.SelFontName = "Courier New"
48970 rtb.SelBold = False
48980 rtb.SelText = "   Bacteria: "
48990 rtb.SelBold = True
49000    Set Ux = Uxs.Item("Bacteria")
49010    If Ux Is Nothing Then R = "" Else R = Ux.Result
49020 rtb.SelText = Left$(Trim$(R) & Space(20), 20)

49030 rtb.SelBold = False
49040 rtb.SelText = "Crystals:"
49050 rtb.SelBold = True
49060 Set Ux = Uxs.Item("Crystals")
49070 If Ux Is Nothing Then R = "" Else R = Ux.Result
49080 rtb.SelText = Trim$(R) & vbCrLf

49090 rtb.SelFontName = "Courier New"
49100 rtb.SelBold = False
49110 rtb.SelText = "        WCC: "
49120 rtb.SelBold = True
49130 Set Ux = Uxs.Item("WCC")
49140 If Ux Is Nothing Then R = "" Else R = Ux.Result
49150 rtb.SelText = Left$(Trim$(R) & Space(20), 20)

49160 rtb.SelBold = False
49170 rtb.SelText = "   Casts:"
49180 rtb.SelBold = True
49190 Set Ux = Uxs.Item("Casts")
49200 If Ux Is Nothing Then R = "" Else R = Ux.Result
49210 rtb.SelText = Trim$(R) & vbCrLf

49220 rtb.SelFontName = "Courier New"
49230 rtb.SelBold = False
49240 rtb.SelText = "        RCC: "
49250 rtb.SelBold = True
49260 Set Ux = Uxs.Item("RCC")
49270 If Ux Is Nothing Then R = "" Else R = Ux.Result
49280 rtb.SelText = Left$(Trim$(R) & Space(20), 20)

49290 rtb.SelBold = False
49300 rtb.SelText = "    Misc: "
49310 rtb.SelBold = True
49320 Set Ux = Uxs.Item("Misc0")
49330 If Ux Is Nothing Then R = "" Else R = Ux.Result
49340 rtb.SelText = Trim$(R)
49350 Set Ux = Uxs.Item("Misc1")
49360 If Ux Is Nothing Then R = "" Else R = Ux.Result
49370 rtb.SelText = " " & Trim$(R)
49380 Set Ux = Uxs.Item("Misc2")
49390 If Ux Is Nothing Then R = "" Else R = Ux.Result
49400 rtb.SelText = " " & Trim$(R) & vbCrLf

49410 rtb.SelFontName = "MS Sans Serif"
49420 rtb.SelText = vbCrLf

49430 Exit Sub

FillMicroscopyRTB_Error:

      Dim strES As String
      Dim intEL As Integer

49440 intEL = Erl
49450 strES = Err.Description
49460 LogError "frmMicroReport", "FillMicroscopyRTB", intEL, strES

End Sub
Private Sub FillGramStainRTB(ByVal SampleID As String)

      Dim ID As IdentResult
      Dim IDs As New IdentResults
      Dim Title As String

49470 On Error GoTo FillGramStainRTB_Error

49480 IDs.Load SampleID
49490 If Not IDs Is Nothing Then
49500   Title = "Gram Stain: "
49510   For Each ID In IDs
49520     If UCase$(ID.TestType) = "GRAMSTAIN" Then
49530       rtb.SelFontName = "Courier New"
49540       rtb.SelColor = vbBlack
49550       rtb.SelBold = False
49560       rtb.SelFontSize = 10
49570       rtb.SelText = Title & ID.TestName & " " & ID.Result & vbCrLf
49580       Title = "            "
49590     End If
49600   Next

49610   Title = "  Wet Prep: "
49620   For Each ID In IDs
49630     If UCase$(ID.TestType) = "WETPREP" Then
49640       rtb.SelFontName = "Courier New"
49650       rtb.SelColor = vbBlack
49660       rtb.SelBold = False
49670       rtb.SelFontSize = 10
49680       rtb.SelText = Title & ID.TestName & " " & ID.Result & vbCrLf
49690       Title = "            "
49700     End If
49710   Next
49720 End If

49730 Exit Sub

FillGramStainRTB_Error:

      Dim strES As String
      Dim intEL As Integer

49740 intEL = Erl
49750 strES = Err.Description
49760 LogError "frmMicroReport", "FillGramStainRTB", intEL, strES

End Sub


Private Sub FillCSF(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String

49770     On Error GoTo FillCSF_Error

      '20        sql = "SELECT * FROM CSFResults WHERE " & _
      '                "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
49780     sql = "SELECT * FROM CSFResults WHERE " & _
                "SampleID = '" & Val(SampleID) & "'"
49790     Set tb = New Recordset
49800     RecOpenServer 0, tb, sql
49810     If tb.EOF Then Exit Sub

49820     With rtb
49830         .SelFontName = "Courier New"
49840         .SelColor = vbBlack
49850         .SelBold = False
49860         .SelFontSize = 10
49870         .SelText = "Appearance:" & Space(40) & "Gram Stain" & vbCrLf

49880         .SelFontName = "Courier New"
49890         .SelFontSize = 10
49900         .SelBold = False
49910         .SelText = "Sample 1 "
49920         .SelBold = True
49930         .SelText = Left$(tb!Appearance0 & Space(42), 42) & tb!Gram & vbCrLf

49940         .SelFontName = "Courier New"
49950         .SelFontSize = 10
49960         .SelBold = False
49970         .SelText = "Sample 2 "
49980         .SelBold = True
49990         .SelText = tb!Appearance1 & vbCrLf

50000         .SelFontName = "Courier New"
50010         .SelFontSize = 10
50020         .SelBold = False
50030         .SelText = "Sample 3 "
50040         .SelBold = True
50050         .SelText = tb!Appearance2 & vbCrLf & vbCrLf

50060         .SelFontName = "Courier New"
50070         .SelFontSize = 10
50080         .SelBold = False
50090         .SelText = "        Sample 1        Sample 2        Sample 3        White Cell Differential" & vbCrLf

50100         .SelFontName = "Courier New"
50110         .SelFontSize = 10
50120         .SelBold = False
50130         .SelText = "WCC/cmm    "
50140         .SelBold = True
50150         .SelText = Left$(tb!WCC0 & Space(16), 16)
50160         .SelText = Left$(tb!WCC1 & Space(16), 16)
50170         .SelText = Left$(tb!WCC2 & Space(16), 16)
50180         .SelText = tb!WCCDiff0 & " % Neutrophils" & vbCrLf

50190         .SelFontName = "Courier New"
50200         .SelFontSize = 10
50210         .SelBold = False
50220         .SelText = "RCC/cmm    "
50230         .SelBold = True
50240         .SelText = Left$(tb!RCC0 & Space(16), 16)
50250         .SelText = Left$(tb!RCC1 & Space(16), 16)
50260         .SelText = Left$(tb!RCC2 & Space(16), 16)
50270         .SelText = tb!WCCDiff1 & " % Mononuclear Cells" & vbCrLf

50280         .SelFontName = "MS Sans Serif"
50290         .SelText = vbCrLf
50300     End With

50310     Exit Sub

FillCSF_Error:

          Dim strES As String
          Dim intEL As Integer

50320     intEL = Erl
50330     strES = Err.Description
50340     LogError "frmMicroReport", "FillCSF", intEL, strES, sql

End Sub


Private Sub FillSensitivitiesRTB(ByVal SampleID As String)

      Dim sx As Sensitivity
      Dim Sxs As New Sensitivities
      Dim tb As Recordset
      Dim sql As String
      Dim strGroup(1 To 8) As OrgGroup
      Dim SampleIDWithOffset As Long
      Dim SensPrintMax As Integer
      Dim X As Integer
      Dim SDS As New SiteDetails
      Dim Ix As Isolate
      Dim Ixs As New Isolates

50350 On Error GoTo FillSensitivitiesRTB_Error

50360 SampleIDWithOffset = Val(SampleID) '+ sysOptMicroOffset(0)

50370 SensPrintMax = 3
      '+++ Junaid 20-05-2024
      '40    SDS.Load SampleIDWithOffset
50380 SDS.Load Trim(SampleID)
      '--- Junaid
50390 If SDS.Count > 0 Then
50400   sql = "Select [Default] as D from Lists where " & _
              "ListType = 'SI' " & _
              "and Text = '" & SDS(1).Site & "'"
50410   Set tb = New Recordset
50420   RecOpenServer 0, tb, sql
50430   If Not tb.EOF Then
50440     SensPrintMax = Val(tb!D & "")
50450   End If
50460 End If
      '+++ Junaid 20-05-2024
      '130   FillOrgGroups strGroup(), SampleIDWithOffset
      '
      '140   Ixs.Load Val(SampleID) + sysOptMicroOffset(0)

50470 FillOrgGroups strGroup(), Trim(SampleID)

50480 Ixs.Load Val(SampleID)
      '--- Junaid

50490 If Not Ixs.Count = 0 Then
50500   Set Ix = Ixs(1)
50510   If Not Ix Is Nothing Then
50520     If Not WordResultPrinted Then
50530       rtb.SelColor = vbBlue
50540       rtb.SelBold = True
50550       rtb.SelFontSize = 12
50560       rtb.SelText = "Results:" & vbCrLf
50570       WordResultPrinted = True
50580     End If
50590     rtb.SelColor = vbBlack
50600     rtb.SelBold = False
50610     rtb.SelFontSize = 10
50620     rtb.SelText = vbCrLf
50630     rtb.SelFontSize = 10
50640     If Ix.Valid = 0 Or Ix.OrganismGroup = "" Then
50650       rtb.SelText = "Isolates not yet available." & vbCrLf
50660     Else
50670       sql = "UPDATE Sensitivities SET Valid = 1 WHERE SampleID = '" & Trim(SampleID) & "' "
50680       Cnxn(0).Execute sql
50690     End If
50700   End If
50710 End If

50720 Sxs.Load Trim(SampleID)
50730 If Sxs.Count > 0 Then
50740   For X = 1 To 4
50750     For Each sx In Sxs
50760       If strGroup(X).OrgName <> "" And sx.Valid Then
        
50770         rtb.SelFontSize = 10
50780         rtb.SelBold = True
50790         rtb.SelText = "Culture " & X & " : " & _
                            strGroup(X).OrgName & " " & _
                            strGroup(X).Qualifier & vbCrLf

50800         FillSensitivityResults Sxs, X, 1

           'Not for reporting

50810         If AnyNotForReporting(Sxs, X) Then
50820           rtb.SelColor = vbBlack
50830           rtb.SelBold = False
50840           rtb.SelFontSize = 10
50850           rtb.SelText = vbCrLf
50860           rtb.SelFontSize = 10
50870           rtb.SelText = "Sensitivities not for reporting:-" & vbCrLf
50880           rtb.SelFontName = "MS Sans Serif"
50890           rtb.SelColor = vbBlack
50900           rtb.SelText = vbCrLf
50910           FillSensitivityResults Sxs, X, False
50920         End If
50930         rtb.SelText = vbCrLf
50940         Exit For
50950       End If
50960     Next
50970   Next
50980 Else
50990   For Each Ix In Ixs
51000     If Ix.Valid Then
51010       rtb.SelFontSize = 10
51020       rtb.SelText = Ix.OrganismName & " " & Ix.Qualifier & vbCrLf
51030     End If
51040   Next
51050 End If

51060 Exit Sub

FillSensitivitiesRTB_Error:

          Dim strES As String
          Dim intEL As Integer

51070     intEL = Erl
51080     strES = Err.Description
51090     LogError "frmMicroReport", "FillSensitivitiesRTB", intEL, strES, sql

End Sub
Private Sub FillMicroUrineComment(ByVal SampleID As String, _
                                  ByVal Cn As Integer)

          Dim tb As Recordset
          Dim sql As String


51100 On Error GoTo FillMicroUrineComment_Error

      Dim SDS As New SiteDetails
51110 SDS.Load Val(SampleID) + sysOptMicroOffset(Cn)
51120 If SDS.Count > 0 Then
51130   If SDS(1).Site = "Urine" Then

51140     sql = "Select * from Isolates where " & _
                "SampleID = '" & Val(SampleID) + sysOptMicroOffset(Cn) & "' " & _
                "AND OrganismGroup <> 'Negative results' " & _
                "AND OrganismName <> ''"
51150     Set tb = New Recordset
51160     RecOpenServer 0, tb, sql
51170     If tb.EOF Then Exit Sub

51180     rtb.SelText = vbCrLf
51190     rtb.SelColor = vbBlack
51200     rtb.SelBold = False
51210     rtb.SelFontSize = 10

51220     rtb.SelText = "Positive cultures "
51230     rtb.SelUnderline = True
51240     rtb.SelText = "must"
51250     rtb.SelUnderline = False
51260     rtb.SelText = " be correlated with signs and symptoms of UTI "
51270     rtb.SelText = "Particularly with low colony counts"
51280     rtb.SelText = vbCrLf
51290   End If
51300 End If


51310 Exit Sub

FillMicroUrineComment_Error:

      Dim strES As String
      Dim intEL As Integer

51320 intEL = Erl
51330 strES = Err.Description
51340 LogError "frmMicroReport", "FillMicroUrineComment", intEL, strES, sql


End Sub

Private Function FillOrgGroups(ByRef strGroup() As OrgGroup, _
                               ByVal SampleIDWithOffset As Long) _
                               As Integer

      Dim tbO As Recordset
      Dim sql As String
      Dim n As Integer
      Dim IsoNum As Integer
      Dim Isos As New Isolates
      Dim Iso As Isolate

51350 On Error GoTo FillOrgGroups_Error

51360 Isos.Load SampleIDWithOffset
51370 n = 1
51380 For Each Iso In Isos
51390   IsoNum = Iso.IsolateNumber
51400   With strGroup(IsoNum)
51410     .OrgGroup = Iso.OrganismGroup
51420     If Trim$(Iso.OrganismName) = "" Then
51430       .OrgName = .OrgGroup
51440     Else
51450       .OrgName = Iso.OrganismName
51460     End If
51470     .Qualifier = Iso.Qualifier
51480     sql = "Select ShortName, ReportName from Organisms where " & _
                "Name = '" & Iso.OrganismName & "'"
51490     Set tbO = New Recordset
51500     RecOpenClient 0, tbO, sql
51510     If Not tbO.EOF Then
51520       .ShortName = tbO!ShortName & ""
51530       .ReportName = Trim$(tbO!ReportName & "")
51540     Else
51550       .ShortName = .OrgName
51560       .ReportName = .OrgName
51570     End If
51580     If .ReportName = "" Then
51590       .ShortName = .OrgName
51600       .ReportName = .OrgName
51610       .OrgName = .OrgName
51620     End If
51630   End With
51640   n = n + 1
51650 Next

51660 FillOrgGroups = n - 1

51670 Exit Function

FillOrgGroups_Error:

      Dim strES As String
      Dim intEL As Integer

51680 intEL = Erl
51690 strES = Err.Description
51700 LogError "frmMicroReport", "FillOrgGroups", intEL, strES, sql

End Function

Private Sub grdSID_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As Date
          Dim d2 As Date
          Dim Column As Integer

51710     With grdSID
51720         Column = .Col
51730         Cmp = 0
51740         If IsDate(.TextMatrix(Row1, Column)) Then
51750             d1 = Format(.TextMatrix(Row1, Column), "dd/mmm/yyyy")
51760             If IsDate(.TextMatrix(Row2, Column)) Then
51770                 d2 = Format(.TextMatrix(Row2, Column), "dd/mmm/yyyy")
51780                 Cmp = Sgn(DateDiff("d", d1, d2))
51790             End If
51800         End If
51810     End With

End Sub





Private Sub grdSID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51820     pBar = 0

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

51830     pBar = 0

End Sub


Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

51840     pBar = 0

End Sub


Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

51850     pBar = 0

End Sub


Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

51860     pBar = 0

End Sub


Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

51870     pBar = 0

End Sub


Private Sub Label7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

51880     pBar = 0

End Sub


Private Sub lblAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51890     pBar = 0

End Sub


Private Sub lblAge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51900     pBar = 0

End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51910     pBar = 0

End Sub


Private Sub lblClDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51920     pBar = 0

End Sub


Private Sub lblDemogComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51930     pBar = 0

End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51940     pBar = 0

End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51950     pBar = 0

End Sub


Private Sub lblSex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

51960     pBar = 0

End Sub


Private Sub optNew_Click()
      Dim X As Integer
51970 rtb = ""
51980 For X = 1 To grdSID.Cols - 1
51990     grdSID.Col = X
52000     grdSID.CellBackColor = &H80000018
52010 Next
End Sub

Private Sub optOld_Click()
      Dim X As Integer
52020 rtb = ""
52030 For X = 1 To grdSID.Cols - 1
52040     grdSID.Col = X
52050     grdSID.CellBackColor = &H80000018
52060 Next
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)

52070     KeyAscii = 0

End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

52080     rtb.SelLength = 0
End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
52090     pBar = pBar + 1

52100     If pBar = pBar.max Then
52110         Unload Me
52120     End If

End Sub

Private Sub PrintThis()

      Dim tb As Recordset
      Dim sql As String
      Dim tbP As Recordset
      Dim SampleID As String
      Dim strTime As String
      Dim OB As Observation
      Dim OBs As Observations
      Dim n As Integer

52130 On Error GoTo PrintThis_Error

52140 pBar = 0

52150 SampleID = ""
52160 grdSID.Col = 1
52170 For n = 1 To grdSID.Rows - 1
52180   grdSID.row = n
52190   If grdSID.CellBackColor = vbYellow Then
52200     SampleID = grdSID.TextMatrix(grdSID.row, 0) ' + sysOptMicroOffset(0)
52210     Exit For
52220   End If
52230 Next
52240 If SampleID = "" Then Exit Sub

52250 sql = "SELECT SampleID, PatName, Sex, Ward, Clinician, GP, SampleDate, RecDate " & _
            "FROM Demographics " & _
            "WHERE SampleID = '" & SampleID & "'"
52260 Set tb = New Recordset
52270 RecOpenServer 0, tb, sql

52280 If tb.EOF Then
52290   iMsg "No demographic details found", vbInformation
52300   Exit Sub
52310 End If

52320 If tb!Sex & "" = "" Then
52330   iMsg "Sex is not entered. Please enter sex first", vbInformation
52340   Exit Sub
52350 End If

52360 If SurName(tb!PatName & "") <> "" Then
52370   If Trim$(tb!Ward & "") = "" Then
52380     iMsg "Must have Ward entry.", vbCritical
52390     Exit Sub
52400   End If

52410   If UCase$(Trim$(tb!Ward & "")) = "GP" Then
52420     If Trim$(tb!GP & "") = "" Then
52430       iMsg "Must have Ward or GP entry.", vbCritical
52440       Exit Sub
52450     End If
52460   End If
52470 End If
          
52480 If Format(tb!SampleDate, "hh:mm") = "00:00" Then
52490 Set OBs = New Observations
52500   Set OBs = OBs.Load(SampleID, "Demographic")
52510   If Not OBs Is Nothing Then
52520     Set OB = OBs.Item(1)
52530     If InStr(OB.Comment, "Sample Time Unknown.") = 0 Then
52540       If iMsg("Is Sample Time unknown?", vbQuestion + vbYesNo) = vbYes Then
52550         Set OBs = New Observations
52560         OBs.Save SampleID, True, "Demographic", OB.Comment & " Sample Time Unknown."
52570       Else
52580         strTime = iTIME("Sample Time?")
52590         If IsDate(strTime) Then
52600           tb!SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy") & " " & Format(strTime, "hh:mm")
52610           tb.Update
52620         Else
52630           Exit Sub
52640         End If
52650       End If
52660     End If
52670   Else
52680     If iMsg("Is Sample Time unknown?", vbQuestion + vbYesNo) = vbYes Then
52690       Set OBs = New Observations
52700       OBs.Save SampleID, True, "Demographic", "Sample Time Unknown."
52710     Else
52720       strTime = iTIME("Sample Time?")
52730       If IsDate(strTime) Then
52740         tb!SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy") & " " & Format(strTime, "hh:mm")
52750         tb.Update
52760       Else
52770         Exit Sub
52780       End If
52790     End If
52800   End If
52810 End If

52820 If Format(tb!RecDate, "hh:mm") = "00:00" Then
52830   strTime = iTIME("Received Time?")
52840   If IsDate(strTime) Then
52850     tb!RecDate = Format(tb!RecDate, "dd/MMM/yyyy") & " " & Format(strTime, "hh:mm")
52860     tb.Update
52870   Else
52880     Exit Sub
52890   End If
52900 End If

52910 sql = "SELECT * FROM PrintPending WHERE " & _
            "Department = 'M' " & _
            "AND SampleID = '" & grdSID.TextMatrix(SelectedIndex, 0) & "'"
52920 Set tbP = New Recordset
52930 RecOpenClient 0, tbP, sql
52940 If tbP.EOF Then
52950   tbP.AddNew
52960 End If
52970 tbP!SampleID = grdSID.TextMatrix(SelectedIndex, 0)
52980 tbP!Ward = tb!Ward & ""
52990 tbP!Clinician = tb!Clinician & ""
53000 tbP!GP = tb!GP & ""
53010 tbP!Department = "M"
53020 tbP!Initiator = UserName
53030 tbP!UsePrinter = pPrintToPrinter
53040 tbP!ThisIsCopy = 1
53050 tbP.Update

53060 Exit Sub

PrintThis_Error:

      Dim strES As String
      Dim intEL As Integer

53070 intEL = Erl
53080 strES = Err.Description
53090 LogError "frmEditMicrobiology", "PrintThis", intEL, strES, sql

End Sub


Public Property Get ReturnedSampleID() As String

53100 ReturnedSampleID = pReturnedSampleID

End Property

