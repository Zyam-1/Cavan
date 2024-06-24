VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmMicroReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ward Enquiry --- Microbiology"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13770
   ControlBox      =   0   'False
   HelpContextID   =   10035
   Icon            =   "frmMicroReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   13770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSignOffMicro 
      Caption         =   "Sign OFF"
      Height          =   1000
      Left            =   10110
      Picture         =   "frmMicroReport.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   615
      Width           =   1100
   End
   Begin VB.ComboBox cmbDays 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmMicroReport.frx":1794
      Left            =   10110
      List            =   "frmMicroReport.frx":17A8
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   180
      Width           =   3465
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1000
      Left            =   11280
      Picture         =   "frmMicroReport.frx":1804
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   615
      Width           =   1100
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   7005
      Left            =   120
      TabIndex        =   18
      Top             =   3705
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   12356
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMicroReport.frx":26CE
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   1935
      Left            =   90
      TabIndex        =   16
      Top             =   1725
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   11
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmMicroReport.frx":2750
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13380
      Top             =   4140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1000
      Left            =   12450
      Picture         =   "frmMicroReport.frx":2849
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   615
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   9750
      Begin VB.CommandButton cmdViewScan 
         Caption         =   "&View Scan"
         Height          =   1155
         Left            =   8550
         Picture         =   "frmMicroReport.frx":3713
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblClDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   150
         TabIndex        =   17
         Top             =   1110
         Width           =   8325
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
         Left            =   4560
         TabIndex        =   10
         Top             =   510
         Width           =   3915
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
         Left            =   7755
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
         Left            =   7440
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
         Top             =   840
         Width           =   8325
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
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

Private Activated As Boolean

Private Type OrgGroup
    OrgGroup As String
    OrgName As String
    ShortName As String
    ReportName As String
    Qualifier As String
End Type

Dim WordResultPrinted As Boolean
Public ReportDept As String
Dim SampleIDForSignOff As String
Private SortOrder As Boolean
Private Sub FillCommentsRTB(ByVal SampleID As String)

      Dim OB As Observation
      Dim OBS As Observations
      Dim S As String
      Dim sql As String
      Dim tb As Recordset
      Dim CommentsFound As Boolean
      Dim DemComment As String
      Dim ConsComment As String
      Dim CSComment As String

10    On Error GoTo FillCommentsRTB_Error

20    CommentsFound = False
30    Set OBS = New Observations
40    Set OBS = OBS.Load(Val(SampleID) + sysOptMicroOffset(0), "Demographic", "MicroConsultant", "MicroCS")
50    If Not OBS Is Nothing Then
60        For Each OB In OBS
70            Select Case UCase$(OB.Discipline)
              Case "DEMOGRAPHIC": DemComment = OB.Comment
80            Case "MICROCONSULTANT": ConsComment = OB.Comment
90            Case "MICROCS": CSComment = OB.Comment
100           End Select
110       Next
120       CommentsFound = True
130       rtb.SelColor = vbBlue
140       rtb.SelFontSize = 12
150       rtb.SelBold = True
160       rtb.SelText = "Comments:" & vbCrLf
170       rtb.SelColor = vbBlack
180       rtb.SelFontSize = 10
190       If Trim$(DemComment) <> "" Then
200           rtb.SelBold = True
210           rtb.SelText = "Demographic Comment: "
220           rtb.SelBold = False
230           rtb.SelText = DemComment & vbCrLf
240       End If
          'SHOW MICRO CONSULTANT AND MICRO SCIENTIST COMMENT IF SAMPLE IS VALIDATED
250       If grdSID.TextMatrix(grdSID.Row, 10) = 1 Then
260           If Trim$(ConsComment) <> "" Then
270               rtb.SelBold = True
280               rtb.SelText = "Consultant Comment: "
290               rtb.SelBold = False
300               rtb.SelText = ConsComment & vbCrLf
310           End If
320           If Trim$(CSComment) <> "" Then
330               rtb.SelBold = True
340               rtb.SelText = "Medical Scientist Comment: "
350               rtb.SelBold = False
360               rtb.SelText = CSComment & vbCrLf
370           End If
380       End If
390   End If

400   sql = "SELECT PCA0, PCA1, PCA2, PCA3 " & _
            "FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
410   Set tb = New Recordset
420   RecOpenClient 0, tb, sql
430   If Not tb.EOF Then
440       S = Trim$(tb!PCA0 & " " & tb!PCA1 & " " & tb!PCA2 & " " & tb!PCA3 & "")
450       If Trim$(S) <> "" Then
460           If Not CommentsFound Then
470               rtb.SelColor = vbBlue
480               rtb.SelFontSize = 12
490               rtb.SelBold = True
500               rtb.SelText = "Comments:" & vbCrLf
510               rtb.SelColor = vbBlack
520               rtb.SelFontSize = 10
530           End If
540           rtb.SelBold = True
550           rtb.SelText = "Current Antibiotics: "
560           rtb.SelBold = False
570           rtb.SelText = S & vbCrLf
580           CommentsFound = True
590       End If
600   End If

610   If CommentsFound Then
620       rtb.SelText = vbCrLf
630       rtb.SelText = String$(80, "-") & vbCrLf
640   End If

650   Exit Sub

FillCommentsRTB_Error:

      Dim strES As String
      Dim intEL As Integer

660   intEL = Erl
670   strES = Err.Description
680   LogError "frmMicroReport", "FillCommentsRTB", intEL, strES, sql

End Sub
Private Sub FillGenericResultsRTB(ByVal SampleID As String, _
                                  ByVal Cn As Integer)

      Dim sql As String
      Dim tb As Recordset
      Dim TestName As String

10    On Error GoTo FillGenericResultsRTB_Error

20    sql = "SELECT COALESCE(D.Valid,0) AS Valid, G.* FROM PrintValidLog AS D, GenericResults AS G WHERE " & _
            "D.SampleID = '" & Val(SampleID) + sysOptMicroOffset(Cn) & "' " & _
            "AND D.SampleID = G.SampleID"
30    Set tb = New Recordset
40    RecOpenServer Cn, tb, sql
50    If Not tb.EOF Then
60        rtb.SelText = vbCrLf
70        If Not WordResultPrinted Then
80            rtb.SelColor = vbBlue
90            rtb.SelBold = True
100           rtb.SelFontSize = 12
110           rtb.SelText = "Results:" & vbCrLf
120           WordResultPrinted = True
130       End If
140       rtb.SelColor = vbBlack
150       rtb.SelBold = False
160       rtb.SelFontSize = 10
170       Do While Not tb.EOF
180           rtb.SelColor = vbBlack
190           rtb.SelBold = False
200           rtb.SelFontSize = 10
210           TestName = tb!TestName & ""
220           If UCase$(TestName) = "REDSUB" Then
230               TestName = "Reducing Substances"
240           End If
250           rtb.SelText = TestName & " : "
260           rtb.SelBold = True
270           rtb.SelText = IIf(tb!Valid <> 0, tb!Result, "Not yet available") & vbCrLf
280           tb.MoveNext
290       Loop
300       rtb.SelText = vbCrLf
310   End If

320   Exit Sub

FillGenericResultsRTB_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmMicroReport", "FillGenericResultsRTB", intEL, strES, sql

End Sub

Private Sub PrintMicroCSF(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim S As String

10    On Error GoTo PrintMicroCSF_Error

20    sql = "SELECT * FROM CSFResults WHERE " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If Not tb.EOF Then
60        If Not WordResultPrinted Then
70            rtb.SelColor = vbBlue
80            rtb.SelBold = True
90            rtb.SelFontSize = 12
100           rtb.SelText = "Results:" & vbCrLf
110           WordResultPrinted = True
120       End If
130       rtb.SelColor = vbBlack
140       rtb.SelBold = False
150       rtb.SelFontName = "Courier New"
160       rtb.SelFontSize = 10
170       rtb.SelText = vbCrLf
180       rtb.SelFontName = "Courier New"
190       rtb.SelBold = True
200       rtb.SelText = "Appearance: "
210       rtb.SelBold = False
220       rtb.SelText = "Sample 1: " & tb!Appearance0 & vbCrLf
230       rtb.SelFontName = "Courier New"
240       rtb.SelFontSize = 10
250       rtb.SelText = "            Sample 2: " & tb!Appearance1 & vbCrLf
260       rtb.SelFontName = "Courier New"
270       rtb.SelFontSize = 10
280       rtb.SelText = "            Sample 3: " & tb!Appearance2 & vbCrLf
290       rtb.SelFontName = "Courier New"
300       rtb.SelFontSize = 10

310       rtb.SelBold = True
320       rtb.SelText = "Gram Stain: "
330       rtb.SelBold = False
340       rtb.SelText = tb!Gram & vbCrLf
350       rtb.SelFontName = "Courier New"
360       rtb.SelFontSize = 10

370       rtb.SelText = "         "
380       rtb.SelUnderline = True
390       rtb.SelText = "Sample 1"
400       rtb.SelUnderline = False
410       rtb.SelText = "        "
420       rtb.SelUnderline = True
430       rtb.SelFontSize = 10
440       rtb.SelText = "Sample 2"
450       rtb.SelUnderline = False
460       rtb.SelText = "        "
470       rtb.SelUnderline = True
480       rtb.SelFontSize = 10
490       rtb.SelText = "Sample 3" & vbCrLf
500       rtb.SelUnderline = False
510       rtb.SelFontName = "Courier New"
520       rtb.SelFontSize = 10

530       rtb.SelBold = True
540       rtb.SelText = "WCC/cmm    "
550       rtb.SelBold = False
560       S = Left$(tb!WCC0 & Space(16), 16)
570       S = S & Left$(tb!WCC1 & Space(16), 16)
580       S = S & tb!WCC2 & ""
590       rtb.SelText = S & vbCrLf
600       rtb.SelFontName = "Courier New"
610       rtb.SelFontSize = 10

620       rtb.SelBold = True
630       rtb.SelText = "RCC/cmm    "
640       rtb.SelBold = False

650       S = Left$(tb!RCC0 & Space(16), 16)
660       S = S & Left$(tb!RCC1 & Space(16), 16)
670       S = S & tb!RCC2 & ""
680       rtb.SelText = S & vbCrLf
690       rtb.SelFontName = "Courier New"
700       rtb.SelFontSize = 10

710       rtb.SelBold = True
720       rtb.SelText = "White Cell Differential: "
730       rtb.SelBold = False
740       rtb.SelText = tb!WCCdiff0 & " % Neutrophils " & tb!WCCdiff1 & "% Mononuclear Cells" & vbCrLf
750       rtb.SelFontName = "Courier New"
760       rtb.SelFontSize = 10

770   End If

780   Exit Sub

PrintMicroCSF_Error:

      Dim strES As String
      Dim intEL As Integer

790   intEL = Erl
800   strES = Err.Description
810   LogError "frmMicroReport", "PrintMicroCSF", intEL, strES, sql

End Sub

Private Sub FillFaecesRTB(ByVal SampleID As String, _
                          ByVal Cn As Integer)

      Dim n As Integer
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

10    On Error GoTo FillFaecesRTB_Error

20    Fxs.Load Val(SampleID) + sysOptMicroOffset(Cn)

30    If Fxs.Count > 0 Then
40        rtb.SelText = vbCrLf
50        If Not WordResultPrinted Then
60            rtb.SelColor = vbBlue
70            rtb.SelBold = True
80            rtb.SelFontSize = 12
90            rtb.SelText = "Results:" & vbCrLf
100           WordResultPrinted = True
110       End If
120       rtb.SelColor = vbBlack

          'Occult Blood
130       For n = 0 To 2
140           Set Fx = Fxs.Item("OB" & Format$(n))
150           If Not Fx Is Nothing Then
160               rtb.SelBold = False
170               rtb.SelText = "Occult Blood (" & n & ") : "
180               rtb.SelBold = True
190               If Fx.Valid = 0 Then
200                   rtb.SelText = "Not yet available." & vbCrLf
210               Else
220                   Select Case Fx.Result
                      Case "N": rtb.SelText = "Negative" & vbCrLf
230                   Case "P": rtb.SelText = "Positive" & vbCrLf
240                   Case Else: rtb.SelText = "Not available" & vbCrLf
250                   End Select
260               End If
270           End If
280       Next

          'Rota and Adeno
290       Set Fx = Fxs.Item("Rota")
300       If Not Fx Is Nothing Then
310           rtb.SelText = vbCrLf
320           rtb.SelBold = False
330           rtb.SelText = "Rota Virus : "
340           rtb.SelBold = True
350           If Fx.Valid = 0 Then
360               rtb.SelText = "Not yet available." & vbCrLf
370           Else
380               Select Case Fx.Result
                  Case "N": rtb.SelText = "Negative" & vbCrLf
390               Case "P": rtb.SelText = "Positive" & vbCrLf
400               End Select
410           End If
420       End If
430       Set Fx = Fxs.Item("Adeno")
440       If Not Fx Is Nothing Then
450           rtb.SelBold = False
460           rtb.SelText = vbCrLf
470           rtb.SelBold = False
480           rtb.SelText = "Adeno Virus : "
490           rtb.SelBold = True
500           If Fx.Valid = 0 Then
510               rtb.SelText = "Not yet available." & vbCrLf
520           Else
530               Select Case Fx.Result
                  Case "N": rtb.SelText = "Negative" & vbCrLf
540               Case "P": rtb.SelText = "Positive" & vbCrLf
550               End Select
560           End If
570       End If

          'C.diff
580       Set Fx = Fxs.Item("ToxinAL")
590       If Not Fx Is Nothing Then
600           rtb.SelText = vbCrLf
610           rtb.SelBold = False
620           rtb.SelText = "C. difficile : "
630           rtb.SelBold = True
640           If Fx.Valid = 0 Then
650               rtb.SelText = "Not yet available." & vbCrLf
660           Else
670               Select Case Fx.Result
                  Case "N": rtb.SelText = "Not detected" & vbCrLf
680               Case "P": rtb.SelText = "Positive" & vbCrLf
690               Case "I": rtb.SelText = "Inconclusive" & vbCrLf
700               Case "R": rtb.SelText = "Sample Rejected" & vbCrLf
710               End Select
720           End If
730       End If

          'Cryptosporidium
740       Set Fx = Fxs.Item("AUS")
750       If Not Fx Is Nothing Then
760           rtb.SelText = vbCrLf
770           rtb.SelBold = False
780           rtb.SelText = "Cryptosporidium : "
790           rtb.SelBold = True
800           If Fx.Valid = 0 Then
810               rtb.SelText = "Not yet available." & vbCrLf
820           Else
830               Select Case Fx.Result
                  Case "N": rtb.SelText = "Negative" & vbCrLf
840               Case "P": rtb.SelText = "Positive" & vbCrLf
850               End Select
860           End If
870       End If

          'Ova/Parasites
880       Set Fx = Fxs.Item("OP0")
890       If Not Fx Is Nothing Then
900           rtb.SelText = vbCrLf
910           rtb.SelBold = False
920           rtb.SelText = "Ova/Parasites(1) : "
930           rtb.SelBold = True
940           If Fx.Valid = 0 Then
950               rtb.SelText = "Not yet available." & vbCrLf
960           Else
970               rtb.SelText = Fx.Result & vbCrLf
980           End If
990       End If
1000      Set Fx = Fxs.Item("OP1")
1010      If Not Fx Is Nothing Then
1020          rtb.SelText = vbCrLf
1030          rtb.SelBold = False
1040          rtb.SelText = "Ova/Parasites(2) : "
1050          rtb.SelBold = True
1060          If Fx.Valid = 0 Then
1070              rtb.SelText = "Not yet available." & vbCrLf
1080          Else
1090              rtb.SelText = Fx.Result & vbCrLf
1100          End If
1110      End If
1120      Set Fx = Fxs.Item("OP2")
1130      If Not Fx Is Nothing Then
1140          rtb.SelText = vbCrLf
1150          rtb.SelBold = False
1160          rtb.SelText = "Ova/Parasites(3) : "
1170          rtb.SelBold = True
1180          If Fx.Valid = 0 Then
1190              rtb.SelText = "Not yet available." & vbCrLf
1200          Else
1210              rtb.SelText = Fx.Result & vbCrLf
1220          End If
1230      End If

1240  End If

1250  Exit Sub

FillFaecesRTB_Error:

      Dim strES As String
      Dim intEL As Integer

1260  intEL = Erl
1270  strES = Err.Description
1280  LogError "frmMicroReport", "FillFaecesRTB", intEL, strES

End Sub




Private Sub FillGrid()

      Dim sql As String
      Dim tb As Recordset
      Dim S As String
      Dim DaysBack As Integer
      Dim n As Integer
      Dim SearchCriteria As String

10    On Error GoTo FillGrid_Error
20    cmdSignOffMicro.Enabled = False
30    With grdSID
40        .ColWidth(5) = 0    'report number
50        .ColWidth(6) = 0    'counter
60        .ColWidth(10) = 0    'Valid
70        .Rows = 2
80        .AddItem ""
90        .RemoveItem 1
100   End With

105
    
110   With frmViewResultsWE
120       For n = 1 To .grd.Rows - 1
130           If cmbDays.ItemData(cmbDays.ListIndex) > 0 Then
140               SearchCriteria = "AND PrintTime BETWEEN '" & Format(DateAdd("d", -cmbDays.ItemData(cmbDays.ListIndex), Now), "dd/MMM/yyyy HH:mm:ss") & _
                                   "' AND '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "' "
150           End If

160           If n > 1 Then sql = sql & " UNION "

170           sql = sql & "SELECT D.SampleID, D.Age, D.Sex, D.Addr0, D.Addr1, D.RunDate, D.SampleDate, R.PrintTime, R.ReportNumber, R.Counter, " & _
                    "R.Hidden,ISNULL(R.ReportType,'') AS ReportType, COALESCE(P.Valid, 0) Valid, P.SignOff , P.SignOffBy, P.SignOffDateTime " & _
                    "FROM demographics D " & _
                    "LEFT JOIN (SELECT * FROM Reports WHERE Dept = 'Microbiology' AND COALESCE(Hidden, 0) <> 1 AND COALESCE(Hidden, 0) <> 2 " & SearchCriteria & ") R " & _
                    "ON D.SampleID = R.SampleID + " & sysOptMicroOffset(0) & " " & _
                    "LEFT JOIN PrintValidLog P " & _
                    "ON D.SampleID = P.SampleID + " & sysOptMicroOffset(0) & "  " & _
                    "WHERE D.SampleID > 2000000 AND D.SampleID < 300000000 AND D.PatName = '" & AddTicks(.grd.TextMatrix(n, 2)) & "' "
180           If Len(SearchCriteria) > 0 Then
190               sql = sql & "AND (D.RunDate BETWEEN '" & Format(DateAdd("d", -cmbDays.ItemData(cmbDays.ListIndex), Now), "dd/MMM/yyyy HH:mm:ss") & "' AND '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "' " & _
                        "OR R.PrintTime BETWEEN '" & Format(DateAdd("d", -cmbDays.ItemData(cmbDays.ListIndex), Now), "dd/MMM/yyyy HH:mm:ss") & "' AND '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "') "
200           End If
210           If Trim$(.grd.TextMatrix(n, 0)) <> "" Then
220               sql = sql & "AND D.Chart = '" & .grd.TextMatrix(n, 0) & "' "
230           Else
240               sql = sql & "AND COALESCE(D.Chart, '') = '' "
250           End If
260           If IsDate(.grd.TextMatrix(n, 1)) Then
270               sql = sql & "AND DoB = '" & Format$(.grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
280           Else
290               sql = sql & "AND COALESCE(DoB, '') = '' "
300           End If


310       Next n

320   End With

330   sql = sql & "ORDER BY ReportType DESC, PrintTime DESC"
340   Set tb = New Recordset
350   RecOpenServer 0, tb, sql
360   Do While Not tb.EOF
370       S = Format$(Val(tb!SampleID) - sysOptMicroOffset(0)) & vbTab & _
              tb!Rundate & vbTab
380       If IsDate(tb!SampleDate) Then
390           If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
400               S = S & Format(tb!SampleDate, "dd/MM/yy hh:mm")
410           Else
420               S = S & Format(tb!SampleDate, "dd/MM/yy")
430           End If
440       Else
450           S = S & "Not Specified"
460       End If
470       S = S & vbTab & _
              Format$(tb!PrintTime, "dd/MM/yy HH:nn") & vbTab & _
              LoadOutstandingMicro(tb!SampleID) & " (Pre-Printed)" & vbTab & _
              tb!ReportNumber & vbTab & _
              tb!Counter & vbTab & _
              tb!ReportType & "" & vbTab & _
              Format$(tb!SignOffDateTime, "dd/MM/yyyy hh:mm:ss") & "" & vbTab & _
              tb!SignOffBy & "" & vbTab & _
              tb!Valid
480       grdSID.AddItem S

490       lblAge = tb!Age & ""
500       Select Case Left$(UCase$(tb!Sex & ""), 1)
          Case "M": lblSex = "Male"
510       Case "F": lblSex = "Female"
520       Case Else: lblSex = ""
530       End Select
540       lblAddress = tb!Addr0 & " " & tb!Addr1 & ""
550       tb.MoveNext
560   Loop

570   If grdSID.Rows > 2 Then
580       grdSID.RemoveItem 1
590   End If
600   grdSID.col = 2
      'grdSID.Sort = 9

610   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "frmMicroReport", "FillGrid", intEL, strES, sql

End Sub

Private Sub FillGridNotReported()

      Dim sql As String
      Dim tb As Recordset
      Dim S As String
      Dim n As Integer
      Dim i As Integer

10    On Error GoTo FillGridNotReported_Error

20    With frmViewResultsWE
30        For n = 1 To .grd.Rows - 1
40            If n > 1 Then sql = sql & " UNION "

50            sql = sql & "SELECT SampleID, Age, Sex, Addr0, Addr1, RunDate, SampleDate " & _
                    "FROM Demographics " & _
                    "WHERE PatName = '" & AddTicks(lblName) & "' "
60            If Trim$(.grd.TextMatrix(n, 0)) <> "" Then
70                sql = sql & "AND Chart = '" & .grd.TextMatrix(n, 0) & "' "
80            Else
90                sql = sql & "AND ( Chart IS NULL OR Chart = '' ) "
100           End If
110           If IsDate(.grd.TextMatrix(n, 1)) Then
120               sql = sql & "AND DoB = '" & Format$(.grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
130           Else
140               sql = sql & "AND ( DoB IS NULL OR DoB = '' ) "
150           End If
160           If cmbDays.ItemData(cmbDays.ListIndex) > 0 Then
170               sql = sql & "AND RunDate BETWEEN '" & Format(DateAdd("d", -cmbDays.ItemData(cmbDays.ListIndex), Now), "dd/MMM/yyyy HH:mm:ss") & _
                        "' AND '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "' "
180           End If
190           sql = sql & "AND SampleID > '" & sysOptMicroOffsetOLD(0) & "' AND SampleID < 300000000 " & _
                    "AND SampleID NOT IN ("
200           For i = 2 To grdSID.Rows - 1
210               sql = sql & Val(grdSID.TextMatrix(i, 0)) + sysOptMicroOffset(0) & ", "
220           Next
230           sql = sql & "1)"
240       Next n
250   End With

260   Set tb = New Recordset
270   RecOpenClient 0, tb, sql

280   Do While Not tb.EOF
290       S = Format$(Val(tb!SampleID) - sysOptMicroOffset(0)) & vbTab & _
              tb!Rundate & vbTab
300       If IsDate(tb!SampleDate) Then
310           If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
320               S = S & Format(tb!SampleDate, "dd/MM/yy hh:mm")
330           Else
340               S = S & Format(tb!SampleDate, "dd/MM/yy")
350           End If
360       Else
370           S = S & "Not Specified"
380       End If
390       S = S & vbTab & "In Lab - Not ready" & vbTab
400       S = S & LoadOutstandingMicro(tb!SampleID)
410       grdSID.AddItem S


420       tb.MoveNext
430   Loop

440   Exit Sub

FillGridNotReported_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmMicroReport", "FillGridNotReported", intEL, strES, sql

End Sub

Private Sub FillGridSemen()

      Dim sql As String
      Dim tb As Recordset
      Dim S As String

10    On Error GoTo FillGridSemen_Error

20    With grdSID
30        .ColWidth(5) = 0    'report number
40        .ColWidth(6) = 0    'counter
50        .ColWidth(10) = 0    'Valid
60        .Rows = 2
70        .AddItem ""
80        .RemoveItem 1
90    End With

100   sql = "SELECT D.SampleID, D.Age, D.Sex, D.Addr0, D.Addr1, D.RunDate, D.SampleDate, R.PrintTime, R.ReportNumber, R.Counter, R.Hidden,isnull(R.ReportType,'') as ReportType " & _
            "FROM Reports R Join Demographics D " & _
            "ON D.SampleID = R.SampleID + " & sysOptSemenOffset(0) & " " & _
            "WHERE R.Dept = 'Semen' " & _
            "AND D.PatName = '" & AddTicks(lblName) & "' " & _
            "AND R.Hidden <> 1 AND R.Hidden <> 2 "
110   If Trim$(lblChart) <> "" Then
120       sql = sql & "AND D.Chart = '" & lblChart & "' "
130   Else
140       sql = sql & "AND COALESCE(Chart, '') = '' "
150   End If
160   If IsDate(lblDoB) Then
170       sql = sql & "AND DoB = '" & Format$(lblDoB, "dd/mmm/yyyy") & "' "
180   Else
190       sql = sql & "AND COALESCE(DoB, '') = '' "
200   End If
210   If cmbDays.ItemData(cmbDays.ListIndex) > 0 Then
220       sql = sql & "AND PrintTime BETWEEN '" & Format(DateAdd("d", -cmbDays.ItemData(cmbDays.ListIndex), Now), "dd/MMM/yyyy") & _
                "' AND '" & Format(Now, "dd/MMM/yyyy") & "' "
230   End If
240   sql = sql & " and ISNULL(R.ReportType,'') <> 'Interim Report' "
250   sql = sql & "ORDER BY PrintTime DESC"

260   Set tb = New Recordset
270   RecOpenServer 0, tb, sql
280   Do While Not tb.EOF
290       S = Format$(Val(tb!SampleID) - sysOptSemenOffset(0)) & vbTab & _
              tb!Rundate & vbTab
300       If IsDate(tb!SampleDate) Then
310           If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
320               S = S & Format(tb!SampleDate, "dd/MM/yy hh:mm")
330           Else
340               S = S & Format(tb!SampleDate, "dd/MM/yy")
350           End If
360       Else
370           S = S & "Not Specified"
380       End If
390       S = S & vbTab & _
              Format$(tb!PrintTime, "dd/MM/yy HH:nn") & vbTab & _
              "Semen Analysis" & " (Pre-Printed)" & vbTab & _
              tb!ReportNumber & vbTab & _
              tb!Counter & vbTab & _
              tb!ReportType
400       grdSID.AddItem S
410       lblAge = tb!Age & ""
420       Select Case Left$(UCase$(tb!Sex & ""), 1)
              Case "M": lblSex = "Male"
430           Case "F": lblSex = "Female"
440           Case Else: lblSex = ""
450       End Select
460       lblAddress = tb!Addr0 & " " & tb!Addr1 & ""
470       tb.MoveNext
480   Loop








490   sql = "SELECT D.SampleID, D.Age, D.Sex, D.Addr0, D.Addr1, D.RunDate, D.SampleDate, P.PrintedDateTime from Demographics D, PrintValidLog P WHERE " & _
            "PatName = '" & AddTicks(lblName) & "' "
500   If Trim$(lblChart) <> "" Then
510       sql = sql & "AND Chart = '" & lblChart & "' "
520   Else
530       sql = sql & "AND COALESCE(Chart, '') = '' "
540   End If
550   If IsDate(lblDoB) Then
560       sql = sql & "AND DoB = '" & Format$(lblDoB, "dd/mmm/yyyy") & "' "
570   Else
580       sql = sql & "AND COALESCE(DoB, '') = '' "
590   End If
600   sql = sql & "AND D.SampleID = P.SampleID " & _
            "AND D.SampleID > '" & sysOptSemenOffset(0) & "' " & _
            "AND D.SampleID < '" & sysOptMicroOffsetOLD(0) & "' " & _
            "AND SampleDate < '" & GetOptionSetting("WardEnqV7Date", "01/May/2011", "") & "'"


610   Set tb = New Recordset
620   RecOpenClient 0, tb, sql

630   Do While Not tb.EOF
640       S = Format$(Val(tb!SampleID) - sysOptSemenOffset(0)) & vbTab & _
              tb!Rundate & vbTab
650       If IsDate(tb!SampleDate) Then
660           If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
670               S = S & Format(tb!SampleDate, "dd/MM/yy hh:mm")
680           Else
690               S = S & Format(tb!SampleDate, "dd/MM/yy")
700           End If
710       Else
720           S = S & "Not Specified"
730       End If
740       S = S & vbTab
750       If Not IsNull(tb!PrintedDateTime) Then
760           S = S & Format$(tb!PrintedDateTime, "dd/MM/yy HH:mm")
770       End If
780       S = S & vbTab
790       S = S & "Semen Analysis"
800       S = S & vbTab & HospName(0)
810       grdSID.AddItem S

820       lblAge = tb!Age & ""
830       Select Case Left$(UCase$(tb!Sex & ""), 1)
              Case "M": lblSex = "Male"
840           Case "F": lblSex = "Female"
850           Case Else: lblSex = ""
860       End Select
870       lblAddress = tb!Addr0 & " " & tb!Addr1 & ""
880       tb.MoveNext
890   Loop

900   If grdSID.Rows > 2 Then
910       grdSID.RemoveItem 1
920   End If
930   grdSID.col = 1
940   grdSID.Sort = 9

950   Exit Sub

FillGridSemen_Error:

      Dim strES As String
      Dim intEL As Integer

960   intEL = Erl
970   strES = Err.Description
980   LogError "frmMicroReport", "FillGridSemen", intEL, strES, sql


End Sub

Private Function LoadOutstandingMicro(ByVal SampleIDWithOffset As String) As String

      Dim sql As String
      Dim S As String
      Dim UR As UrineRequest
      Dim URs As New UrineRequests
      Dim fr As FaecesRequest
      Dim FRs As New FaecesRequests
      Dim SDS As New SiteDetails

10    On Error GoTo LoadOutstandingMicro_Error

20    SDS.Load SampleIDWithOffset
30    If SDS.Count > 0 Then
40        S = SDS(1).Site & " " & SDS(1).SiteDetails & " "
50        If SDS(1).Site = "Faeces" Then
60            FRs.Load SampleIDWithOffset
70            For Each fr In FRs
80                S = S & fr.Request & " "
90            Next
100       ElseIf SDS(1).Site = "Urine" Then
110           URs.Load SampleIDWithOffset
120           For Each UR In URs
130               S = S & UR.Request & " "
140           Next
150       End If
160   End If
170   LoadOutstandingMicro = S

180   Exit Function

LoadOutstandingMicro_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmMicroReport", "LoadOutstandingMicro", intEL, strES, sql

End Function
Private Sub RePrint()

      Dim sql As String

10    On Error GoTo RePrint_Error

20    If iMsg("Report will be printed on" & vbCrLf & _
              WardEnqForcedPrinter & "." & vbCrLf & _
              "OK?", vbQuestion + vbYesNo) = vbYes Then

30        sql = "IF NOT EXISTS(SELECT * FROM PrintPending WHERE " & _
                "              SampleID = '" & grdSID.TextMatrix(grdSID.Row, 0) & "' " & _
                "              AND Department = 'M' ) " & _
                "  INSERT INTO PrintPending " & _
                "  (SampleID, Department, Initiator, UsePrinter, FaxNumber, ptime, UseConnection, " & _
                "  Hyear, Ward, Clinician, GP, PrintOnCondition, DateTimeOfRecord, ThisIsCopy, ReprintReportCounter ) " & _
                "  VALUES " & _
                "  ('" & grdSID.TextMatrix(grdSID.Row, 0) & "', " & _
                "  'M', " & _
                "  '" & AddTicks(UserName) & "', " & _
                "  '" & WardEnqForcedPrinter & "', " & _
                "  '', " & _
                "  getdate(), " & _
                "  '0', '', '', '', '', 0, getdate(), 1, " & _
                "  '" & grdSID.TextMatrix(grdSID.Row, 6) & "')"
40        Cnxn(0).Execute sql

50        LogAsViewed "N", grdSID.TextMatrix(grdSID.Row, 0), lblChart
60    End If

70    Exit Sub

RePrint_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmMicroReport", "RePrint", intEL, strES, sql

End Sub

Private Sub cmbDays_Click()

10    On Error GoTo cmbDays_Click_Error

20    If Activated Then

30        If ReportDept = "SEMEN" Then
              'this report is complete. if there is a need to work on thi
              'in the future. then it can be started from here.
              'SEMEN ANALYSIS REPORT BUTTON IS PERMANETLY DISABLED.
          
40            FillGridSemen
50        ElseIf ReportDept = "MICRO" Then
          
60            FillGrid
70        End If
80    End If
90    Exit Sub

cmbDays_Click_Error:

       Dim strES As String
       Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "frmMicroReport", "cmbDays_Click", intEL, strES
          
End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim Ward As String
      Dim Clinician As String
      Dim SID As Long

10    On Error GoTo cmdPrint_Click_Error
20    If grdSID.TextMatrix(grdSID.Row, 0) = "" Then Exit Sub

30    If grdSID.TextMatrix(grdSID.Row, 7) = "" Then Exit Sub

40    SID = Val(grdSID.TextMatrix(grdSID.Row, 0))
50    If grdSID.TextMatrix(grdSID.Row, 6) <> "" Then
60        RePrint
70        Exit Sub
80    End If

90    If InStr(grdSID.TextMatrix(grdSID.Row, 3), "Not") <> 0 Then
100       iMsg "Cannot Print." & vbCrLf & "Report not ready", vbOKOnly
110       Exit Sub
120   End If

130   sql = "Select Ward, Clinician from Demographics where " & _
            "SampleID = '" & SID + sysOptMicroOffset(0) & "'"
140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql
160   If tb.EOF Then
170       Exit Sub
180   End If

190   Ward = tb!Ward & ""
200   Clinician = tb!Clinician & ""

210   If iMsg("Report will be printed on" & vbCrLf & _
              WardEnqForcedPrinter & "." & vbCrLf & _
              "OK?", vbQuestion + vbYesNo) = vbYes Then

220       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'M' " & _
                "AND SampleID = '" & SID & "'"
230       Set tb = New Recordset
240       RecOpenClient 0, tb, sql
250       If tb.EOF Then
260           tb.AddNew
270       End If
280       tb!SampleID = SID
290       tb!Ward = Ward
300       tb!Clinician = Clinician
310       tb!GP = ""
320       tb!Department = "M"
330       tb!Initiator = UserName
340       tb!UsePrinter = WardEnqForcedPrinter
350       tb!ThisIsCopy = 1
360       tb.Update

370       LogAsViewed "N", SID, lblChart

380   End If

390   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "frmMicroReport", "cmdPrint_Click", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSignOffMicro_Click
' Author    : Masood
' Date      : 23/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSignOffMicro_Click()

10    On Error GoTo cmdSignOffMicro_Click_Error

      Dim sql        As String

20    sql = " UPDATE printvalidlog SET SignOff = 1 , "
30    sql = sql & " SignOffBy = '" & UserName & "'"
40    sql = sql & " , SignOffDateTime ='" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "'"
50    sql = sql & " WHERE SampleID = '" & SampleIDForSignOff & "'"

60    Cnxn(0).Execute sql

70    cmdSignOffMicro.Enabled = False
80    grdSID.TextMatrix(grdSID.Row, 8) = Format$(Now, "dd/MM/yyyy hh:mm:ss")
90    grdSID.TextMatrix(grdSID.Row, 9) = UserName


100   Exit Sub


cmdSignOffMicro_Click_Error:

      Dim strES      As String
      Dim intEL      As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmMicroReport", "cmdSignOffMicro_Click", intEL, strES, sql
End Sub

Private Sub cmdViewScan_Click()
10    frmViewScan.CallerDepartment = "WardEnq Micro"
20    frmViewScan.SampleID = grdSID.TextMatrix(grdSID.Row, 0)
30    frmViewScan.txtSampleID = grdSID.TextMatrix(grdSID.Row, 0)
40    frmViewScan.Show 1
End Sub

Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    PBar.Max = LogOffDelaySecs
30    PBar = 0

40    SingleUserUpdateLoggedOn UserName

50    Timer1.Enabled = True

60    If Activated Then Exit Sub

70    cmbDays.ListIndex = 1
80    Activated = True


90    If ReportDept = "SEMEN" Then
          'this report is complete. if there is a need to work on thi
          'in the future. then it can be started from here.
          'SEMEN ANALYSIS REPORT BUTTON IS PERMANETLY DISABLED.

100       FillGridSemen
110   ElseIf ReportDept = "MICRO" Then

120       FillGrid
130   End If

140   cmdPrint.Visible = UserCanPrint

150   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmMicroReport", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub


Private Sub Form_Load()

10    Activated = False

20    PBar.Max = LogOffDelaySecs
30    PBar = 0

40    SortOrder = True


End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub grdSID_Click()

      Dim x As Integer
      Dim y As Integer

10    On Error GoTo grdSID_Click_Error
20    If grdSID.TextMatrix(1, 0) = "" Then Exit Sub

30    lblClDetails = ""
40    cmdSignOffMicro.Enabled = False
50    rtb.Text = ""
60    If grdSID.MouseRow = 0 Then
70        If InStr(grdSID.TextMatrix(0, grdSID.col), "Date") Then
80            grdSID.Sort = 9
90        Else
100           If SortOrder Then
110               grdSID.Sort = flexSortGenericAscending
120           Else
130               grdSID.Sort = flexSortGenericDescending
140           End If
150       End If
160       SortOrder = Not SortOrder
170       Exit Sub
180   End If

190   For y = 1 To grdSID.Rows - 1
200       grdSID.Row = y
210       For x = 1 To grdSID.Cols - 1
220           grdSID.col = x
230           grdSID.CellBackColor = 0
240       Next
250   Next

260   grdSID.Row = grdSID.MouseRow
270   For x = 1 To grdSID.Cols - 1
280       grdSID.col = x
290       grdSID.CellBackColor = vbYellow
300   Next

310   If grdSID.TextMatrix(grdSID.Row, 3) = "In Lab - Not ready" Then
          '  FillCommentsRTB grdSID.TextMatrix(grdSID.Row, 0)
          '  FillResultMicroRTB grdSID.TextMatrix(grdSID.Row, 0), 0
320       rtb.SelColor = vbBlue
330       rtb.SelBold = True
340       rtb.SelFontSize = 12
350       rtb.SelText = vbCrLf & "Sample in Laboratory - Report not ready."

360   Else
370       If grdSID.TextMatrix(grdSID.Row, 6) = "" Then
              'AddActivity grdSID.TextMatrix(grdSID.Row, 0), "Ward Enquiry Micro Report", "VIEWED", "", lblChart, "", ""
              'FillCommentsRTB grdSID.TextMatrix(grdSID.Row, 0)
              'FillResultMicroRTB grdSID.TextMatrix(grdSID.Row, 0), 0
380       Else
390           AddActivity grdSID.TextMatrix(grdSID.Row, 0), "Ward Enquiry Micro Report", "VIEWED", "", lblChart, "", ""
400           FillReport

410       End If
420   End If

430   cmdPrint.Visible = UserCanPrint
440   SetViewScans grdSID.TextMatrix(grdSID.Row, 0), cmdViewScan

450   Exit Sub

grdSID_Click_Error:

      Dim strES As String
      Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "frmMicroReport", "grdSID_Click", intEL, strES

End Sub
Private Sub FillReport()

      Dim tb As Recordset
      Dim sql As String

10    rtb = ""
20    rtb.SelText = ""

30    sql = "SELECT Report FROM Reports WHERE " & _
            "Counter = '" & grdSID.TextMatrix(grdSID.Row, 6) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        If Trim(tb!Report & "") <> "" Then
80            rtb.SelText = Trim(tb!Report)
90        End If
100   End If

110       SignOFFChecking (grdSID.TextMatrix(grdSID.Row, 0))

        
120   Exit Sub

FillReport_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmMicroReport", "FillReport", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SignOFFChecking
' Author    : Masood
' Date      : 19/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SignOFFChecking(SampleID As String)
10        On Error GoTo SignOFFChecking_Error
          Dim tb As Recordset
          Dim sql As String
20        SampleIDForSignOff = ""
30        sql = "SELECT ISNULL(SignOff,0) AS SignOff FROM PrintValidLog WHERE VALID = 1 AND " & _
                "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' "
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            If tb!SignOff = 0 Then
80                cmdSignOffMicro.Enabled = True
90                SampleIDForSignOff = Val(SampleID) + sysOptMicroOffset(0)
100           End If
110       End If



120       Exit Sub


SignOFFChecking_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmMicroReport", "SignOFFChecking", intEL, strES, sql
End Sub

Private Sub FillResultMicroRTB(ByVal SampleID As String, _
                               ByVal Cn As Integer)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillResultMicroRTB_Error

20    lblClDetails = ""

30    sql = "Select ClDetails from Demographics where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(Cn) & "'"
40    Set tb = New Recordset
50    RecOpenClient Cn, tb, sql
60    If Not tb.EOF Then
70        lblClDetails = tb!cldetails & ""
80    End If

90    WordResultPrinted = False

100   FillMicroscopyRTB SampleID, Cn

110   FillGramStainRTB SampleID

120   FillGenericResultsRTB SampleID, Cn

130   PrintMicroCSF grdSID.TextMatrix(grdSID.Row, 0)

140   FillFaecesRTB SampleID, Cn

150   FillMicroUrineComment SampleID, Cn

160   FillSensitivitiesRTB SampleID

170   Exit Sub

FillResultMicroRTB_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmMicroReport", "FillResultMicroRTB", intEL, strES, sql


End Sub
Private Sub FillGramStainRTB(ByVal SampleID As String)

      Dim objIdent As New IdentResult
      Dim IDs As IdentResults
      Dim Title As String

10    On Error GoTo FillGramStainRTB_Error

20    Set IDs = objIdent.LoadIdentList(SampleID, "GramStain")
30    If Not IDs Is Nothing Then
40        Title = "Gram Stain: "
50        For Each objIdent In IDs
60            rtb.SelFontName = "Courier New"
70            rtb.SelColor = vbBlack
80            rtb.SelBold = False
90            rtb.SelFontSize = 10
100           rtb.SelText = Title & objIdent.TestName & " " & objIdent.Result & vbCrLf
110           Title = "            "
120       Next
130   End If

140   Set IDs = objIdent.LoadIdentList(SampleID, "WetPrep")
150   If Not IDs Is Nothing Then
160       Title = "  Wet Prep: "
170       For Each objIdent In IDs
180           rtb.SelFontName = "Courier New"
190           rtb.SelColor = vbBlack
200           rtb.SelBold = False
210           rtb.SelFontSize = 10
220           rtb.SelText = Title & objIdent.TestName & " " & objIdent.Result & vbCrLf
230           Title = "            "
240       Next
250   End If

260   Exit Sub

FillGramStainRTB_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmMicroReport", "FillGramStainRTB", intEL, strES

End Sub

Private Sub FillMicroscopyRTB(ByVal SampleID As String, _
                              ByVal Cn As Integer)

      Dim Ux As UrineResult
      Dim Uxs As New UrineResults
      Dim R As String

10    On Error GoTo FillMicroscopyRTB_Error

20    Uxs.Load Val(SampleID) + sysOptMicroOffset(Cn)
30    If Uxs.Count = 0 Then Exit Sub

40    If Not WordResultPrinted Then
50        rtb.SelColor = vbBlue
60        rtb.SelBold = True
70        rtb.SelFontSize = 12
80        rtb.SelText = "Results:" & vbCrLf
90        WordResultPrinted = True
100   End If
110   rtb.SelColor = vbBlack
120   rtb.SelBold = False
130   rtb.SelFontSize = 10
140   rtb.SelText = "Microscopy:" & vbCrLf

150   rtb.SelFontName = "Courier New"
160   rtb.SelBold = False
170   rtb.SelText = "   Bacteria: "
180   rtb.SelBold = True
190   Set Ux = Uxs.Item("Bacteria")
200   If Ux Is Nothing Then R = "" Else R = Ux.Result
210   rtb.SelText = Left$(Trim$(R) & Space(20), 20)

220   rtb.SelBold = False
230   rtb.SelText = "Crystals:"
240   rtb.SelBold = True
250   Set Ux = Uxs.Item("Crystals")
260   If Ux Is Nothing Then R = "" Else R = Ux.Result
270   rtb.SelText = Trim$(R) & vbCrLf

280   rtb.SelFontName = "Courier New"
290   rtb.SelBold = False
300   rtb.SelText = "        WCC: "
310   rtb.SelBold = True
320   Set Ux = Uxs.Item("WCC")
330   If Ux Is Nothing Then R = "" Else R = Ux.Result
340   rtb.SelText = Left$(Trim$(R) & Space(20), 20)

350   rtb.SelBold = False
360   rtb.SelText = "   Casts:"
370   rtb.SelBold = True
380   Set Ux = Uxs.Item("Casts")
390   If Ux Is Nothing Then R = "" Else R = Ux.Result
400   rtb.SelText = Trim$(R) & vbCrLf

410   rtb.SelFontName = "Courier New"
420   rtb.SelBold = False
430   rtb.SelText = "        RCC: "
440   rtb.SelBold = True
450   Set Ux = Uxs.Item("RCC")
460   If Ux Is Nothing Then R = "" Else R = Ux.Result
470   rtb.SelText = Left$(Trim$(R) & Space(20), 20)

480   rtb.SelBold = False
490   rtb.SelText = "    Misc: "
500   rtb.SelBold = True
510   Set Ux = Uxs.Item("Misc0")
520   If Ux Is Nothing Then R = "" Else R = Ux.Result
530   rtb.SelText = Trim$(R)
540   Set Ux = Uxs.Item("Misc1")
550   If Ux Is Nothing Then R = "" Else R = Ux.Result
560   rtb.SelText = " " & Trim$(R)
570   Set Ux = Uxs.Item("Misc2")
580   If Ux Is Nothing Then R = "" Else R = Ux.Result
590   rtb.SelText = " " & Trim$(R) & vbCrLf

600   rtb.SelFontName = "MS Sans Serif"
610   rtb.SelText = vbCrLf

620   Exit Sub

FillMicroscopyRTB_Error:

      Dim strES As String
      Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "frmMicroReport", "FillMicroscopyRTB", intEL, strES

End Sub

Private Sub FillSensitivitiesRTB(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim strGroup(1 To 8) As OrgGroup
      Dim SampleIDWithOffset As Long
      Dim SensPrintMax As Integer
      Dim Site As String
      Dim x As Integer

10    On Error GoTo FillSensitivitiesRTB_Error

20    SampleIDWithOffset = Val(SampleID) + sysOptMicroOffset(0)

30    SensPrintMax = 3
40    sql = "Select Site from MicroSiteDetails where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        Site = Trim$(tb!Site & "")
90        If Site <> "" Then
100           sql = "Select [Default] as D from Lists where " & _
                    "ListType = 'SI' " & _
                    "and Text = '" & Site & "'"
110           Set tb = New Recordset
120           RecOpenServer 0, tb, sql
130           If Not tb.EOF Then
140               SensPrintMax = Val(tb!D & "")
150           End If
160       End If
170   End If

180   FillOrgGroups strGroup(), SampleIDWithOffset
190   sql = "SELECT COALESCE(D.Valid, 0) AS Valid, I.* FROM PrintValidLog AS D, Isolates AS I WHERE " & _
            "D.SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' " & _
            "AND D.SampleID = I.SampleID " & _
            "ORDER BY I.IsolateNumber"
200   Set tb = New Recordset
210   RecOpenServer 0, tb, sql
220   If Not tb.EOF Then
230       If Not WordResultPrinted Then
240           rtb.SelColor = vbBlue
250           rtb.SelBold = True
260           rtb.SelFontSize = 12
270           rtb.SelText = "Results:" & vbCrLf
280           WordResultPrinted = True
290       End If
300       rtb.SelColor = vbBlack
310       rtb.SelBold = False
320       rtb.SelFontSize = 10
330       rtb.SelText = vbCrLf
340       rtb.SelFontSize = 10
350       If tb!Valid = 0 And (Trim$(tb!OrganismGroup & "") <> "" Or Trim$(tb!OrganismGroup & "") <> "") Then
360           rtb.SelText = "Isolates not yet available." & vbCrLf
370       Else
380           sql = "UPDATE Sensitivities SET Valid = 1 WHERE SampleID = '" & SampleIDWithOffset & "' "
390           Cnxn(0).Execute sql
400       End If
410   End If

      Dim sx As Sensitivity
      Dim Sxs As New Sensitivities
420   Sxs.Load SampleIDWithOffset
430   For x = 1 To 4
440       For Each sx In Sxs
450           If strGroup(x).OrgName <> "" Then

460               rtb.SelFontSize = 10
470               rtb.SelBold = True
480               rtb.SelText = "Culture " & x & " : " & _
                                strGroup(x).OrgName & " " & _
                                strGroup(x).Qualifier & vbCrLf

490               FillSensitivityResults Sxs, x, 1

                  'Not for reporting
500               If Not (InStr(UCase$(App.Path), "WARD") > 0) Then
510                   If AnyNotForReporting(Sxs, x) Then
520                       rtb.SelColor = vbBlack
530                       rtb.SelBold = False
540                       rtb.SelFontSize = 10
550                       rtb.SelText = vbCrLf
560                       rtb.SelFontSize = 10
570                       rtb.SelText = "Sensitivities not for reporting:-" & vbCrLf
580                       rtb.SelFontName = "MS Sans Serif"
590                       rtb.SelColor = vbBlack
600                       rtb.SelText = vbCrLf
610                       FillSensitivityResults Sxs, x, False
620                   End If
630               End If
640               rtb.SelText = vbCrLf
650               Exit For
660           End If
670       Next
680   Next

690   Exit Sub

FillSensitivitiesRTB_Error:

      Dim strES As String
      Dim intEL As Integer

700   intEL = Erl
710   strES = Err.Description
720   LogError "frmMicroReport", "FillSensitivitiesRTB", intEL, strES, sql

End Sub

Private Function AnyNotForReporting(ByVal Sxs As Sensitivities, ByVal IsolateNumber As Integer) As Boolean

      Dim RetVal As Boolean
      Dim sx As Sensitivity

10    RetVal = False

20    For Each sx In Sxs
30        If sx.IsolateNumber = IsolateNumber And sx.Report = False Then
40            RetVal = True
50            Exit For
60        End If
70    Next

80    AnyNotForReporting = RetVal

End Function


Private Sub FillSensitivityResults(ByVal Sxs As Sensitivities, ByVal IsolateNumber As Integer, ByVal Report As Integer)

      Dim sx As Sensitivity

10    On Error GoTo FillSensitivityResults_Error

20    For Each sx In Sxs
30        If sx.IsolateNumber = IsolateNumber And sx.Report = Report Then
              'Only report if R/S/I not for X
40            If Trim$(sx.RSI) <> "" And sx.RSI <> "X" Then
50                rtb.SelFontName = "Courier New"
60                rtb.SelColor = vbBlack
70                rtb.SelText = Left$(sx.AntibioticName & Space$(20), 20)
80                If sx.RSI = "R" Then
90                    rtb.SelColor = vbRed
100                   rtb.SelFontName = "Courier New"
110                   rtb.SelText = "Resistant" & vbCrLf
120               ElseIf sx.RSI = "S" Then
130                   rtb.SelColor = vbGreen
140                   rtb.SelFontName = "Courier New"
150                   rtb.SelText = "Sensitive" & vbCrLf
160               ElseIf sx.RSI = "I" Then
170                   rtb.SelColor = vbBlue
180                   rtb.SelFontName = "Courier New"
190                   rtb.SelText = "Intermediate" & vbCrLf
200               End If
210           End If
220       End If
230   Next

240   Exit Sub

FillSensitivityResults_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmMicroReport", "FillSensitivityResults", intEL, strES

End Sub


Private Sub FillMicroUrineComment(ByVal SampleID As String, _
                                  ByVal Cn As Integer)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillMicroUrineComment_Error

20    sql = "Select Site from MicroSiteDetails where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(Cn) & "' " & _
            "AND Site like 'Urine'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    sql = "Select * from Isolates where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(Cn) & "' " & _
            "AND OrganismGroup <> 'Negative results' " & _
            "AND OrganismName <> ''"
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If tb.EOF Then Exit Sub

100   rtb.SelText = vbCrLf
110   rtb.SelColor = vbBlack
120   rtb.SelBold = False
130   rtb.SelFontSize = 10

140   rtb.SelText = "Positive cultures "
150   rtb.SelUnderline = True
160   rtb.SelText = "must"
170   rtb.SelUnderline = False
180   rtb.SelText = " be correlated with signs and symptoms of UTI "
190   rtb.SelText = "Particularly with low colony counts"
200   rtb.SelText = vbCrLf

210   Exit Sub

FillMicroUrineComment_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmMicroReport", "FillMicroUrineComment", intEL, strES, sql

End Sub

Private Function FillOrgGroups(ByRef strGroup() As OrgGroup, _
                               ByVal SampleIDWithOffset As Long) _
                               As Integer

      Dim tb As Recordset
      Dim tbO As Recordset
      Dim sql As String
      Dim n As Integer
      Dim IsoNum As Integer

10    On Error GoTo FillOrgGroups_Error

20    sql = "Select OrganismGroup, OrganismName, Qualifier, IsolateNumber " & _
            "from Isolates where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    n = 1
60    Do While Not tb.EOF
70        IsoNum = tb!IsolateNumber
80        With strGroup(IsoNum)
90            .OrgGroup = tb!OrganismGroup & ""
100           If Trim$(tb!OrganismName & "") = "" Then
110               .OrgName = .OrgGroup
120           Else
130               .OrgName = tb!OrganismName & ""
140           End If
150           .Qualifier = tb!Qualifier & ""
160           sql = "Select ShortName, ReportName from Organisms where " & _
                    "Name = '" & tb!OrganismName & "'"
170           Set tbO = New Recordset
180           RecOpenClient 0, tbO, sql
190           If Not tbO.EOF Then
200               .ShortName = tbO!ShortName & ""
210               .ReportName = Trim$(tbO!ReportName & "")
220           Else
230               .ShortName = .OrgName
240               .ReportName = .OrgName
250           End If
260           If .ReportName = "" Then
270               .ShortName = .OrgName
280               .ReportName = .OrgName
290               .OrgName = .OrgName
300           End If
310       End With
320       n = n + 1
330       tb.MoveNext
340   Loop

350   FillOrgGroups = n - 1

360   Exit Function

FillOrgGroups_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "frmMicroReport", "FillOrgGroups", intEL, strES, sql


End Function

Private Sub grdSID_Compare(ByVal Row1 As Long, ByVal Row2 As Long, cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(grdSID.TextMatrix(Row1, grdSID.col)) Then
20        cmp = 0
30        Exit Sub
40    End If

50    If Not IsDate(grdSID.TextMatrix(Row2, grdSID.col)) Then
60        cmp = 0
70        Exit Sub
80    End If

90    d1 = Format(grdSID.TextMatrix(Row1, grdSID.col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(grdSID.TextMatrix(Row2, grdSID.col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120       cmp = Sgn(DateDiff("s", d1, d2))
130   Else
140       cmp = Sgn(DateDiff("s", d2, d1))
150   End If


End Sub

Private Sub grdSID_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblAge_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblClDetails_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblDemogComment_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblSex_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub rtb_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10    rtb.SelLength = 0
End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1

20    If PBar = PBar.Max Then
30        Unload Me
40    End If

End Sub


