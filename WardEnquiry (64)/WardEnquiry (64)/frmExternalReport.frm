VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmExternalReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - External Tests"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "frmExternalReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   90
      TabIndex        =   5
      Top             =   270
      Width           =   10035
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
         Left            =   750
         TabIndex        =   19
         Top             =   840
         Width           =   8505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   3045
         TabIndex        =   18
         Top             =   540
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   8220
         TabIndex        =   17
         Top             =   210
         Width           =   285
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3375
         TabIndex        =   16
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8535
         TabIndex        =   15
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   4515
         TabIndex        =   14
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   5730
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   12
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   11
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5115
         TabIndex        =   10
         Top             =   510
         Width           =   4140
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6645
         TabIndex        =   9
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1665
         TabIndex        =   8
         Top             =   540
         Width           =   1200
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
         Left            =   1665
         TabIndex        =   7
         Top             =   210
         Width           =   3540
      End
      Begin VB.Label lblClDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   750
         TabIndex        =   6
         Top             =   1110
         Width           =   8505
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   885
      Left            =   9450
      Picture         =   "frmExternalReport.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   705
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9000
      Top             =   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report"
      Height          =   3705
      Left            =   90
      TabIndex        =   3
      Top             =   3780
      Width           =   10035
      Begin RichTextLib.RichTextBox rtb 
         Height          =   3345
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   5900
         _Version        =   393217
         TextRTF         =   $"frmExternalReport.frx":1534
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   885
      Left            =   8640
      Picture         =   "frmExternalReport.frx":15B6
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "bprint"
      Top             =   2280
      Width           =   705
   End
   Begin VB.CommandButton cmdFAX 
      Caption         =   "FAX"
      Enabled         =   0   'False
      Height          =   885
      Left            =   7800
      Picture         =   "frmExternalReport.frx":1C20
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   1935
      Left            =   90
      TabIndex        =   2
      Top             =   1740
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmExternalReport.frx":2062
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   60
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmExternalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean


Private pWard As String
Private pClinician As String
Private pGP As String

Private Sub FillGrid()

      Dim sql As String
      Dim tb As Recordset
      Dim S As String

10    On Error GoTo FillGrid_Error

20    With grdSID
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60    End With

70    sql = "SELECT D.SampleID, Chart, PatName, DoB, Age, Sex, Addr0, Addr1, RunDate, SampleDate FROM Demographics D, ExtResults E WHERE "
80    If Trim$(lblChart) <> "" Then
90      sql = sql & "Chart = '" & lblChart & "' AND "
100   End If
110   sql = sql & "PatName = '" & AddTicks(lblName) & "' "
120   If IsDate(lblDoB) Then
130     sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
140   End If
150   sql = sql & "AND D.SampleID = E.SampleID "
160   sql = sql & "GROUP BY D.SampleID, Chart, PatName, DoB, Age, Sex, Addr0, Addr1, RunDate, SampleDate "
170   sql = sql & "ORDER BY D.SampleID DESC"
  
180   Set tb = New Recordset
190   RecOpenClient 0, tb, sql

200   Do While Not tb.EOF
210     S = Format$(tb!SampleID) & vbTab & _
            tb!RunDate & vbTab
220     If IsDate(tb!SampleDate) Then
230       If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
240         S = S & Format(tb!SampleDate, "dd/MM/yy hh:mm")
250       Else
260         S = S & Format(tb!SampleDate, "dd/MM/yy")
270       End If
280     Else
290       S = S & "Not Specified"
300     End If
310     If Not IsNull(tb!DoB) Then
320       lblDoB = tb!DoB
330     Else
340       lblDoB = ""
350     End If
360     lblAge = tb!Age & ""
370     Select Case Left$(UCase$(tb!Sex & ""), 1)
          Case "M": lblSex = "Male"
380       Case "F": lblSex = "Female"
390       Case Else: lblSex = ""
400     End Select
410     lblAddress = tb!Addr0 & " " & tb!Addr1 & ""
420     grdSID.AddItem S
430     tb.MoveNext
440   Loop

450   If grdSID.Rows > 2 Then
460     grdSID.RemoveItem 1
470   End If

480   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

490   intEL = Erl
500   strES = Err.Description
510   LogError "frmExternalReport", "FillGrid", intEL, strES, sql


End Sub

Private Sub FillReport(ByVal SampleID As String)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillReport_Error

20    sql = "Select * from ExtResults where " & _
            "SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    Do While Not tb.EOF

60      With rtb
70        .SelIndent = 0
80        .SelColor = vbBlue
          '.SelBold = False
          '.SelText = "Analyte: "
90        .SelBold = True
100       .SelText = .SelText & tb!Analyte & ": "
110       .SelColor = vbBlack
120       .SelBold = True
130       .SelIndent = 200
140       If Trim$(tb!Result & "") <> "" Then
150         .SelText = .SelText & tb!Result & ""
160       Else
170         .SelText = .SelText & "Not yet Available."
180       End If
190       .SelBold = False
200       .SelText = .SelText & " (Sent To " & tb!SendTo & " )"
210       .SelText = .SelText & vbCrLf
220     End With

230     tb.MoveNext
240   Loop

250   Exit Sub

FillReport_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmExternalReport", "FillReport", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdFAX_Click()
                                                                                                                              '
                                                                                                                              'Dim sql As String
                                                                                                                              'Dim tb As Recordset
                                                                                                                              'Dim SampleID As String
                                                                                                                              'Dim FaxNumber As String
                                                                                                                              '
                                                                                                                              'SampleID = grdSID.TextMatrix(grdSID.Row, 0)
                                                                                                                              '
                                                                                                                              'sql = "Select * from PrintPending where " & _
                                                                                                                              '      "Department = 'M' " & _
                                                                                                                              '      "and SampleID = '" & SampleID & "' " & _
                                                                                                                              '      "and UsePrinter = 'FAX'"
                                                                                                                              'Set tb = New Recordset
                                                                                                                              'RecOpenClient 0, tb, sql
                                                                                                                              'If tb.EOF Then
                                                                                                                              '  tb.AddNew
                                                                                                                              'End If
                                                                                                                              'tb!SampleID = SampleID
                                                                                                                              'tb!Ward = pWard
                                                                                                                              'tb!Clinician = pClinician
                                                                                                                              'tb!GP = pGP
                                                                                                                              'tb!UsePrinter = "FAX"
                                                                                                                              '
                                                                                                                              'FaxNumber = IsFaxable("GPs", pGP)
                                                                                                                              'If FaxNumber = "" Then
                                                                                                                              '  FaxNumber = IsFaxable("Wards", pWard)
                                                                                                                              'End If
                                                                                                                              'FaxNumber = iBOX("Confirm FAX Number" & vbCrLf & "(Leave blank to Cancel FAX)", , FaxNumber)
                                                                                                                              'If FaxNumber = "" Then
                                                                                                                              '  iMsg "FAX Cancelled!", vbInformation
                                                                                                                              '  Exit Sub
                                                                                                                              'End If
                                                                                                                              '
                                                                                                                              'tb!FaxNumber = FaxNumber
                                                                                                                              '
                                                                                                                              'tb.Update

End Sub



Private Sub cmdPrint_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim SampleID As String

10    On Error GoTo cmdPrint_Click_Error

20    SampleID = grdSID.TextMatrix(grdSID.Row, 0)

30    If SampleID = "" Then Exit Sub

40    If iMsg("Report will be printed on" & vbCrLf & _
              WardEnqForcedPrinter & "." & vbCrLf & _
              "OK?", vbQuestion + vbYesNo) = vbYes Then

50      sql = "Select * from PrintPending where " & _
              "Department = 'M' " & _
              "and SampleID = '" & SampleID & "'"
60      Set tb = New Recordset
70      RecOpenClient 0, tb, sql
80      If tb.EOF Then
90        tb.AddNew
100     End If
110     tb!SampleID = SampleID
120     tb!Ward = pWard
130     tb!Clinician = pClinician
140     tb!GP = pGP
150     tb!Department = "M"
160     tb!Initiator = UserName
170     tb!UsePrinter = WardEnqForcedPrinter
180     tb!ThisIsCopy = 1
190     tb.Update

200   End If

210   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmExternalReport", "cmdPrint_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

10    PBar.Max = LogOffDelaySecs
20    PBar = 0
30    SingleUserUpdateLoggedOn UserName

40    Timer1.Enabled = True

50    If Activated Then Exit Sub
60    Activated = True

70    FillGrid

End Sub

Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub


Private Sub Form_Load()

10    Activated = False

20    PBar.Max = LogOffDelaySecs
30    PBar = 0

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


Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub grdSID_Click()

      Static SortOrder As Boolean
      Dim x As Integer
      Dim y As Integer

10    rtb.Text = ""
20    cmdFAX.Enabled = False

30    If grdSID.MouseRow = 0 Then
40      If SortOrder Then
50        grdSID.Sort = flexSortGenericAscending
60      Else
70        grdSID.Sort = flexSortGenericDescending
80      End If
90      SortOrder = Not SortOrder
100     Exit Sub
110   End If

120   For y = 1 To grdSID.Rows - 1
130     grdSID.Row = y
140     For x = 1 To grdSID.Cols - 1
150       grdSID.col = x
160       grdSID.CellBackColor = 0
170     Next
180   Next

190   grdSID.Row = grdSID.MouseRow
200   For x = 1 To grdSID.Cols - 1
210     grdSID.col = x
220     grdSID.CellBackColor = vbYellow
230   Next

240   FillReport grdSID.TextMatrix(grdSID.Row, 0)

250   If Trim$(rtb) <> "" Then
260     cmdFAX.Enabled = True
270   End If

End Sub

Private Sub grdSID_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1
  
20    If PBar = PBar.Max Then
30      Unload Me
40    End If

End Sub


