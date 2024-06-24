VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportCollated 
   Caption         =   "NetAcquire - Collated Report"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWait 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3540
      TabIndex        =   11
      Text            =   "Generating Report. Please wait."
      Top             =   2130
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   0
      Left            =   10800
      Picture         =   "frmReportCollated.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Print This Report"
      Top             =   3270
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid gReport 
      Height          =   7065
      Left            =   150
      TabIndex        =   9
      Top             =   900
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12462
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmReportCollated.frx":066A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   885
      Left            =   10800
      Picture         =   "frmReportCollated.frx":06FA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4230
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   885
      Left            =   10800
      Picture         =   "frmReportCollated.frx":0A04
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Report"
      Height          =   885
      Left            =   10800
      Picture         =   "frmReportCollated.frx":106E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   930
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clinician"
      Height          =   675
      Left            =   3210
      TabIndex        =   3
      Top             =   150
      Width           =   2955
      Begin VB.ComboBox cmbClinician 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "cmbClinician"
         Top             =   240
         Width           =   2685
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   675
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2955
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220200961
         CurrentDate     =   39220
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220200961
         CurrentDate     =   39220
      End
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10560
      TabIndex        =   8
      Top             =   5130
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmReportCollated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearGrid()

45250 With gReport
45260   .Rows = 2
45270   .AddItem ""
45280   .RemoveItem 1
45290 End With

End Sub

Private Sub FillClinicians()

      Dim tb As Recordset
      Dim sql As String
      Dim StartDate As String
      Dim StopDate As String

45300 On Error GoTo FillClinicians_Error

45310 StartDate = Format$(dtFrom, "Long Date")
45320 StopDate = Format$(dtTo, "Long Date")

45330 ClearGrid

45340 cmbClinician.Clear
45350 cmdGenerate.Visible = False

45360 sql = "SELECT DISTINCT Clinician FROM Demographics WHERE " & _
            "RunDate BETWEEN '" & StartDate & "' AND '" & StopDate & "' " & _
            "AND Clinician IS NOT NULL " & _
            "AND Clinician <> '' " & _
            "ORDER BY Clinician"
45370 Set tb = New Recordset
45380 RecOpenServer 0, tb, sql
45390 If Not tb.EOF Then
45400   cmdGenerate.Visible = True
45410   Do While Not tb.EOF
45420     cmbClinician.AddItem tb!Clinician
45430     tb.MoveNext
45440   Loop
45450 End If

45460 Exit Sub

FillClinicians_Error:

      Dim strES As String
      Dim intEL As Integer

45470 intEL = Erl
45480 strES = Err.Description
45490 LogError "frmReportCollated", "FillClinicians", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

45500 Unload Me

End Sub

Private Sub cmdGenerate_Click()

      Dim tb As Recordset
      Dim tbDem As Recordset
      Dim sql As String
      Dim StartDate As String
      Dim StopDate As String
      Dim s As String
      Dim NameFilled As Boolean
      Dim gColour As Boolean
      Dim X As Integer
      Dim t As Single

45510 On Error GoTo cmdGenerate_Click_Error

45520 txtWait.Visible = True
45530 txtWait.Refresh

45540 StartDate = Format$(dtFrom, "Long Date")
45550 StopDate = Format$(dtTo, "Long Date") '& " 23:59:59"

45560 ClearGrid

45570 gReport.Visible = False

45580 sql = "SELECT DISTINCT(D.SampleID), D.RunDate, PatName, DoB, Chart " & _
            "FROM Demographics AS D, BioResults AS R WHERE " & _
            "Clinician = '" & cmbClinician & "' " & _
            "AND D.RunDate BETWEEN '" & StartDate & "' AND '" & StopDate & "' " & _
            "AND D.SampleID = R.SampleID " & _
            "ORDER BY D.SampleID"
45590 Set tbDem = New Recordset
45600 t = Timer
45610 RecOpenServer 0, tbDem, sql
45620 Do While Not tbDem.EOF
45630   NameFilled = False
45640   sql = "SELECT D.ShortName, R.Result, D.DP FROM BioResults AS R, BioTestDefinitions AS D WHERE " & _
              "D.Code = R.Code " & _
              "AND R.SampleID = '" & tbDem!SampleID & "'"
45650   Set tb = New Recordset
45660   RecOpenServer 0, tb, sql
45670   If Not tb.EOF Then
45680     Do While Not tb.EOF
45690       s = IIf(NameFilled, "", tbDem!SampleID) & vbTab & _
                IIf(NameFilled, "", tbDem!Rundate) & vbTab & _
                IIf(NameFilled, "", tbDem!PatName & "") & vbTab & _
                IIf(NameFilled, "", tbDem!DoB & "") & vbTab & _
                IIf(NameFilled, "", tbDem!Chart & "") & vbTab & _
                tb!ShortName & vbTab
45700       If IsNumeric(tb!Result & "") Then
45710         s = s & FormatNumber(tb!Result, tb!DP)
45720       Else
45730         s = s & tb!Result & ""
45740       End If
45750       gReport.AddItem s
45760       If gColour Then
45770         gReport.row = gReport.Rows - 1
45780         For X = 0 To gReport.Cols - 1
45790           gReport.Col = X
45800           gReport.CellBackColor = &HFFFF80
45810         Next
45820       End If
45830       NameFilled = True
45840       tb.MoveNext
45850     Loop
45860     gColour = Not gColour
45870   End If
45880   tbDem.MoveNext
45890 Loop
45900 Debug.Print Timer - t

45910 If gReport.Rows > 2 Then
45920   gReport.RemoveItem 1
45930 End If
45940 gReport.Visible = True

45950 txtWait.Visible = False

45960 Exit Sub

cmdGenerate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

45970 intEL = Erl
45980 strES = Err.Description
45990 LogError "frmReportCollated", "cmdGenerate_Click", intEL, strES, sql


End Sub

Private Sub cmdXL_Click()

46000 ExportFlexGrid gReport, Me

End Sub


Private Sub dtFrom_CloseUp()

46010 If DateDiff("d", dtFrom, dtTo) > 7 Then
46020   dtTo = dtFrom + 7
46030 End If

46040 FillClinicians

End Sub

Private Sub dtTo_CloseUp()

46050 If DateDiff("d", dtFrom, dtTo) > 7 Then
46060   dtFrom = dtTo - 7
46070 End If

46080 FillClinicians

End Sub

Private Sub Form_Load()

46090 dtFrom = Format$(Now - 7, "Short Date")
46100 dtTo = Format$(Now, "Short Date")

46110 FillClinicians

End Sub

