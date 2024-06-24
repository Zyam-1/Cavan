VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExternalWorkList 
   Caption         =   "NetAcquire - External Worklist"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Between Date/Times"
      Height          =   585
      Left            =   150
      TabIndex        =   7
      Top             =   120
      Width           =   4515
      Begin MSComCtl2.DTPicker dtTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   3510
         TabIndex        =   8
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   219217923
         UpDown          =   -1  'True
         CurrentDate     =   0.686805555555556
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   285
         Left            =   2010
         TabIndex        =   9
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   219217921
         CurrentDate     =   39086
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "and"
         Height          =   195
         Left            =   1650
         TabIndex        =   11
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "04/01/2007 16:30"
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.ComboBox cmbSendTo 
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Text            =   "cmbSendTo"
      Top             =   330
      Width           =   2715
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6105
      Left            =   150
      TabIndex        =   4
      Top             =   810
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   10769
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmExternalWorkList.frx":0000
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   585
      Left            =   8490
      Picture         =   "frmExternalWorkList.frx":008D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   540
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   585
      Left            =   8490
      Picture         =   "frmExternalWorkList.frx":0397
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "bprint"
      Top             =   1530
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   585
      Left            =   8490
      Picture         =   "frmExternalWorkList.frx":0A01
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6330
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "External Site"
      Height          =   195
      Left            =   5670
      TabIndex        =   6
      Top             =   120
      Width           =   885
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
      Left            =   8490
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmExternalWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearG()

39170     g.Rows = 2
39180     g.AddItem ""
39190     g.RemoveItem 1

End Sub

Private Sub FillCombo()

          Dim sql As String
          Dim tb As Recordset
          Dim Current As String

39200     On Error GoTo FillCombo_Error

39210     cmbSendTo.Clear

39220     Current = Format$(dt, "dd/MMM/yyyy") & " " & Format$(dtTime, "HH:mm")

39230     sql = "SELECT DISTINCT(CAST(SendTo AS nvarchar(50))) AS S FROM ExtResults WHERE " & _
              "SampleID IN " & _
              " (Select SampleID FROM Demographics WHERE " & _
              "  DateTimeDemographics BETWEEN '" & _
              Format$(lblDateTime, "dd/MMM/yyyy HH:mm:ss") & "' AND '" & _
              Current & "') " & _
              "ORDER BY S"
39240     Set tb = New Recordset
39250     RecOpenServer 0, tb, sql
39260     Do While Not tb.EOF
39270         cmbSendTo.AddItem Trim$(tb!s & "")
39280         tb.MoveNext
39290     Loop

39300     Exit Sub

FillCombo_Error:

          Dim strES As String
          Dim intEL As Integer

39310     intEL = Erl
39320     strES = Err.Description
39330     LogError "frmExternalWorkList", "FillCombo", intEL, strES, sql


End Sub

Private Sub FillG()

          Dim tbSID As Recordset
          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim DetailsFilled As Boolean
          Dim Current As String

39340     On Error GoTo FillG_Error

39350     ClearG

39360     Current = Format$(dt, "dd/MMM/yyyy") & " " & Format$(dtTime, "HH:mm")

39370     sql = "SELECT DISTINCT(SampleID) FROM ExtResults WHERE " & _
              "SampleID IN " & _
              " (Select SampleID FROM Demographics WHERE " & _
              "  DateTimeDemographics BETWEEN '" & _
              Format$(lblDateTime, "dd/MMM/yyyy HH:mm:ss") & "' AND '" & _
              Current & "') " & _
              "AND SendTo LIKE '" & cmbSendTo & "' " & _
              "ORDER BY SampleID"
39380     Set tbSID = New Recordset
39390     RecOpenServer 0, tbSID, sql
39400     Do While Not tbSID.EOF
39410         sql = "SELECT D.SampleID, D.PatName, D.Chart, D.DoB, D.Ward, D.Clinician, D.GP, E.Analyte " & _
                  "FROM Demographics AS D, ExtResults AS E " & _
                  "WHERE D.SampleID = E.SampleID " & _
                  "AND SendTo LIKE '" & cmbSendTo & "' " & _
                  "AND D.SampleID = '" & tbSID!SampleID & "'"
39420         Set tb = New Recordset
39430         RecOpenServer 0, tb, sql
39440         DetailsFilled = False
39450         Do While Not tb.EOF
39460             If Not DetailsFilled Then
39470                 DetailsFilled = True
39480                 s = tb!SampleID & vbTab & _
                          tb!PatName & vbTab & tb!Chart & vbTab & _
                          tb!DoB & vbTab
39490                 If Trim$(tb!Ward & "") <> "" Then
39500                     s = s & tb!Ward
39510                 ElseIf Trim$(tb!Clinician & "") <> "" Then
39520                     s = s & tb!Clinician
39530                 Else
39540                     s = s & tb!GP & ""
39550                 End If
39560                 s = s & vbTab
39570             Else
39580                 s = vbTab & vbTab & vbTab & vbTab & vbTab
39590             End If
39600             s = s & tb!Analyte & ""
39610             g.AddItem s
39620             tb.MoveNext
39630         Loop
39640         tbSID.MoveNext
39650     Loop

39660     If g.Rows > 2 Then
39670         g.RemoveItem 1
39680     End If

39690     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

39700     intEL = Erl
39710     strES = Err.Description
39720     LogError "frmExternalWorkList", "FillG", intEL, strES, sql


End Sub

Private Sub SetDateTime()

          Dim D As Date
          Dim t As Date

          'previous day
39730     D = DateAdd("d", -1, dt)
          'one minute later
39740     t = DateAdd("n", 1, dtTime)

39750     lblDateTime = Format$(D + t, "dd/MM/yyyy HH:mm")

End Sub

Private Sub cmbSendTo_Click()

39760     FillG

End Sub

Private Sub cmbSendTo_KeyPress(KeyAscii As Integer)

39770     KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

39780     Unload Me

End Sub


Private Sub cmdPrint_Click()

          Dim Y As Integer

39790     Printer.Font.Name = "Courier New"
39800     Printer.Font.size = 15
39810     Printer.Print cmbSendTo & " External Request"
39820     Printer.Print
39830     Printer.Font.size = 10
39840     For Y = 0 To g.Rows - 1
              'Sample ID                                              'End of line
39850         Printer.Print Left$(g.TextMatrix(Y, 0) & Space$(10), 10); '    10
              'Patient Name
39860         Printer.Print Left$(g.TextMatrix(Y, 1) & Space$(20), 20); '    30
              'Chart
39870         Printer.Print Left$(g.TextMatrix(Y, 2) & Space$(7), 7); '      37
              'D.o.B.
39880         Printer.Print Left$(g.TextMatrix(Y, 3) & Space$(11), 11); '    48
              'Location
39890         Printer.Print Left$(g.TextMatrix(Y, 4) & Space$(10), 10); '    58
              'Requests
39900         Printer.Print Left$(g.TextMatrix(Y, 5) & Space$(22), 22) '     80
39910     Next

39920     Printer.EndDoc

End Sub

Private Sub cmdXL_Click()

39930     ExportFlexGrid g, Me

End Sub


Private Sub dt_CloseUp()

          Dim Orig As String
          Dim n As Integer

39940     ClearG
39950     Orig = cmbSendTo.Text

39960     SetDateTime

39970     FillCombo
39980     For n = 0 To cmbSendTo.ListCount - 1
39990         If Trim$(UCase$(Orig)) = Trim$(UCase$(cmbSendTo.List(n))) Then
40000             cmbSendTo.Text = Orig
40010             FillG
40020             Exit For
40030         End If
40040     Next

End Sub


Private Sub dtTime_Change()

40050     SetDateTime

End Sub

Private Sub Form_Load()

40060     dt = Format$(Now, "dd/MM/yyyy")
40070     dtTime = "16:29"
40080     SetDateTime

40090     FillCombo

End Sub


