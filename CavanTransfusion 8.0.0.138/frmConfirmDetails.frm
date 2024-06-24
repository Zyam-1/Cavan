VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfirmDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   14205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnvalidate 
      Caption         =   "UnConfirm Details"
      Height          =   800
      Left            =   9810
      Picture         =   "frmConfirmDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1560
      Width           =   2000
   End
   Begin VB.TextBox txtNotes 
      Height          =   315
      Left            =   840
      TabIndex        =   16
      Top             =   2040
      Width           =   8625
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   800
      Left            =   210
      Picture         =   "frmConfirmDetails.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Width           =   2000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel with no Change"
      Height          =   800
      Left            =   11910
      Picture         =   "frmConfirmDetails.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "bCancel"
      Top             =   1560
      Width           =   2000
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "Confirm Details"
      Height          =   800
      Left            =   9810
      Picture         =   "frmConfirmDetails.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   690
      Width           =   2000
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   180
      TabIndex        =   17
      Top             =   7290
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3315
      Left            =   180
      TabIndex        =   23
      Top             =   2940
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   5847
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmConfirmDetails.frx":1680
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Historical Records (Confirmation History)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4657
      TabIndex        =   24
      Top             =   2580
      Width           =   4920
   End
   Begin VB.Label lblAandE 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   870
      TabIndex        =   21
      Top             =   540
      Width           =   1230
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "A / E"
      Height          =   195
      Left            =   450
      TabIndex        =   20
      Top             =   570
      Width           =   375
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3705
      TabIndex        =   19
      Top             =   915
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   2880
      TabIndex        =   18
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Notes:-"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   2100
      Width           =   510
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   2340
      TabIndex        =   14
      Top             =   6780
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblFG 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3570
      TabIndex        =   12
      Top             =   570
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblTypenex 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3705
      TabIndex        =   11
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   870
      TabIndex        =   10
      Top             =   1680
      Width           =   4170
   End
   Begin VB.Label lblSex 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   870
      TabIndex        =   9
      Top             =   1290
      Width           =   1230
   End
   Begin VB.Label lblMRN 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   870
      TabIndex        =   8
      Top             =   915
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Forward Group"
      Height          =   195
      Left            =   2460
      TabIndex        =   5
      Top             =   615
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   405
      TabIndex        =   4
      Top             =   1710
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   555
      TabIndex        =   3
      Top             =   1335
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Typenex"
      Height          =   195
      Left            =   2970
      TabIndex        =   2
      Top             =   1335
      Width           =   615
   End
   Begin VB.Label lblMRNAandE 
      AutoSize        =   -1  'True
      Caption         =   "MRN"
      Height          =   195
      Left            =   450
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Confirm the following details are correct:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   3510
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2445
      Left            =   180
      Top             =   120
      Width           =   13875
   End
End
Attribute VB_Name = "frmConfirmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Confirm(ByVal Valid As Integer)
  
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Confirm_Error

20    sql = "SELECT * FROM ConfirmDetails WHERE 0 = 1"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    tb.AddNew
60    tb!SampleID = lblSampleID
70    tb!AandE = lblAandE
80    tb!Chart = lblMRN
90    tb!Typenex = lblTypenex
100   tb!Sex = lblsex
110   tb!FwdGroup = lblFG
120   tb!Name = lblName
130   tb!Notes = txtNotes
140   tb!Operator = SecondUserName
150   tb!Confirmed = Valid
160   tb.Update



170   Unload Me

180   Exit Sub

Confirm_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmConfirmDetails", "Confirm", intEL, strES, sql

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "SELECT * FROM ConfirmDetails WHERE " & _
            "SampleID = '" & lblSampleID & "' " & _
            "ORDER BY DateTimeOfRecord DESC"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!Chart & vbTab & _
            tb!AandE & vbTab & _
            tb!Typenex & vbTab & _
            tb!Sex & vbTab & _
            tb!Name & vbTab & _
            tb!Notes & vbTab & _
            tb!Operator & vbTab & _
            Format$(tb!DateTimeOfRecord, "dd/MM/yyyy HH:mm:ss") & vbTab
100     If Not IsNull(tb!Confirmed) Then
110       s = s & IIf(tb!Confirmed, "Yes", "No")
120     Else
130       s = s & "No"
140     End If
150     g.AddItem s
160     tb.MoveNext
170   Loop

180   If g.Rows > 2 Then
190     g.RemoveItem 1
200   End If

210   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmConfirmDetails", "FillG", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdUnvalidate_Click()

10    Confirm 0

End Sub

Private Sub cmdValidate_Click()
  
10    Confirm 1

End Sub

Private Sub cmdXL_Click()

10    On Error GoTo cmdXL_Click_Error

20    ExportFlexGrid g, Me

30    Exit Sub

cmdXL_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmConfirmDetails", "cmdXL_Click", intEL, strES

End Sub


Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    With frmxmatch
30      lblMRN.Caption = .txtChart
40      lblAandE = .tAandE
50      lblTypenex = .tTypenex
60      lblsex = .lSex
70      lblFG = .lstfg.Text
80      lblName = .txtName
90      lblSampleID = .tLabNum
100     If .cmdValidate.Caption = "Confirm Details" Then
110       cmdUnvalidate.Enabled = False
120     ElseIf .cmdValidate.Caption = "UnConfirm" Then
130       cmdValidate.Enabled = False
140     End If
150   End With

160   FillG

170   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmConfirmDetails", "Form_Load", intEL, strES

End Sub


