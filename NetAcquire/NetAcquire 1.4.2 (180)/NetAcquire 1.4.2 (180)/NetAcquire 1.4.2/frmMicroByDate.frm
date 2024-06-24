VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMicroByDate 
   Caption         =   "NetAcquire"
   ClientHeight    =   8190
   ClientLeft      =   240
   ClientTop       =   465
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9780
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   180
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6435
      Left            =   180
      TabIndex        =   13
      Top             =   1530
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   11351
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   7140
      Picture         =   "frmMicroByDate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   6300
      Picture         =   "frmMicroByDate.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   825
   End
   Begin VB.CommandButton breCalc 
      Caption         =   "Calculate"
      Height          =   825
      Left            =   3750
      Picture         =   "frmMicroByDate.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between"
      Height          =   1125
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   3315
      Begin VB.OptionButton optSIDs 
         Caption         =   "Sample Numbers"
         Height          =   225
         Left            =   1530
         TabIndex        =   4
         Top             =   0
         Width           =   1515
      End
      Begin VB.OptionButton optDates 
         Alignment       =   1  'Right Justify
         Caption         =   "Dates"
         Height          =   225
         Left            =   780
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1020
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1020
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219742209
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1020
         TabIndex        =   6
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219742209
         CurrentDate     =   38126
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   195
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
      Left            =   7980
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmMicroByDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub breCalc_Click()

36910 FillG

End Sub

Private Sub cmdCancel_Click()

36920 Unload Me

End Sub


Private Sub cmdXL_Click()

36930 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

36940 dtFrom = Format(Now, "dd/mmm/yyyy")
36950 dtTo = dtFrom

36960 g.FormatString = "<Sample ID |"

End Sub


Private Sub optDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

36970 dtFrom.Visible = True
36980 dtTo.Visible = True
36990 txtFrom.Visible = False
37000 txtTo.Visible = False

End Sub


Private Sub optSIDs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

37010 dtFrom.Visible = False
37020 dtTo.Visible = False
37030 txtFrom.Visible = True
37040 txtTo.Visible = True

End Sub


Private Sub FillG()

      Dim tb As Recordset
      Dim tbO As Recordset
      Dim tbRpt As Recordset
      Dim sql As String
      Dim s As String
      Dim lngFrom As Long
      Dim lngTo As Long
      Dim n As Integer
      Dim X As Integer
      Dim gRow As Integer

37050 On Error GoTo FillG_Error

37060 g.Rows = 2
37070 g.AddItem ""
37080 g.RemoveItem 1

37090 sql = "Select Code from Antibiotics where " & _
            "AntibioticName <> 'None' " & _
            "and AntibioticName <> 'Antibiotic not stated' " & _
            "order by listorder"
37100 Set tb = New Recordset
37110 RecOpenClient 0, tb, sql
37120 s = "<Sample ID        |<Organism Group     |<Organism Name     |"
37130 Do While Not tb.EOF
37140   s = s & ">" & tb!Code & "     |^" & tb!Code & "RSI|"
37150   tb.MoveNext
37160 Loop
37170 g.FormatString = s
37180 If tb.RecordCount > 0 Then
37190   g.Cols = tb.RecordCount * 2 + 3
37200 End If

37210 sql = "Select distinct SampleID, IsolateNumber from Sensitivities where "
37220 If optDates Then
37230   If Abs(DateDiff("d", dtFrom, dtTo)) > 60 Then
37240     iMsg "Maximum 60 Days!", vbExclamation
37250     Exit Sub
37260   End If
37270   sql = sql & "Rundate between '" & Format(dtFrom, "dd/mmm/yyyy") & _
                    "' and '" & Format(dtTo, "dd/mmm/yyyy") & "'"
37280 Else
37290   lngFrom = Val(txtFrom)
37300   lngTo = Val(txtTo)
37310   If lngTo < lngFrom Then
37320     txtFrom = Format(lngTo)
37330     txtTo = Format(lngFrom)
37340     lngFrom = Val(txtFrom)
37350     lngTo = Val(txtTo)
37360   End If
37370   If lngFrom < 1 Or lngFrom > 9999999 Then
37380     iMsg "Number <From> is incorrect!", vbExclamation
37390     txtFrom = ""
37400     Exit Sub
37410   End If
37420   If lngTo < 1 Or lngTo > 9999999 Then
37430     iMsg "Number <To> is incorrect!", vbExclamation
37440     txtTo = ""
37450     Exit Sub
37460   End If
37470   If lngTo - lngFrom > 5000 Then
37480     iMsg "Maximum 5000 Records!", vbExclamation
37490     Exit Sub
37500   End If
      '470     sql = sql & "SampleID between '" & Format$(Val(txtFrom) + sysOptMicroOffset(0)) & "' " & _
      '                    " and '" & Format$(Val(txtTo) + sysOptMicroOffset(0)) & "'"
37510   sql = sql & "SampleID between '" & Format$(Val(txtFrom)) & "' " & _
                    " and '" & Format$(Val(txtTo)) & "'"
37520 End If
37530 sql = sql & "order by SampleID, IsolateNumber"

37540 Set tb = New Recordset
37550 RecOpenClient 0, tb, sql
37560 If Not tb.EOF Then
37570   pb.max = tb.RecordCount
37580   pb = 0
37590   pb.Visible = True
37600   g.Visible = False
37610 End If
37620 Do While Not tb.EOF
        
37630   pb = pb + 1
37640   s = tb!SampleID
37650   If tb!IsolateNumber <> 1 Then
37660     s = s & "." & Format$(tb!IsolateNumber)
37670   End If
37680   s = s & vbTab
        
37690   sql = "Select OrganismGroup, OrganismName from Isolates where " & _
              "SampleID = '" & tb!SampleID & "' " & _
              "and IsolateNumber = '" & tb!IsolateNumber & "'"
37700   Set tbO = New Recordset
37710   RecOpenServer 0, tbO, sql
37720   If Not tbO.EOF Then
37730     s = s & tbO!OrganismGroup & vbTab & _
                tbO!OrganismName & vbTab
37740   Else
37750     s = vbTab & vbTab
37760   End If

37770   g.AddItem s
        
37780   gRow = g.Rows - 1
        
37790   sql = "Select AntibioticCode, Result, RSI from Sensitivities where " & _
              "SampleID = '" & tb!SampleID & "' " & _
              "and IsolateNumber = '" & tb!IsolateNumber & "' "
37800   Set tbRpt = New Recordset
37810   RecOpenServer 0, tbRpt, sql
37820   Do While Not tbRpt.EOF
37830     For X = 3 To g.Cols - 1 Step 2
37840       If tbRpt!AntibioticCode = g.TextMatrix(0, X) Then
37850         g.TextMatrix(gRow, X) = tbRpt!Result & ""
37860         g.TextMatrix(gRow, X + 1) = tbRpt!RSI & ""
37870         Exit For
37880       End If
37890     Next
37900     tbRpt.MoveNext
37910   Loop
          
37920   tb.MoveNext
37930 Loop

37940 If g.Rows > 2 Then
37950   g.RemoveItem 1
37960   For X = 1 To g.Cols - 1 Step 4
37970     g.Col = X
37980     For n = 1 To g.Rows - 1
37990       g.row = n
38000       g.CellBackColor = &HFFFFC0
38010     Next
38020     g.Col = X + 1
38030     For n = 1 To g.Rows - 1
38040       g.row = n
38050       g.CellBackColor = &HFFFFC0
38060     Next
38070   Next
38080 End If

38090 g.Visible = True

38100 pb.Visible = False
38110 pb = 0


38120 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

38130 intEL = Erl
38140 strES = Err.Description
38150 LogError "frmMicroByDate", "FillG", intEL, strES, sql


End Sub


