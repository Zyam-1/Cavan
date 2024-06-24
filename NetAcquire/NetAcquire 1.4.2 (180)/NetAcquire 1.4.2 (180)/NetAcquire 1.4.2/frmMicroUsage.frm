VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMicroUsage 
   Caption         =   "NetAcquire"
   ClientHeight    =   6180
   ClientLeft      =   210
   ClientTop       =   525
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   9495
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   1290
      Visible         =   0   'False
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4485
      Left            =   180
      TabIndex        =   7
      Top             =   1560
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   7911
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
      FormatString    =   "<Site                                       |^In          |^AE          |^Out         |^GP         |^MGH       |^Total       "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between"
      Height          =   1125
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   3315
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1020
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1020
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optDates 
         Alignment       =   1  'Right Justify
         Caption         =   "Dates"
         Height          =   225
         Left            =   780
         TabIndex        =   9
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSIDs 
         Caption         =   "Sample Numbers"
         Height          =   225
         Left            =   1530
         TabIndex        =   8
         Top             =   0
         Width           =   1515
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
         Format          =   218824705
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
         Format          =   218824705
         CurrentDate     =   38126
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton breCalc 
      Caption         =   "Calculate"
      Height          =   825
      Left            =   3720
      Picture         =   "frmMicroUsage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   270
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   6300
      Picture         =   "frmMicroUsage.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   825
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   7140
      Picture         =   "frmMicroUsage.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   825
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
      TabIndex        =   2
      Top             =   510
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmMicroUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tbDem As Recordset
      Dim sql As String
      Dim s As String
      Dim lngFrom As Long
      Dim lngTo As Long
      Dim n As Integer
      Dim Isos As New Isolates
      Dim Iso As Isolate

      'g.FormatString = "<SID |<Date     |<Patient Name  |<Chart     " & _
       "|<D.o.B.      |<Clinician  |<Ward  |<G.P. "
56410 On Error GoTo FillG_Error

56420 g.Rows = 2
56430 g.AddItem ""
56440 g.RemoveItem 1

56450 sql = "Select * from Demographics where "

56460 If optDates Then
56470   If Abs(DateDiff("d", dtFrom, dtTo)) > 60 Then
56480     iMsg "Maximum 60 Days!", vbExclamation
56490     Exit Sub
56500   End If
56510   sql = sql & "Rundate between '" & Format(dtFrom, "dd/mmm/yyyy") & _
                    "' and '" & Format(dtTo, "dd/mmm/yyyy") & "'"
56520 Else
56530   lngFrom = Val(txtFrom)
56540   lngTo = Val(txtTo)
56550   If lngTo < lngFrom Then
56560     txtFrom = Format(lngTo)
56570     txtTo = Format(lngFrom)
56580     lngFrom = Val(txtFrom)
56590     lngTo = Val(txtTo)
56600   End If
56610   If lngFrom < 1 Or lngFrom > 9999999 Then
56620     iMsg "Number <From> is incorrect!", vbExclamation
56630     txtFrom = ""
56640     Exit Sub
56650   End If
56660   If lngTo < 1 Or lngTo > 9999999 Then
56670     iMsg "Number <To> is incorrect!", vbExclamation
56680     txtTo = ""
56690     Exit Sub
56700   End If
56710   If lngTo - lngFrom > 5000 Then
56720     iMsg "Maximum 5000 Records!", vbExclamation
56730     Exit Sub
56740   End If
      '350     sql = sql & "SampleID between '" & Format$(Val(txtFrom) + sysOptMicroOffset(0)) & "' " & _
      '                    " and '" & Format$(Val(txtTo) + sysOptMicroOffset(0)) & "'"
56750   sql = sql & "SampleID between '" & Format$(Val(txtFrom)) & "' " & _
                    " and '" & Format$(Val(txtTo)) & "'"
56760 End If
56770 sql = sql & " order by SampleID"

56780 Set tbDem = New Recordset
56790 RecOpenClient 0, tbDem, sql
56800 If tbDem.RecordCount > 0 Then
56810   pb.max = tbDem.RecordCount
56820   pb = 0
56830   pb.Visible = True
56840 End If
56850 Do While Not tbDem.EOF
56860   pb = pb + 1
56870   pb.Refresh
56880   s = Format$(Val(tbDem!SampleID)) & vbTab & _
            tbDem!Rundate & vbTab & _
            tbDem!PatName & vbTab & _
            tbDem!Chart & vbTab & _
            tbDem!DoB & vbTab & _
            tbDem!Clinician & vbTab & _
            tbDem!Ward & vbTab & _
            tbDem!GP & vbTab
        
56890   Isos.Load tbDem!SampleID
56900   For Each Iso In Isos
56910     For n = 1 To 4
56920       If Iso.IsolateNumber = n Then
56930         s = s & Iso.OrganismName & " " & Iso.Qualifier
56940       End If
56950       s = s & vbTab
56960     Next
56970   Next
56980   g.AddItem s
          
56990   tbDem.MoveNext
57000 Loop

57010 If g.Rows > 2 Then
57020   g.RemoveItem 1
57030 End If

57040 pb.Visible = False
57050 pb = 0


57060 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

57070 intEL = Erl
57080 strES = Err.Description
57090 LogError "frmMicroUsage", "FillG", intEL, strES, sql

End Sub

Private Sub breCalc_Click()

57100 FillG

End Sub

Private Sub cmdCancel_Click()

57110 Unload Me

End Sub

Private Sub cmdXL_Click()

57120 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

      Dim n As Integer

57130 dtFrom = Format(Now, "dd/mmm/yyyy")
57140 dtTo = dtFrom

57150 g.FormatString = "<Sample ID |<Date     |<Patient Name  |<Chart     " & _
                       "|<D.o.B.      |<Clinician  |<Ward  |<G.P. "
57160 For n = 1 To 4
57170   g.FormatString = g.FormatString & "|<Org " & Format$(n)
57180 Next

End Sub


Private Sub optDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

57190 dtFrom.Visible = True
57200 dtTo.Visible = True
57210 txtFrom.Visible = False
57220 txtTo.Visible = False

End Sub


Private Sub optSIDs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

57230 dtFrom.Visible = False
57240 dtTo.Visible = False
57250 txtFrom.Visible = True
57260 txtTo.Visible = True

End Sub


