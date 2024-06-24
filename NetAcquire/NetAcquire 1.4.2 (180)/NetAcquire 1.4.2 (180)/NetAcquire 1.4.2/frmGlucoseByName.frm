VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlucoseByName 
   Caption         =   "NetAcquire - Glucose Tolerance"
   ClientHeight    =   6870
   ClientLeft      =   240
   ClientTop       =   645
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8970
   Begin VB.CommandButton cmdViewScanGlo 
      Caption         =   "&View Scan"
      Height          =   1020
      Left            =   7740
      Picture         =   "frmGlucoseByName.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   150
      TabIndex        =   29
      Top             =   1230
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton bPrintGTT 
      Caption         =   "&Print as GTT Report"
      Height          =   1005
      Left            =   7320
      Picture         =   "frmGlucoseByName.frx":57EE
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2820
      Width           =   1305
   End
   Begin VB.CommandButton bPrintSeries 
      Caption         =   "Print as  Glucose Series"
      Height          =   1005
      Left            =   7320
      Picture         =   "frmGlucoseByName.frx":5E58
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3990
      Width           =   1305
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   7620
      Picture         =   "frmGlucoseByName.frx":64C2
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6120
      Width           =   1035
   End
   Begin VB.PictureBox SSPanel2 
      BackColor       =   &H00C0C0C0&
      Height          =   2475
      Left            =   3540
      ScaleHeight     =   2415
      ScaleWidth      =   4995
      TabIndex        =   6
      Top             =   240
      Width           =   5055
      Begin VB.Label txtSampleID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3300
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lsex 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4395
         TabIndex        =   24
         Top             =   480
         Width           =   405
      End
      Begin VB.Label lage 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3255
         TabIndex        =   23
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lgp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   1980
         Width           =   3465
      End
      Begin VB.Label lward 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   21
         Top             =   1680
         Width           =   3465
      End
      Begin VB.Label lclinician 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   20
         Top             =   1380
         Width           =   3465
      End
      Begin VB.Label laddr1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   19
         Top             =   1080
         Width           =   3465
      End
      Begin VB.Label laddr0 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   18
         Top             =   780
         Width           =   3465
      End
      Begin VB.Label ldob 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   17
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lchart 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   16
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Index           =   10
         Left            =   765
         TabIndex        =   15
         Top             =   1710
         Width           =   390
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consultant"
         Height          =   195
         Index           =   9
         Left            =   405
         TabIndex        =   14
         Top             =   1410
         Width           =   750
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   8
         Left            =   2925
         TabIndex        =   13
         Top             =   510
         Width           =   285
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   7
         Left            =   705
         TabIndex        =   12
         Top             =   510
         Width           =   450
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   6
         Left            =   4095
         TabIndex        =   11
         Top             =   510
         Width           =   270
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Addr1"
         Height          =   195
         Index           =   5
         Left            =   735
         TabIndex        =   10
         Top             =   810
         Width           =   420
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Chart Number"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Addr2"
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   1110
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "G. P."
         Height          =   195
         Left            =   795
         TabIndex        =   7
         Top             =   2010
         Width           =   360
      End
   End
   Begin VB.TextBox tName 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   930
      Width           =   2445
   End
   Begin MSFlexGridLib.MSFlexGrid gName 
      Height          =   5265
      Left            =   150
      TabIndex        =   3
      Top             =   1410
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   9287
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Date         |<Name                               "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   675
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2865
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220856321
         CurrentDate     =   37585
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220856321
         CurrentDate     =   37585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3855
      Left            =   3510
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2820
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
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
      ScrollBars      =   2
      FormatString    =   "<Run #        |<Time     |<Serum mmol/L "
   End
   Begin VB.Image iFind 
      Height          =   480
      Left            =   3060
      Picture         =   "frmGlucoseByName.frx":6B2C
      Top             =   900
      Width           =   480
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   960
      Width           =   420
   End
End
Attribute VB_Name = "frmGlucoseByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CodeForGlucose As String

Private Sub FillgName()

          Dim tb As Recordset
          Dim sn As Recordset
          Dim tf As Recordset
          Dim sql As String
          Dim n As Integer
          Dim NameToFind As String
          Dim Found As Integer

52510     On Error GoTo FillgName_Error

52520     gName.Rows = 2
52530     gName.AddItem ""
52540     gName.RemoveItem 1

52550     g.Rows = 2
52560     g.AddItem ""
52570     g.RemoveItem 1

52580     pb.Visible = True

52590     sql = "Select Distinct PatName, D.RunDate " & _
              "from Demographics D, BioResults B where " & _
              "(D.RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & "') " & _
              "and PatName like '" & AddTicks(tName) & "%' " & _
              "and D.SampleID = B.SampleID " & _
              "order by D.rundate desc"

52600     Set tb = New Recordset
52610     RecOpenClient 0, tb, sql
52620     If tb.EOF Then Exit Sub

52630     pb.Min = 0
52640     pb = 0
52650     pb.max = tb.RecordCount

52660     Do While Not tb.EOF
52670         pb = pb + 1
52680         gName.AddItem Format$(tb!Rundate, "dd/mm/yy") & vbTab & tb!PatName & ""
52690         tb.MoveNext
52700     Loop

52710     pb = 0
52720     pb.max = gName.Rows

52730     For n = gName.Rows - 1 To 2 Step -1
52740         pb = pb + 1
52750         NameToFind = AddTicks(gName.TextMatrix(n, 1))
52760         Found = 0
52770         sql = "select * from demographics where " & _
                  "patname = '" & NameToFind & "' " & _
                  "and rundate = '" & Format$(gName.TextMatrix(n, 0), "dd/mmm/yyyy") & "' "
52780         Set sn = New Recordset
52790         RecOpenServer 0, sn, sql
52800         Do While Not sn.EOF
52810             sql = "Select * from BioResults where " & _
                      "SampleID = '" & sn!SampleID & "' " & _
                      "and RunDate = '" & Format$(gName.TextMatrix(n, 0), "dd/mmm/yyyy") & "' " & _
                      "and Code = '" & CodeForGlucose & "'"
52820             Set tf = New Recordset
52830             RecOpenServer 0, tf, sql
          
52840             If Not tf.EOF Then
52850                 Found = Found + 1
52860             End If
52870             sn.MoveNext
52880         Loop
52890         If Found < 2 Then
52900             gName.RemoveItem n
52910         End If
52920     Next

52930     If gName.Rows > 2 Then
52940         gName.RemoveItem 1
52950     End If

52960     pb.Visible = False

52970     Exit Sub

FillgName_Error:

          Dim strES As String
          Dim intEL As Integer

52980     intEL = Erl
52990     strES = Err.Description
53000     LogError "fGluByName", "FillgName", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

53010     Unload Me

End Sub

Private Sub bPrintGTT_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleID As Long

53020     On Error GoTo bPrintGTT_Click_Error

53030     SampleID = Val(g.TextMatrix(1, 0))
53040     If SampleID = 0 Then
53050         iMsg "Nothing to do!" & vbCrLf & "Select a Name to Print.", vbExclamation
53060         Exit Sub
53070     End If

53080     sql = "Select * from PrintPending where " & _
              "Department = 'G' " & _
              "and SampleID = '" & SampleID & "'"
53090     Set tb = New Recordset
53100     RecOpenClient 0, tb, sql
53110     If tb.EOF Then
53120         tb.AddNew
53130     End If
53140     tb!SampleID = SampleID
53150     tb!Ward = lward
53160     tb!Clinician = lclinician
53170     tb!GP = lgp
53180     tb!Department = "G"
53190     tb!Initiator = UserName
53200     tb.Update

53210     Exit Sub

bPrintGTT_Click_Error:

          Dim strES As String
          Dim intEL As Integer

53220     intEL = Erl
53230     strES = Err.Description
53240     LogError "fGluByName", "bPrintGTT_Click", intEL, strES, sql


End Sub

Private Sub bPrintSeries_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleID As Long

53250     On Error GoTo bPrintSeries_Click_Error

53260     SampleID = Val(g.TextMatrix(1, 0))
53270     If SampleID = 0 Then
53280         iMsg "Nothing to do!" & vbCrLf & "Select a Name to Print.", vbExclamation
53290         Exit Sub
53300     End If

53310     sql = "Select * from PrintPending where " & _
              "Department = 'S' " & _
              "and SampleID = '" & SampleID & "'"
53320     Set tb = New Recordset
53330     RecOpenClient 0, tb, sql
53340     If tb.EOF Then
53350         tb.AddNew
53360     End If
53370     tb!SampleID = SampleID
53380     tb!Ward = lward
53390     tb!Clinician = lclinician
53400     tb!GP = lgp
53410     tb!Department = "S"
53420     tb!Initiator = UserName
53430     tb.Update

53440     Exit Sub

bPrintSeries_Click_Error:

          Dim strES As String
          Dim intEL As Integer

53450     intEL = Erl
53460     strES = Err.Description
53470     LogError "fGluByName", "bPrintSeries_Click", intEL, strES, sql


End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdViewScanGlo_Click
' Author    : Masood
' Date      : 09/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewScanGlo_Click()
53480     On Error GoTo cmdViewScanGlo_Click_Error


53490     frmViewScan.SampleID = txtSampleID
53500     frmViewScan.txtSampleID = txtSampleID
53510     frmViewScan.Show 1

       
53520     Exit Sub

       
cmdViewScanGlo_Click_Error:

          Dim strES As String
          Dim intEL As Integer

53530     intEL = Erl
53540     strES = Err.Description
53550     LogError "frmGlucoseByName", "cmdViewScanGlo_Click", intEL, strES

End Sub

Private Sub Form_Load()

53560     dtFrom = Format$(Now - 60, "dd/mm/yyyy")
53570     dtTo = Format$(Now, "dd/mm/yyyy")

53580     CodeForGlucose = GetOptionSetting("BioCodeForGlucose", "")

End Sub


'---------------------------------------------------------------------------------------
' Procedure : g_Click
' Author    : Masood
' Date      : 09/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub g_Click()
53590     On Error GoTo g_Click_Error


53600     cmdViewScanGlo.Visible = False
53610     With g
53620         txtSampleID = .TextMatrix(.RowSel, 0)
53630         SetViewScans .TextMatrix(.RowSel, 0), cmdViewScanGlo
53640     End With


53650     Exit Sub


g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

53660     intEL = Erl
53670     strES = Err.Description
53680     LogError "frmGlucoseByName", "g_Click", intEL, strES
End Sub

Private Sub gName_Click()

          Dim sn As Recordset
          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim Found As Integer

53690     On Error GoTo gName_Click_Error

53700     g.Rows = 2
53710     g.AddItem ""
53720     g.RemoveItem 1

53730     If gName.MouseRow = 0 Then
53740         Exit Sub
53750     End If

53760     sql = "select * from demographics where " & _
              "patname = '" & AddTicks(gName.TextMatrix(gName.row, 1)) & "' " & _
              "and rundate = '" & Format$(gName.TextMatrix(gName.row, 0), "dd/mmm/yyyy") & "' " & _
              "order by SampleID"
53770     Set sn = New Recordset
53780     RecOpenClient 0, sn, sql

53790     If sn.EOF Then
53800         iMsg "No details found", vbExclamation
53810         Exit Sub
53820     End If

53830     lchart = sn!Chart & ""
53840     If Not IsNull(sn!DoB) Then
53850         ldob = sn!DoB
53860     Else
53870         ldob = ""
53880     End If
53890     lage = sn!Age & ""
53900     lsex = sn!Sex & ""
53910     laddr0 = sn!Addr0 & ""
53920     laddr1 = sn!Addr1 & ""
53930     lclinician = sn!Clinician & ""
53940     lward = sn!Ward & ""
53950     lgp = sn!GP & ""

53960     Do While Not sn.EOF
53970         Found = False
53980         s = Format$(sn!SampleDate & "", "hh:mm") & vbTab
53990         sql = "Select * from BioResults where " & _
                  "SampleID = '" & sn!SampleID & "' " & _
                  "and RunDate = '" & Format$(gName.TextMatrix(gName.row, 0), "dd/mmm/yyyy") & "' " & _
                  "and Code = '" & CodeForGlucose & "'"
54000         Set tb = New Recordset
54010         RecOpenClient 0, tb, sql
54020         If Not tb.EOF Then
54030             s = s & Format$(tb!Result, "0.0")
54040             Found = True
54050         End If
54060         If Found Then g.AddItem sn!SampleID & vbTab & s
54070         sn.MoveNext
54080     Loop

54090     If g.Rows > 2 Then
54100         g.RemoveItem 1
54110     End If

54120     Exit Sub

gName_Click_Error:

          Dim strES As String
          Dim intEL As Integer

54130     intEL = Erl
54140     strES = Err.Description
54150     LogError "fGluByName", "gName_Click", intEL, strES, sql

End Sub

Private Sub iFind_Click()
        
54160     If Len(Trim$(tName)) > 1 Then
54170         FillgName
54180     End If

End Sub

Private Sub tName_Change()
        
54190     If Len(Trim$(tName)) > 4 Then
54200         FillgName
54210     End If

End Sub


