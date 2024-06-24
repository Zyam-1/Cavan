VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlucose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Glucose Tolerance Test"
   ClientHeight    =   6390
   ClientLeft      =   510
   ClientTop       =   1695
   ClientWidth     =   11805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6390
   ScaleWidth      =   11805
   Begin VB.CommandButton cmdViewScanGlo 
      Caption         =   "&View Scan"
      Height          =   1020
      Left            =   10680
      Picture         =   "frmGlucose.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3180
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton bPrintSeries 
      Caption         =   "Print as Glucose Series"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7410
      Picture         =   "frmGlucose.frx":57EE
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Height          =   2985
      Left            =   6570
      TabIndex        =   5
      Top             =   90
      Width           =   4995
      Begin VB.Label txtSampleID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   2460
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lsex 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4365
         TabIndex        =   23
         Top             =   510
         Width           =   405
      End
      Begin VB.Label lage 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3225
         TabIndex        =   22
         Top             =   510
         Width           =   555
      End
      Begin VB.Label lgp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   21
         Top             =   2010
         Width           =   3465
      End
      Begin VB.Label lward 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   20
         Top             =   1710
         Width           =   3465
      End
      Begin VB.Label lclinician 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   19
         Top             =   1410
         Width           =   3465
      End
      Begin VB.Label laddr1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   18
         Top             =   1110
         Width           =   3465
      End
      Begin VB.Label laddr0 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   17
         Top             =   810
         Width           =   3465
      End
      Begin VB.Label ldob 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   16
         Top             =   510
         Width           =   1245
      End
      Begin VB.Label lchart 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   15
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   735
         TabIndex        =   14
         Top             =   1740
         Width           =   390
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consultant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   375
         TabIndex        =   13
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2895
         TabIndex        =   12
         Top             =   540
         Width           =   285
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   675
         TabIndex        =   11
         Top             =   540
         Width           =   450
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   4065
         TabIndex        =   10
         Top             =   540
         Width           =   270
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Addr1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   705
         TabIndex        =   9
         Top             =   840
         Width           =   420
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Chart Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Addr2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   705
         TabIndex        =   7
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "G. P."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   765
         TabIndex        =   6
         Top             =   2040
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   210
      TabIndex        =   3
      Top             =   90
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid gNames 
         Height          =   2535
         Left            =   60
         TabIndex        =   26
         Top             =   210
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         FormatString    =   "<Name                             |<Date of Birth  "
      End
      Begin MSComCtl2.DTPicker dtRun 
         Height          =   315
         Left            =   4500
         TabIndex        =   24
         Top             =   1020
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   37501
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   165
         Left            =   60
         TabIndex        =   4
         Top             =   2760
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Run Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4770
         TabIndex        =   25
         Top             =   810
         Width           =   690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2985
      Left            =   1590
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   5265
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
      FormatString    =   "<Run #          |<Date/Time                |<Serum  |<Urine  "
   End
   Begin VB.CommandButton bPrintGTT 
      Caption         =   "&Print as GTT Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7410
      Picture         =   "frmGlucose.frx":5E58
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4140
      Width           =   1875
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9900
      Picture         =   "frmGlucose.frx":64C2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4980
      Width           =   1245
   End
End
Attribute VB_Name = "frmGlucose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim CodeForGlucose As String

Private Sub FillNames()

          Dim tb As Recordset
          Dim tb1 As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Name2Find As String
          Dim Found As Integer
          Dim s As String
          Dim Count As Integer

50580     On Error GoTo FillNames_Error

50590     g.Rows = 2
50600     g.AddItem ""
50610     g.RemoveItem 1

50620     sql = "select distinct patname, DoB  from demographics D, BioResults B where " & _
              "D.RunDate = '" & Format$(dtRun, "dd/mmm/yyyy") & "' " & _
              "and D.SampleID = B.SampleID "

50630     Set tb = New Recordset
50640     RecOpenClient 0, tb, sql

50650     gNames.Visible = False
50660     gNames.Rows = 2
50670     gNames.AddItem ""
50680     gNames.RemoveItem 1

50690     If tb.EOF Then
50700         gNames.AddItem "None Found"
50710         gNames.RemoveItem 1
50720         gNames.Visible = True
50730         Exit Sub
50740     End If

50750     Do While Not tb.EOF
50760         s = tb!PatName & vbTab
50770         If IsDate(tb!DoB) Then
50780             s = s & Format$(tb!DoB, "dd/mm/yyyy")
50790         End If
50800         gNames.AddItem s
50810         tb.MoveNext
50820     Loop

50830     pb.Visible = True
50840     pb.max = gNames.Rows

50850     For n = gNames.Rows - 1 To 2 Step -1
50860         pb = pb.max - n
50870         Name2Find = gNames.TextMatrix(n, 0)
50880         Found = 0
        
50890         sql = "Select SampleID from Demographics where " & _
                  "patname = '" & AddTicks(Name2Find) & "' " & _
                  "and demographics.rundate = '" & Format$(dtRun, "dd/mmm/yyyy") & "'"
50900         If IsDate(gNames.TextMatrix(n, 1)) Then
50910             sql = sql & " and DoB = '" & Format$(gNames.TextMatrix(n, 1), "dd/mmm/yyyy") & "'"
50920         End If
50930         Set tb = New Recordset
50940         RecOpenClient 0, tb, sql
        
50950         Count = 0
50960         Do While Not tb.EOF
50970             sql = "select count (Code) as tCount from BioResults where " & _
                      "Code = '" & CodeForGlucose & "' " & _
                      "and SampleID = '" & tb!SampleID & "'"
50980             Set tb1 = New Recordset
50990             RecOpenClient 0, tb1, sql
51000             Count = Count + tb1!tCount
51010             tb.MoveNext
51020         Loop
        
51030         If Count < 2 Then
51040             gNames.RemoveItem n
51050         End If
51060     Next
51070     pb.Visible = False

51080     If gNames.Rows > 2 Then
51090         gNames.RemoveItem 1
51100     End If
51110     gNames.Visible = True

51120     Exit Sub

FillNames_Error:

          Dim strES As String
          Dim intEL As Integer

51130     intEL = Erl
51140     strES = Err.Description
51150     LogError "fglucose", "FillNames", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

51160     Unload Me

End Sub

Private Sub bPrintGTT_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleID As String

51170     On Error GoTo bPrintGTT_Click_Error

51180     SampleID = g.TextMatrix(1, 0)
51190     If Trim$(SampleID) = "" Then
51200         iMsg "Nothing to do!" & vbCrLf & "Select a Name to Print.", vbExclamation
51210         Exit Sub
51220     End If

51230     sql = "Select * from PrintPending where " & _
              "Department = 'G' " & _
              "and SampleID = '" & SampleID & "'"
51240     Set tb = New Recordset
51250     RecOpenClient 0, tb, sql
51260     If tb.EOF Then
51270         tb.AddNew
51280     End If
51290     tb!SampleID = SampleID
51300     tb!Ward = lward
51310     tb!Clinician = lclinician
51320     tb!GP = lgp
51330     tb!Department = "G"
51340     tb!Initiator = UserName
51350     tb.Update

51360     Exit Sub

bPrintGTT_Click_Error:

          Dim strES As String
          Dim intEL As Integer

51370     intEL = Erl
51380     strES = Err.Description
51390     LogError "fglucose", "bPrintGTT_Click", intEL, strES, sql


End Sub


Private Sub bPrintSeries_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleID As String

51400     On Error GoTo bPrintSeries_Click_Error

51410     SampleID = g.TextMatrix(1, 0)
51420     If Trim$(SampleID) = "" Then
51430         iMsg "Nothing to do!" & vbCrLf & "Select a Name to Print.", vbExclamation
51440         Exit Sub
51450     End If

51460     sql = "Select * from PrintPending where " & _
              "Department = 'S' " & _
              "and SampleID = '" & SampleID & "'"
51470     Set tb = New Recordset
51480     RecOpenClient 0, tb, sql
51490     If tb.EOF Then
51500         tb.AddNew
51510     End If
51520     tb!SampleID = SampleID
51530     tb!Ward = lward
51540     tb!Clinician = lclinician
51550     tb!GP = lgp
51560     tb!Department = "S"
51570     tb!Initiator = UserName
51580     tb.Update

51590     Exit Sub

bPrintSeries_Click_Error:

          Dim strES As String
          Dim intEL As Integer

51600     intEL = Erl
51610     strES = Err.Description
51620     LogError "fglucose", "bPrintSeries_Click", intEL, strES, sql


End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdViewScanGlo_Click
' Author    : Masood
' Date      : 09/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewScanGlo_Click()
51630     On Error GoTo cmdViewScanGlo_Click_Error


51640     frmViewScan.SampleID = txtSampleID
51650     frmViewScan.txtSampleID = txtSampleID
51660     frmViewScan.Show 1

       
51670     Exit Sub

       
cmdViewScanGlo_Click_Error:

          Dim strES As String
          Dim intEL As Integer

51680     intEL = Erl
51690     strES = Err.Description
51700     LogError "frmGlucose", "cmdViewScanGlo_Click", intEL, strES
End Sub

Private Sub dtRun_CloseUp()

51710     FillNames

51720     lchart = ""
51730     ldob = ""
51740     lage = ""
51750     lsex = ""
51760     laddr0 = ""
51770     laddr1 = ""
51780     lclinician = ""
51790     lward = ""
51800     lgp = ""

End Sub


Private Sub Form_Activate()

51810     FillNames

51820     CodeForGlucose = GetOptionSetting("BioCodeForGlucose", "")

End Sub

Private Sub Form_Load()


51830     dtRun = Format$(Now, "dd/mm/yyyy")

End Sub

'---------------------------------------------------------------------------------------
' Procedure : g_Click
' Author    : Masood
' Date      : 09/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub g_Click()

51840     On Error GoTo g_Click_Error

51850     cmdViewScanGlo.Visible = False
51860     With g
51870         txtSampleID = .TextMatrix(.RowSel, 0)
51880         SetViewScans .TextMatrix(.RowSel, 0), cmdViewScanGlo
51890     End With


51900     Exit Sub


g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

51910     intEL = Erl
51920     strES = Err.Description
51930     LogError "frmGlucose", "g_Click", intEL, strES
End Sub

Private Sub gNames_Click()

          Dim tb As Recordset
          Dim sn As Recordset
          Dim sql As String
          Dim s As String
          Dim Found As Integer
          Dim Name2Find As String

51940     On Error GoTo gNames_Click_Error

51950     If gNames.MouseRow = 0 Then
51960         Exit Sub
51970     End If
        
51980     Name2Find = gNames.TextMatrix(gNames.row, 0)

51990     sql = "select * from demographics where " & _
              "patname = '" & AddTicks(Name2Find) & "' " & _
              "and rundate = '" & Format$(dtRun, "dd/mmm/yyyy") & "' "
52000     If IsDate(gNames.TextMatrix(gNames.row, 1)) Then
52010         sql = sql & "and DoB = '" & Format$(gNames.TextMatrix(gNames.row, 1), "dd/mmm/yyyy") & "' "
52020     End If
52030     sql = sql & "order by SampleDate"
52040     Set tb = New Recordset
52050     RecOpenClient 0, tb, sql

52060     If tb.EOF Then
52070         iMsg "No details found", vbInformation
52080         Exit Sub
52090     End If

52100     lchart = tb!Chart & ""
52110     If Not IsNull(tb!DoB) Then
52120         ldob = tb!DoB
52130     Else
52140         ldob = ""
52150     End If
52160     lage = tb!Age & ""
52170     lsex = tb!Sex & ""
52180     laddr0 = tb!Addr0 & ""
52190     laddr1 = tb!Addr1 & ""
52200     lclinician = tb!Clinician & ""
52210     lward = tb!Ward & ""
52220     lgp = tb!GP & ""

52230     g.Rows = 2
52240     g.AddItem ""
52250     g.RemoveItem 1

52260     Do While Not tb.EOF
52270         Found = False
52280         s = Format$(tb!SampleDate & "", "dd/mm/yyyy hh:mm") & vbTab
52290         sql = "Select COALESCE(SampleType, 'S') SampleType, Result from BioResults where " & _
                  "SampleID = '" & tb!SampleID & "' " & _
                  "and Code = '" & CodeForGlucose & "'"
52300         Set sn = New Recordset
52310         RecOpenClient 0, sn, sql
52320         If Not sn.EOF Then
52330             If sn!SampleType = "S" Then
52340                 s = s & Format$(sn!Result, "0.0")
52350                 Found = True
52360             Else
52370                 s = s & vbTab & Format$(sn!Result, "0.0")
52380                 Found = True
52390             End If
52400         End If
52410         If Found Then g.AddItem tb!SampleID & vbTab & s
52420         tb.MoveNext
52430     Loop

52440     If g.Rows > 2 Then
52450         g.RemoveItem 1
52460     End If

52470     Exit Sub

gNames_Click_Error:

          Dim strES As String
          Dim intEL As Integer

52480     intEL = Erl
52490     strES = Err.Description
52500     LogError "fglucose", "gNames_Click", intEL, strES, sql

End Sub



