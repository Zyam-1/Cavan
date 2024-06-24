VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUPro 
   Caption         =   "NetAcquire - Urine Protein"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewScan 
      Caption         =   "&View Scan"
      Height          =   1020
      Left            =   9540
      Picture         =   "frmUPro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3960
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   705
      Left            =   6930
      Picture         =   "frmUPro.frx":57EE
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   660
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   6930
      Picture         =   "frmUPro.frx":5C30
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4410
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   150
      TabIndex        =   19
      Top             =   210
      Width           =   5445
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   20
         Top             =   840
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   3180
         TabIndex        =   21
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1710
         TabIndex        =   22
         Top             =   420
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   37505
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Patient Name"
         Height          =   195
         Left            =   660
         TabIndex        =   24
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Run Dates Between"
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Demographic Details"
      Height          =   2085
      Left            =   5850
      TabIndex        =   7
      Top             =   1770
      Width           =   4605
      Begin VB.Label lblRunDate 
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
         Left            =   840
         TabIndex        =   18
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblComment 
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
         Left            =   840
         TabIndex        =   17
         Top             =   1560
         Width           =   3525
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
         Left            =   840
         TabIndex        =   16
         Top             =   630
         Width           =   3525
      End
      Begin VB.Label lblChart 
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
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblDoB 
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
         Left            =   840
         TabIndex        =   14
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   13
         Top             =   1260
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Run Date"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SID"
         Height          =   195
         Index           =   0
         Left            =   2730
         TabIndex        =   9
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblSID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3060
         TabIndex        =   8
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1725
      Left            =   150
      TabIndex        =   0
      Top             =   3990
      Width           =   5445
      Begin VB.TextBox txtHours 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1830
         TabIndex        =   33
         Text            =   "24"
         Top             =   300
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   705
         Left            =   3900
         Picture         =   "frmUPro.frx":629A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   690
         Width           =   1125
      End
      Begin VB.TextBox txtVolume 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         TabIndex        =   27
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         Height          =   195
         Left            =   2910
         TabIndex        =   34
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Collection Period"
         Height          =   195
         Left            =   570
         TabIndex        =   32
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "mL"
         Height          =   195
         Left            =   2910
         TabIndex        =   29
         Top             =   690
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Urinary Volume"
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   660
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "g/24Hr"
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "g/L"
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   990
         Width           =   255
      End
      Begin VB.Label lblUPro24 
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
         Left            =   1830
         TabIndex        =   4
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Urinary Protein"
         Height          =   195
         Left            =   735
         TabIndex        =   3
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblUPro 
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
         Left            =   1830
         TabIndex        =   2
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Urinary Protein"
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2025
      Left            =   150
      TabIndex        =   26
      Top             =   1860
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   3572
      _Version        =   393216
      Cols            =   3
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
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Patient Name                              |<Run Date               |<Sample ID #      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmUPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CodeForUProt As String

Private Sub ClearAll()

28900 lblRunDate = ""
28910 lblName = ""
28920 lblChart = ""
28930 lblDoB = ""
28940 lblComment = ""

28950 txtHours = "24"
28960 lblUPro = ""
28970 lblUPro24 = ""

28980 lblSID = ""

28990 txtVolume = ""

End Sub

Private Sub FillDetails(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim OBs As Observations
      Dim n As Integer
      Dim v As Long

29000 On Error GoTo FillDetails_Error

29010 lblSID = SampleID
29020 cmdViewScan.Visible = False


29030 sql = "select * from demographics where " & _
            "SampleID = '" & SampleID & "'"
29040 Set tb = New Recordset
29050 RecOpenServer 0, tb, sql
29060 If tb.EOF Then
29070   lblRunDate = ""
29080   lblName = ""
29090   lblChart = ""
29100   lblDoB = ""
29110   lblComment = ""
29120 Else
29130   lblRunDate = tb!Rundate
29140   lblChart = tb!Chart & ""
29150   lblName = tb!PatName & ""
29160   lblDoB = tb!DoB & ""
29170 End If

29180 lblComment = ""
29190 Set OBs = New Observations
29200 Set OBs = OBs.Load(lblSID, "Demographic")
29210 If Not OBs Is Nothing Then
29220   lblComment = OBs.Item(1).Comment
29230 End If
29240 Set OBs = New Observations
29250 Set OBs = OBs.Load(lblSID, "Biochemistry")
29260 If Not OBs Is Nothing Then
29270   lblComment = lblComment & OBs.Item(1).Comment
29280 End If

29290 sql = "SELECT COALESCE(Result, '0') AS Result FROM BioResults WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Code = '" & CodeForUProt & "'"
29300 Set tb = New Recordset
29310 RecOpenServer 0, tb, sql
29320 If Not tb.EOF Then
29330   lblUPro = Format$(tb!Result, "###0.000")
29340 Else
29350   lblUPro = ""
29360 End If

29370 sql = "SELECT Result FROM BioResults WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Code = 'TUV'" 'total urine volume
29380 Set tb = New Recordset
29390 RecOpenServer 0, tb, sql
29400 If Not tb.EOF Then
29410   txtVolume = tb!Result & ""
29420 Else
29430   txtVolume = ""
29440 End If

29450 If txtVolume = "" And lblComment <> "" Then
29460   For n = 1 To Len(lblComment)
29470     v = Val(Mid$(lblComment, n))
29480     If v <> 0 Then
29490       txtVolume = Format$(v)
29500       Exit For
29510     End If
29520   Next
29530 End If

29540 Calculate
29550 SetViewScans lblSID, cmdViewScan
29560 Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

29570 intEL = Erl
29580 strES = Err.Description
29590 LogError "frmUPro", "FillDetails", intEL, strES, sql

End Sub
Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String

29600 On Error GoTo FillG_Error

29610 g.Rows = 2
29620 g.AddItem ""
29630 g.RemoveItem 1

29640 If Trim$(txtName) = "" Then Exit Sub

29650 Screen.MousePointer = vbHourglass

29660 sql = "SELECT DISTINCT D.SampleID, PatName, B.RunDate " & _
            "FROM Demographics AS D, BioResults AS B WHERE " & _
            "D.PatName LIKE '" & AddTicks(txtName) & "%' " & _
            "AND D.SampleID = B.SampleID " & _
            "AND B.RunDate BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
            "AND '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
            "AND B.Code = '" & CodeForUProt & "'"

29670 Set tb = New Recordset
29680 RecOpenServer 0, tb, sql
29690 Do While Not tb.EOF
29700   g.AddItem tb!PatName & vbTab & tb!Rundate & vbTab & tb!SampleID
29710   tb.MoveNext
29720 Loop

29730 If g.Rows > 2 Then
29740   g.RemoveItem 1
29750 End If

29760 Screen.MousePointer = 0

29770 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

29780 intEL = Erl
29790 strES = Err.Description
29800 LogError "frmUPro", "FillG", intEL, strES, sql


End Sub




Private Sub Calculate()

29810 If Val(lblUPro) <> 0 And Val(txtVolume) <> 0 Then
29820   lblUPro24 = Format$((Val(lblUPro) * (Val(txtVolume) / 1000)) * 24 / (Val(txtHours)), "0.000")
29830 End If

End Sub
Private Sub cmdCancel_Click()

29840 Unload Me

End Sub


Private Sub cmdPrint_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim Ward As String
      Dim Clin As String
      Dim GP As String

29850 On Error GoTo cmdPrint_Click_Error

29860 sql = "SELECT * FROM UPro WHERE " & _
            "SampleID = '" & lblSID & "'"
29870 Set tb = New Recordset
29880 RecOpenServer 0, tb, sql
29890 If tb.EOF Then
29900   tb.AddNew
29910   tb!SampleID = Val(lblSID)
29920 End If
29930 tb!CollectionPeriod = Val(txtHours)
29940 tb!TotalVolume = Val(txtVolume)
29950 tb!UPgPerL = Val(lblUPro)
29960 tb!UP24H = Val(lblUPro24)
29970 tb!PrintedDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
29980 tb!PrintedBy = UserName
29990 tb.Update

30000 GetWardClinGP lblSID, Ward, Clin, GP

30010 sql = "Select * from PrintPending where " & _
            "Department = 'U' " & _
            "and SampleID = '" & lblSID & "'"
        
30020 Set tb = New Recordset
30030 RecOpenClient 0, tb, sql
30040 If tb.EOF Then
30050   tb.AddNew
30060 End If
30070 tb!SampleID = lblSID
30080 tb!Ward = Ward
30090 tb!Clinician = Clin
30100 tb!GP = GP
30110 tb!Department = "U"
30120 tb!Initiator = UserName
30130 tb.Update

30140 Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

30150 intEL = Erl
30160 strES = Err.Description
30170 LogError "frmUPro", "cmdPrint_Click", intEL, strES, sql


End Sub

Private Sub cmdSearch_Click()

30180 ClearAll
30190 FillG

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdViewScan_Click
' Author    : Masood
' Date      : 02/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewScan_Click()
30200   On Error GoTo cmdViewScan_Click_Error


30210 frmViewScan.SampleID = lblSID
30220 frmViewScan.txtSampleID = lblSID
30230 frmViewScan.Show 1

       
30240 Exit Sub

       
cmdViewScan_Click_Error:

      Dim strES As String
      Dim intEL As Integer

30250 intEL = Erl
30260 strES = Err.Description
30270 LogError "frmUPro", "cmdViewScan_Click", intEL, strES
End Sub

Private Sub dtFrom_CloseUp()

30280 cmdSearch.Visible = True

End Sub


Private Sub dtTo_CloseUp()

30290 cmdSearch.Visible = True

End Sub


Private Sub Form_Load()

30300 dtFrom = Format$(Now - 30, "dd/mm/yyyy")
30310 dtTo = Format$(Now, "dd/mm/yyyy")

30320 CodeForUProt = GetOptionSetting("BioCodeForUProt", "")

End Sub


Private Sub g_Click()

      Dim ySave As Integer
      Dim n As Integer

30330 If g.MouseRow = 0 Then Exit Sub

30340 If g.TextMatrix(g.row, 0) = "" Then Exit Sub

30350 ySave = g.row

30360 If g.TextMatrix(ySave, 2) <> "" Then
30370   g.Col = 2
30380 Else
30390   g.Col = 3
30400 End If
30410 For n = 1 To g.Rows - 1
30420   g.row = n
30430   g.CellBackColor = 0
30440 Next
30450 g.row = ySave
30460 g.CellBackColor = vbRed
        
30470 If g.Col = 2 Then
30480   FillDetails g.TextMatrix(ySave, 2)
30490 End If

End Sub


Private Sub txtHours_Click()

30500 Select Case txtHours
        Case "24": txtHours = "3"
30510   Case "3": txtHours = "4"
30520   Case "4": txtHours = "6"
30530   Case "6": txtHours = "12"
30540   Case "12": txtHours = "24"
30550 End Select
       
30560 Calculate
       
End Sub


Private Sub txtHours_KeyUp(KeyCode As Integer, Shift As Integer)

30570 KeyCode = 0

End Sub


Private Sub txtVolume_LostFocus()

30580 Calculate

End Sub


