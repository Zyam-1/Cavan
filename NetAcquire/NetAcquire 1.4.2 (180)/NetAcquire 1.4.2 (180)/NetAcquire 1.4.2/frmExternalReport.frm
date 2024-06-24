VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExternalReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - External Tests"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
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
      Picture         =   "frmExternalReport.frx":0000
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
         Enabled         =   -1  'True
         TextRTF         =   $"frmExternalReport.frx":066A
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Left            =   8640
      Picture         =   "frmExternalReport.frx":06EC
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
      Picture         =   "frmExternalReport.frx":0D56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
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
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmExternalReport.frx":1198
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
          Dim tbDem As Recordset
          Dim s As String
          Dim DemDone As Boolean

33570     On Error GoTo FillGrid_Error

33580     With grdSID
33590         .Rows = 2
33600         .AddItem ""
33610         .RemoveItem 1
33620     End With

33630     sql = "SELECT * FROM Demographics WHERE "
33640     If Trim$(lblChart) <> "" Then
33650         sql = sql & "Chart = '" & lblChart & "' AND "
33660     End If
33670     sql = sql & "PatName = '" & AddTicks(lblName) & "' "
33680     If IsDate(lblDoB) Then
33690         sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
33700     End If
33710     sql = sql & "ORDER BY SampleID DESC"
        
33720     Set tbDem = New Recordset
33730     RecOpenClient 0, tbDem, sql

33740     Do While Not tbDem.EOF
33750         sql = "SELECT * FROM ExtResults WHERE " & _
                  "SampleID = '" & tbDem!SampleID & "'"
33760         Set tb = New Recordset
33770         RecOpenServer 0, tb, sql
33780         DemDone = False
33790         Do While Not tb.EOF
33800             If Not DemDone Then
33810                 If Not IsNull(tbDem!DoB) Then
33820                     lblDoB = tbDem!DoB
33830                 Else
33840                     lblDoB = ""
33850                 End If
33860                 lblAge = tbDem!Age & ""
33870                 Select Case Left$(UCase$(tbDem!Sex & ""), 1)
                          Case "M": lblSex = "Male"
33880                     Case "F": lblSex = "Female"
33890                     Case Else: lblSex = ""
33900                 End Select
33910                 lblAddress = tbDem!Addr0 & " " & tbDem!Addr1 & ""
33920                 s = Format$(tb!SampleID) & vbTab
33930                 If IsDate(tbDem!SampleDate) Then
33940                     If Format(tbDem!SampleDate, "HH:mm") <> "00:00" Then
33950                         s = s & Format(tbDem!SampleDate, "dd/MM/yy HH:mm")
33960                     Else
33970                         s = s & Format(tbDem!SampleDate, "dd/MM/yy")
33980                     End If
33990                 Else
34000                     s = s & "Not Specified"
34010                 End If
34020                 DemDone = True
34030             Else
34040                 s = vbTab
34050             End If
          
34060             s = s & vbTab & _
                      tb!Analyte & vbTab & _
                      tb!Result & ""
34070             grdSID.AddItem s
34080             tb.MoveNext
34090         Loop
34100         tbDem.MoveNext
34110     Loop

34120     If grdSID.Rows > 2 Then
34130         grdSID.RemoveItem 1
34140     End If

34150     Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

34160     intEL = Erl
34170     strES = Err.Description
34180     LogError "frmExternalReport", "FillGrid", intEL, strES, sql


End Sub

Private Sub FillReport(ByVal SampleID As String)

          Dim sql As String
          Dim tb As Recordset

34190     On Error GoTo FillReport_Error

34200     sql = "Select * from ExtResults where " & _
              "SampleID = '" & SampleID & "'"
34210     Set tb = New Recordset
34220     RecOpenServer 0, tb, sql
34230     Do While Not tb.EOF

34240         With rtb
34250             .SelIndent = 0
34260             .SelColor = vbBlue
                  '.SelBold = False
                  '.SelText = "Analyte: "
34270             .SelBold = True
34280             .SelText = .SelText & tb!Analyte & ": "
34290             .SelColor = vbBlack
34300             .SelBold = True
34310             .SelIndent = 200
34320             If Trim$(tb!Result & "") <> "" Then
34330                 .SelText = .SelText & tb!Result & ""
34340             Else
34350                 .SelText = .SelText & "Not yet Available."
34360             End If
34370             .SelBold = False
34380             .SelText = .SelText & " (Sent To " & tb!SendTo & " )"
34390             .SelText = .SelText & vbCrLf
34400         End With

34410         tb.MoveNext
34420     Loop

34430     Exit Sub

FillReport_Error:

          Dim strES As String
          Dim intEL As Integer

34440     intEL = Erl
34450     strES = Err.Description
34460     LogError "frmExternalReport", "FillReport", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

34470     Unload Me

End Sub


Private Sub cmdFAX_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleID As String
          Dim FaxNumber As String

34480     On Error GoTo cmdFAX_Click_Error

34490     SampleID = grdSID.TextMatrix(grdSID.row, 0)

34500     sql = "Select * from PrintPending where " & _
              "Department = 'M' " & _
              "and SampleID = '" & SampleID & "' " & _
              "and UsePrinter = 'FAX'"
34510     Set tb = New Recordset
34520     RecOpenClient 0, tb, sql
34530     If tb.EOF Then
34540         tb.AddNew
34550     End If
34560     tb!SampleID = SampleID
34570     tb!Ward = pWard
34580     tb!Clinician = pClinician
34590     tb!GP = pGP
34600     tb!UsePrinter = "FAX"
        
          Dim Gx As New GP
34610     Gx.LoadName pGP
34620     FaxNumber = Gx.FAX

34630     If FaxNumber = "" Then
34640         FaxNumber = IsFaxable("Wards", pWard)
34650     End If
34660     FaxNumber = iBOX("Confirm FAX Number" & vbCrLf & "(Leave blank to Cancel FAX)", , FaxNumber)
34670     If FaxNumber = "" Then
34680         iMsg "FAX Cancelled!", vbInformation
34690         Exit Sub
34700     End If
        
34710     tb!FaxNumber = FaxNumber
        
34720     tb.Update

34730     Exit Sub

cmdFAX_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34740     intEL = Erl
34750     strES = Err.Description
34760     LogError "frmExternalReport", "cmdFAX_Click", intEL, strES, sql


End Sub



Private Sub cmdPrint_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleID As String

34770     On Error GoTo cmdPrint_Click_Error

34780     SampleID = grdSID.TextMatrix(grdSID.row, 0)

34790     sql = "Select * from PrintPending where " & _
              "Department = 'M' " & _
              "and SampleID = '" & SampleID & "'"
34800     Set tb = New Recordset
34810     RecOpenClient 0, tb, sql
34820     If tb.EOF Then
34830         tb.AddNew
34840     End If
34850     tb!SampleID = SampleID
34860     tb!Ward = pWard
34870     tb!Clinician = pClinician
34880     tb!GP = pGP
34890     tb!Department = "M"
34900     tb!Initiator = UserName
34910     tb!UsePrinter = ""
34920     tb.Update

34930     Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34940     intEL = Erl
34950     strES = Err.Description
34960     LogError "frmExternalReport", "cmdPrint_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

34970     pBar.max = LogOffDelaySecs
34980     pBar = 0

34990     Timer1.Enabled = True

35000     If Activated Then Exit Sub
35010     Activated = True

35020     FillGrid

End Sub

Private Sub Form_Deactivate()

35030     Timer1.Enabled = False

End Sub


Private Sub Form_Load()

35040     Activated = False

35050     pBar.max = LogOffDelaySecs
35060     pBar = 0

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35070     pBar = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

35080     Activated = False

End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35090     pBar = 0

End Sub


Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35100     pBar = 0

End Sub


Private Sub grdSID_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

35110     rtb.Text = ""
35120     cmdFAX.Enabled = False
35130     cmdPrint.Enabled = False

35140     If grdSID.MouseRow = 0 Then
35150         If SortOrder Then
35160             grdSID.Sort = flexSortGenericAscending
35170         Else
35180             grdSID.Sort = flexSortGenericDescending
35190         End If
35200         SortOrder = Not SortOrder
35210         Exit Sub
35220     End If


35230     For Y = grdSID.row To 1 Step -1
35240         If grdSID.TextMatrix(Y, 0) <> "" Then
35250             ySave = Y
35260             Exit For
35270         End If
35280     Next

35290     For Y = 1 To grdSID.Rows - 1
35300         grdSID.row = Y
35310         For X = 1 To grdSID.Cols - 1
35320             grdSID.Col = X
35330             grdSID.CellBackColor = 0
35340         Next
35350     Next

35360     grdSID.row = ySave
35370     For X = 1 To grdSID.Cols - 1
35380         grdSID.Col = X
35390         grdSID.CellBackColor = vbYellow
35400     Next

35410     FillReport grdSID.TextMatrix(grdSID.row, 0)

35420     If Trim$(rtb) <> "" Then
35430         cmdFAX.Enabled = True
35440         cmdPrint.Enabled = True
35450     End If

End Sub

Private Sub grdSID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35460     pBar = 0

End Sub


Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35470     pBar = 0

End Sub


Private Sub Timer1_Timer()

          'tmrRefresh.Interval set to 1000
35480     pBar = pBar + 1
        
35490     If pBar = pBar.max Then
35500         Unload Me
35510     End If

End Sub


