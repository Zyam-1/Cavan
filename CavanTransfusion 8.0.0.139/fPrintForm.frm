VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fPrintForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Print Form"
   ClientHeight    =   9300
   ClientLeft      =   570
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "fPrintForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1035
      Left            =   10815
      Picture         =   "fPrintForm.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   7980
      Width           =   915
   End
   Begin VB.CommandButton bPrintCord 
      Caption         =   "Print Co&rd Report"
      Height          =   705
      Left            =   6795
      Picture         =   "fPrintForm.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   8250
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   915
      Left            =   9495
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Text            =   "fPrintForm.frx":1DFE
      Top             =   3150
      Width           =   1695
   End
   Begin VB.ComboBox cmbComment 
      Height          =   315
      Left            =   3420
      TabIndex        =   52
      Text            =   "cmbComment"
      Top             =   4080
      Width           =   5115
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Print Pre&view"
      Height          =   255
      Left            =   8355
      TabIndex        =   50
      Top             =   7710
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrintAN 
      Caption         =   "Print &A/N Report"
      Height          =   705
      Left            =   6795
      Picture         =   "fPrintForm.frx":1E3E
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7470
      Width           =   1425
   End
   Begin VB.Frame Frame4 
      Height          =   1545
      Left            =   5535
      TabIndex        =   39
      Top             =   7380
      Width           =   1185
      Begin VB.TextBox txtMaxHrs 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   180
         TabIndex        =   41
         Text            =   "4"
         Top             =   780
         Width           =   435
      End
      Begin MSComCtl2.UpDown udMaxHrs 
         Height          =   375
         Left            =   630
         TabIndex        =   40
         Top             =   750
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   4
         BuddyControl    =   "txtMaxHrs"
         BuddyDispid     =   196616
         OrigLeft        =   630
         OrigTop         =   780
         OrigRight       =   1020
         OrigBottom      =   1065
         Max             =   12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Use Plasma within"
         Height          =   405
         Left            =   120
         TabIndex        =   43
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         Height          =   195
         Left            =   330
         TabIndex        =   42
         Top             =   1140
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Index           =   0
      Left            =   1980
      TabIndex        =   36
      Top             =   7380
      Width           =   3405
      Begin VB.CommandButton bPrintXM 
         Caption         =   "Print &EI XM Form"
         Height          =   645
         Index           =   1
         Left            =   1755
         Picture         =   "fPrintForm.frx":24A8
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   750
         Width           =   1410
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   810
         TabIndex        =   47
         Top             =   240
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         Value           =   72
         BuddyControl    =   "txtHoldFor"
         BuddyDispid     =   196621
         OrigLeft        =   840
         OrigTop         =   270
         OrigRight       =   1080
         OrigBottom      =   585
         Increment       =   12
         Max             =   96
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHoldFor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   46
         Text            =   "72"
         Top             =   300
         Width           =   405
      End
      Begin VB.CommandButton bPrintXM 
         Caption         =   "Print &XM Form"
         Height          =   645
         Index           =   0
         Left            =   330
         Picture         =   "fPrintForm.frx":2B12
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         Height          =   195
         Left            =   1110
         TabIndex        =   48
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Hold for"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   240
      TabIndex        =   29
      Top             =   7380
      Width           =   1695
      Begin VB.CommandButton bPrintGH 
         Caption         =   "Print &G/S Form"
         Height          =   645
         Left            =   240
         Picture         =   "fPrintForm.frx":317C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   750
         Width           =   1245
      End
      Begin MSComCtl2.UpDown udDays 
         Height          =   375
         Left            =   900
         TabIndex        =   30
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   56
         BuddyControl    =   "lblAvailableForDays"
         BuddyDispid     =   196627
         OrigLeft        =   810
         OrigTop         =   390
         OrigRight       =   1170
         OrigBottom      =   645
         Max             =   90
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Serum available"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label lblAvailableForDays 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "56"
         Height          =   255
         Left            =   390
         TabIndex        =   34
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "days"
         Height          =   195
         Left            =   1170
         TabIndex        =   33
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "for"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.CommandButton bSelect 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   2970
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Serum From Specimen Labelled"
      Height          =   2835
      Left            =   2655
      TabIndex        =   2
      Top             =   150
      Width           =   5535
      Begin VB.TextBox txtReceivedDate 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2460
         Width           =   1485
      End
      Begin VB.TextBox tAddr2 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1860
         Width           =   3915
      End
      Begin VB.TextBox tSex 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox tClinician 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   3915
      End
      Begin VB.TextBox taddr1 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   9
         Top             =   1560
         Width           =   3915
      End
      Begin VB.TextBox txtChart 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1260
         Width           =   1395
      End
      Begin VB.TextBox tgroup 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2160
         Width           =   1395
      End
      Begin VB.TextBox tdob 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3210
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1260
         Width           =   1155
      End
      Begin VB.TextBox tward 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         Top             =   600
         Width           =   3915
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         Top             =   300
         Width           =   3915
      End
      Begin VB.TextBox tSampleDate 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Received Date"
         Height          =   195
         Left            =   2520
         TabIndex        =   53
         Top             =   2490
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Clinican"
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Group/Rh"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   570
         TabIndex        =   15
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   1290
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   690
         TabIndex        =   13
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Chart No."
         Height          =   195
         Left            =   465
         TabIndex        =   11
         Top             =   1290
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date"
         Height          =   195
         Left            =   2670
         TabIndex        =   10
         Top             =   2220
         Width           =   915
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   2655
      TabIndex        =   28
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2655
      Left            =   240
      TabIndex        =   59
      Top             =   4560
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   7
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
      AllowUserResizing=   1
      FormatString    =   $"fPrintForm.frx":37E6
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   10005
      Picture         =   "fPrintForm.frx":3888
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblComment 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "The following Unit(s) is/are"
      Height          =   285
      Left            =   3420
      TabIndex        =   55
      Top             =   3780
      Width           =   2040
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   2700
      TabIndex        =   51
      Top             =   3840
      Width           =   660
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Typenex"
      Height          =   195
      Left            =   330
      TabIndex        =   45
      Top             =   660
      Width           =   615
   End
   Begin VB.Label lTypenex 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   990
      TabIndex        =   44
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "RED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   25
      Top             =   3600
      Width           =   825
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   600
      Picture         =   "fPrintForm.frx":3CCA
      Top             =   4020
      Width           =   480
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Units in             will Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   390
      TabIndex        =   26
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Antibodies"
      Height          =   195
      Index           =   0
      Left            =   2610
      TabIndex        =   24
      Top             =   3330
      Width           =   735
   End
   Begin VB.Label lAB 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3420
      TabIndex        =   23
      Top             =   3270
      Width           =   5115
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   300
      Width           =   870
   End
   Begin VB.Label lLabNumber 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   990
      TabIndex        =   17
      Top             =   270
      Width           =   1455
   End
   Begin VB.Label lOp 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   1
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Operator"
      Height          =   195
      Left            =   330
      TabIndex        =   0
      Top             =   1020
      Width           =   615
   End
End
Attribute VB_Name = "fPrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private mSampleID As String

Private optSerumDays As Integer
Private optHoldFor As Integer
Private optPlasma As Integer
Private Sub ClickCol0()
  
      Dim SaveX As Integer
      Dim SaveY As Integer
      Dim Product As String
      Dim Y As Integer

10    On Error GoTo ClickCol0_Error

20    SaveX = g.col
30    SaveY = g.row

40    g.col = 0
50    If g.CellBackColor = vbRed Then
60      g.CellBackColor = &H80000018
70      g.CellForeColor = &H8000000D
80    Else
90      Product = g.TextMatrix(g.row, 3)
100     g.CellBackColor = vbRed
110     g.CellForeColor = vbYellow
120     For Y = 1 To g.Rows - 1
130       If g.TextMatrix(Y, 3) <> Product Then
140         g.row = Y
150         g.CellBackColor = &H80000018
160         g.CellForeColor = &H8000000D
170       End If
180     Next
190   End If

200   g.row = SaveY
210   g.col = SaveX

220   Exit Sub

ClickCol0_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "fPrintForm", "ClickCol0", intEL, strES

End Sub

Private Sub FillcmbComment()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillcmbComment_Error

20    cmbComment.Clear
30    cmbComment.AddItem ""
      '"The following Unit(s) is/are "
'40    cmbComment.AddItem "Compatible with the Serum labelled as above"
'50    cmbComment.AddItem "Least Incompatible with Serum Labelled "
'60    cmbComment.AddItem "Issued For "

40    sql = "SELECT * FROM Lists WHERE " & _
            "ListType = 'XC' " & _
            "ORDER BY ListOrder"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    Do While Not tb.EOF
80      cmbComment.AddItem tb!Text & ""
90      tb.MoveNext
100   Loop

110   cmbComment = ""

120   Exit Sub

FillcmbComment_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "fPrintForm", "FillcmbComment", intEL, strES, sql

End Sub

Private Sub PrintXM(ByVal Index As Integer)

10    ReDim RowNumbersToPrint(1 To 1) As Integer
      Dim Y As Integer
      Dim n As Integer
      Dim HoldFor As Integer
      'Dim MaxHrs As String
      'Dim RNTP(1 To 4) As Integer
      Dim Comment As String
      
20    HoldFor = Val(txtHoldFor)

30    n = 0
40    g.col = 0
50    For Y = 1 To g.Rows - 1
60      g.row = Y
70      If g.CellBackColor = vbRed Then
80        n = n + 1
90        ReDim Preserve RowNumbersToPrint(1 To n) As Integer
100       RowNumbersToPrint(n) = Y
110     End If
120   Next
130   If n = 0 Then
140     iMsg "Select Pack Numbers to Print", vbExclamation
150     If TimedOut Then Unload Me: Exit Sub
160     Exit Sub
170   End If

180   CurrentReceivedDate = txtReceivedDate

190   Comment = ""
200   If cmbComment <> "" Then
210     Comment = lblComment & " " & cmbComment
220   End If
      
230   PrintXMForm lLabNumber, RowNumbersToPrint(), tSampleDate, HoldFor, Comment, Index

End Sub

Private Sub bPrintXM_Click(Index As Integer)

10    PrintXM (Index)

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub bPrintCord_Click()

10    CurrentReceivedDate = txtReceivedDate
20    PrintCordFormCavan lLabNumber, tSampleDate
30    UpdatePrinted lLabNumber, "Form"

End Sub

Private Sub bPrintGH_Click()

'260     With frmEligibility
'270       .Chart = txtChart
'280       .PatName = txtName
'290       .DoB = tDoB
'300       .Show 1
'310     End With

CurrentReceivedDate = txtReceivedDate
PrintGHFormCavan lLabNumber, Val(lblAvailableForDays), tSampleDate
UpdatePrinted lLabNumber, "Form"



End Sub



Private Sub bSelect_Click()

      Dim Y As Integer

10    g.col = 0

20    If bSelect.Caption = "Select All" Then
30      bSelect.Caption = "De-Select All"
40      For Y = 1 To g.Rows - 1
50        g.row = Y
60        g.CellBackColor = vbRed
70        g.CellForeColor = vbYellow
80      Next
90    Else
100     bSelect.Caption = "Select All"
110     For Y = 1 To g.Rows - 1
120       g.row = Y
130       g.CellBackColor = &H80000018
140       g.CellForeColor = &H8000000D
150     Next
160   End If

End Sub



Private Sub cmdPrintAN_Click()

10    CurrentReceivedDate = txtReceivedDate
20    PrintANFormCavan lLabNumber, tSampleDate
30    UpdatePrinted lLabNumber, "Form"

End Sub



Private Sub FillDetails()

      Dim Y As Integer
      Dim s As String
      Dim ok As Integer
      Dim issued As Integer
      Dim sql As String
      Dim tb As Recordset
      Dim Product As String
      Dim Multi As Boolean
      Dim BP As BatchProduct
      Dim BPs As New BatchProducts
      Dim Generic As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo FillDetails_Error

20    g.Font.Bold = True

30    sql = "Select * from PatientDetails where " & _
            "LabNumber = '" & mSampleID & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70      lTypenex = tb!Typenex & ""
80      Select Case tb!Sex & ""
          Case "M": tSex = "Male"
90        Case "F": tSex = "Female"
100       Case Else: tSex = ""
110     End Select
120     tClinician = tb!Clinician & ""
130     txtName = tb!Name & ""
140     tward = tb!Ward & ""
150     lOp = UserName
160     lAB = ""
170       If Trim$(tb!Anti3Reported & "") <> "" Then lAB = tb!Anti3Reported
180       If Trim$(tb!AIDS & "") <> "" Then lAB = tb!AIDS
190       If Trim$(tb!AIDR & "") <> "" Then lAB = tb!AIDR
200       If lAB = "Negative" Then
210         lAB = "No Atypical Antibodies detected."
220       End If
230     cmbComment = StripComment(tb!SampleComment & "")
240     txtChart = tb!Patnum
250     taddr1 = tb!Addr1
260     tAddr2 = tb!Addr2
270     tgroup = tb!fGroup & ""
280     tDoB = tb!DoB & ""

290     If Format(tb!DateReceived, "HH:mm") = "00:00" Then
300       txtReceivedDate = Format(tb!DateReceived, "dd/MM/yyyy")
310     Else
320       txtReceivedDate = Format(tb!DateReceived, "dd/MM/yyyy HH:mm")
330     End If
  
340     If Format(tb!SampleDate, "HH:mm") = "00:00" Then
350       tSampleDate = Format(tb!SampleDate, "dd/MM/yyyy")
360     Else
370       tSampleDate = Format(tb!SampleDate, "dd/MM/yyyy HH:mm")
380     End If

390     lLabNumber = mSampleID

400     Ps.LoadLatestBySampleID mSampleID
410     For Each p In Ps

420       If InStr("XIKYV", p.PackEvent) > 0 Then
430         ok = True: issued = False
440         If p.crtr Or p.ccor Or p.cenr Then ok = False
450         If Not ok Then
460           Answer = iMsg("Units are more compatible than the Patients Auto.", vbYesNo + vbQuestion)
470           If TimedOut Then Unload Me: Exit Sub
480           If Answer = vbYes Then
490             ok = True
500           End If
510         End If
520         If p.PackEvent = "I" Or p.PackEvent = "V" Then 'Issue or Electronic Issue
530           ok = True
540           issued = True
550         End If
560         If ok Or issued Then
570           s = p.ISBT128 & vbTab
580           s = s & Bar2Group(p.GroupRh) & vbTab
590           s = s & p.DateExpiry & vbTab
600           s = s & ProductWordingFor(p.BarCode) & vbTab
610           s = s & p.UserName & vbTab
620           s = s & p.RecordDateTime
630           g.AddItem s
640         End If
650       End If
660     Next
670   End If

680   BPs.LoadSampleIDNoAudit mSampleID
690   For Each BP In BPs
700     If BP.EventCode = "I" Then
710       s = BP.BatchNumber & vbTab & _
              BP.UnitGroup & vbTab & _
              Format$(BP.DateExpiry, "dd/mm/yy") & vbTab & _
              BP.Product & vbTab & _
              BP.UserName & vbTab & _
              Format$(BP.RecordDateTime, "dd/mm/yyyy hh:mm:ss")
720       g.AddItem s
730     End If
740   Next

750   If g.Rows > 2 Then
760     g.RemoveItem 1
770   End If

780   Multi = False
790   g.col = 0
800   Product = g.TextMatrix(1, 3)
810   For Y = 1 To g.Rows - 1
820     If g.TextMatrix(Y, 3) = Product Then
830       g.row = Y
840       g.CellBackColor = vbRed
850       g.CellForeColor = vbYellow
860     Else
870       Multi = True
880     End If
890   Next

900   g.Sort = 9

910   If Multi Then
920     bSelect.Visible = False
930   End If

940   AdjustWording

950    For Y = 1 To g.Rows - 1
960      Generic = UCase$(g.TextMatrix(Y, 3))
970      If UCase(Generic) = "OCTAPLAS" Or UCase(Generic) = "LG OCTAPLAS" Or UCase(Generic) = "UNIPLAS" Or InStr(Generic, "PLASMA") Then
980        g.TextMatrix(Y, 6) = DateAdd("H", txtMaxHrs, g.TextMatrix(Y, 5))
990      Else
1000       Generic = UCase$(ProductGenericFor(ProductBarCodeFor(g.TextMatrix(Y, 3))))
1010       If Generic = "RED CELLS" Or Generic = "WHOLE BLOOD" Then
1020         g.TextMatrix(Y, 6) = DateAdd("H", txtHoldFor, tSampleDate)
1030       Else
1040         g.TextMatrix(Y, 6) = g.TextMatrix(Y, 2) '& " 23:59"
1050       End If
1060     End If
1070     If IsDate(g.TextMatrix(Y, 6)) Then
1080       g.TextMatrix(Y, 6) = Format$(MinDate(g.TextMatrix(Y, 6), g.TextMatrix(Y, 2)), "dd/MM/yyyy HH:nn")
1090     End If
1100   Next

1110  Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

1120  intEL = Erl
1130  strES = Err.Description
1140  LogError "fPrintForm", "FillDetails", intEL, strES

End Sub


Private Sub Form_Load()

10    FillcmbComment

20    chkPreview.Visible = False
30    If blnPrintingWithPreview Then
40      chkPreview.Visible = True
50    End If

60      chkPreview.Visible = False
70      bPrintCord.Visible = True
80      cmdPrintAN.Visible = True
90      txtHoldFor = "24"
  
100   lblAvailableForDays = GetOptionSetting("TransfusionAvailableForDays", "56")
110   optSerumDays = Val(lblAvailableForDays)

120   txtHoldFor = GetOptionSetting("TransfusionHoldFor", "72")
130   optHoldFor = Val(txtHoldFor)

140   txtMaxHrs = GetOptionSetting("TransfusionPlasmaHours", "4")
150   optPlasma = Val(txtMaxHrs)

      '*****NOTE
          'FillDetail might be dependent on many components so for any future
          'update in code try to keep FillDetails on bottom most line of form load.
160       FillDetails
      '**************************************

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If optSerumDays <> Val(lblAvailableForDays) Then

30      Answer = iMsg("Do you want to set the default" & vbCrLf & _
                "Serum Available value to " & lblAvailableForDays & _
                " days?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbYes Then
60        SaveOptionSetting "TransfusionAvailableForDays", lblAvailableForDays
70      End If

80    End If

90    If optHoldFor <> Val(txtHoldFor) Then

100     Answer = iMsg("Do you want to set the default" & vbCrLf & _
                "'Hold For' value to " & txtHoldFor & _
                " hours?", vbQuestion + vbYesNo)
110     If TimedOut Then Unload Me: Exit Sub
120     If Answer = vbYes Then
130       SaveOptionSetting "TransfusionHoldFor", txtHoldFor
140     End If

150   End If

160   If optPlasma <> Val(txtMaxHrs) Then

170     Answer = iMsg("Do you want to set the default" & vbCrLf & _
                "'Use Plasma within' value to " & txtMaxHrs & _
                " hours?", vbQuestion + vbYesNo)
180     If TimedOut Then Unload Me: Exit Sub
190     If Answer = vbYes Then
200       SaveOptionSetting "TransfusionPlasmaHours", txtMaxHrs
210     End If

220   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub


Private Sub g_Click()

      Dim n As Integer
      Dim SaveX As Integer
      Dim f As Form
      Dim NewTime As String

10    On Error GoTo g_Click_Error

20    If g.MouseRow < 1 Then Exit Sub

30    SaveX = g.MouseCol

40    ClickCol0

50    If SaveX = 6 Then
60      g.col = 0
70      g.CellBackColor = vbRed
80      g.CellForeColor = vbYellow
  
90      Set f = New frmAskDateTime
100     With f
110       .DateTime = g.TextMatrix(g.row, SaveX)
120       .Prompt = "Enter 'Do not commence Transfusion' time"
130       .Show 1
140       If IsDate(.DateTime) Then
150         NewTime = Format(.DateTime, "dd/MM/yyyy HH:nn")
160       Else
170         NewTime = Format$(Now, "dd/MM/yyyy HH:nn")
180       End If
190     End With
200     Unload f
210     Set f = Nothing
  
220     NewTime = Format$(MinDate(NewTime, g.TextMatrix(g.row, 2)), "dd/MM/yyyy HH:nn")

230     g.col = 0
240     For n = 1 To g.Rows - 1
250       g.row = n
260       If g.CellBackColor = vbRed Then
270         g.TextMatrix(n, 6) = NewTime
280       End If
290     Next

300   End If

310   AdjustWording

320   Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "fPrintForm", "g_Click", intEL, strES

End Sub


Private Sub AdjustWording()

      Dim n As Integer
      Dim Counter As Integer

10    Counter = 0
20    g.col = 0
30    For n = 1 To g.Rows - 1
40      g.row = n
50      If g.CellBackColor = vbRed Then
60        Counter = Counter + 1
70      End If
80    Next

90    If Counter = 0 Then
100     lblComment = ""
110     cmbComment = ""
120   ElseIf Counter = 1 Then
130     lblComment = "The following Unit is "
'140     cmbComment = "Compatible with Serum Labelled "
140   Else
150     lblComment = "The following Units are "
'170     cmbComment = "Compatible with Serum Labelled "
160   End If

End Sub

Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    d1 = Format(g.TextMatrix(Row1, 2), "dd/mmm/yyyy hh:mm:ss")
20    d2 = Format(g.TextMatrix(Row2, 2), "dd/mmm/yyyy hh:mm:ss")

30    Cmp = Sgn(DateDiff("D", d2, d1))

End Sub


