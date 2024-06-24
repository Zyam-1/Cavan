VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCreatClearance 
   Caption         =   "NetAcquire - Creatinine Clearance"
   ClientHeight    =   6705
   ClientLeft      =   675
   ClientTop       =   915
   ClientWidth     =   11850
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6705
   ScaleWidth      =   11850
   Begin VB.CommandButton bprinturine 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12300
      Picture         =   "frmCreatClearance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5310
      Width           =   825
   End
   Begin VB.Frame Frame4 
      Height          =   1965
      Left            =   180
      TabIndex        =   33
      Top             =   4440
      Width           =   6525
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Serum Creatinine"
         Height          =   195
         Left            =   1410
         TabIndex        =   48
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Urine Creatinine"
         Height          =   195
         Left            =   1485
         TabIndex        =   47
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Urinary Protein"
         Height          =   195
         Left            =   1575
         TabIndex        =   46
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Creatinine Clearance"
         Height          =   195
         Left            =   1140
         TabIndex        =   45
         Top             =   1380
         Width           =   1470
      End
      Begin VB.Label lsc 
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
         Left            =   2880
         TabIndex        =   44
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label luc 
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
         Left            =   2880
         TabIndex        =   43
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label lup 
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
         Left            =   2880
         TabIndex        =   42
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lcc 
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
         Left            =   2880
         TabIndex        =   41
         Top             =   1380
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Urinary Protein"
         Height          =   195
         Left            =   1575
         TabIndex        =   40
         Top             =   1110
         Width           =   1035
      End
      Begin VB.Label lupc 
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
         Left            =   2880
         TabIndex        =   39
         Top             =   1110
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "umol/L"
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   38
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "umol/L"
         Height          =   195
         Index           =   1
         Left            =   3930
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "g/L"
         Height          =   195
         Left            =   3930
         TabIndex        =   36
         Top             =   870
         Width           =   255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "g/24Hr"
         Height          =   195
         Left            =   3930
         TabIndex        =   35
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "mL/min"
         Height          =   195
         Left            =   3930
         TabIndex        =   34
         Top             =   1410
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serum Details"
      Height          =   2835
      Left            =   7080
      TabIndex        =   13
      Top             =   210
      Width           =   4605
      Begin VB.CommandButton cmdViewScanSerum 
         Caption         =   "&View Scan"
         Height          =   1020
         Left            =   3420
         Picture         =   "frmCreatClearance.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1800
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdSetPrinter 
         Height          =   585
         Left            =   2790
         Picture         =   "frmCreatClearance.frx":5E58
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
         Top             =   930
         Width           =   735
      End
      Begin VB.CommandButton bprintserum 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3630
         Picture         =   "frmCreatClearance.frx":6162
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   930
         Width           =   735
      End
      Begin VB.Label lblSerumSID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3060
         TabIndex        =   51
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SID"
         Height          =   195
         Index           =   0
         Left            =   2730
         TabIndex        =   49
         Top             =   300
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Run Date"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   22
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   20
         Top             =   1260
         Width           =   315
      End
      Begin VB.Label ldob 
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
         Index           =   0
         Left            =   840
         TabIndex        =   19
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lchart 
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
         Index           =   0
         Left            =   840
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lname 
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
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   630
         Width           =   3525
      End
      Begin VB.Label lcomment 
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
         Index           =   0
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   3525
      End
      Begin VB.Label lserumdate 
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
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Urine Details"
      Height          =   2775
      Left            =   7080
      TabIndex        =   12
      Top             =   3060
      Width           =   4605
      Begin VB.CommandButton cmdViewScanSerumUrine 
         Caption         =   "&View Scan"
         Height          =   1020
         Left            =   3420
         Picture         =   "frmCreatClearance.frx":67CC
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1740
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblUrineSID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3090
         TabIndex        =   52
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SID"
         Height          =   195
         Index           =   1
         Left            =   2730
         TabIndex        =   50
         Top             =   300
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Run Date"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   330
         Width           =   690
      End
      Begin VB.Label lname 
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
         Index           =   1
         Left            =   870
         TabIndex        =   31
         Top             =   630
         Width           =   3525
      End
      Begin VB.Label lchart 
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
         Index           =   1
         Left            =   870
         TabIndex        =   30
         Top             =   930
         Width           =   1695
      End
      Begin VB.Label ldob 
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
         Index           =   1
         Left            =   870
         TabIndex        =   29
         Top             =   1230
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   28
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   27
         Top             =   930
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   26
         Top             =   630
         Width           =   420
      End
      Begin VB.Label lcomment 
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
         Index           =   1
         Left            =   870
         TabIndex        =   25
         Top             =   1530
         Width           =   3525
      End
      Begin VB.Label lurinedate 
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
         Left            =   870
         TabIndex        =   24
         Top             =   300
         Width           =   1695
      End
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
      Height          =   1995
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Width           =   6525
      Begin VB.TextBox tvolume 
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
         Left            =   1710
         TabIndex        =   9
         Top             =   1260
         Width           =   1425
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Search"
         Height          =   525
         Left            =   5100
         TabIndex        =   4
         Top             =   630
         Width           =   1095
      End
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
         TabIndex        =   3
         Top             =   840
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   3180
         TabIndex        =   5
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218890241
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         Top             =   420
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218890241
         CurrentDate     =   37505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Urinary Volume"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1320
         Width           =   1470
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "mL"
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Run Dates Between"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Patient Name"
         Height          =   195
         Left            =   660
         TabIndex        =   7
         Top             =   900
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   8700
      Picture         =   "frmCreatClearance.frx":BFBA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1545
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2025
      Left            =   180
      TabIndex        =   1
      Top             =   2310
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3572
      _Version        =   393216
      Cols            =   4
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
      FormatString    =   "<Patient Name                           |<Run Date            |<Serum #              |<Urine #              "
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Print Urine >>"
      Height          =   195
      Left            =   10710
      TabIndex        =   54
      Top             =   6120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuCode 
         Caption         =   "Serum &Creatinine"
         Index           =   0
      End
      Begin VB.Menu mnuCode 
         Caption         =   "&Urine Creatinine"
         Index           =   1
      End
      Begin VB.Menu mnuCode 
         Caption         =   "Urine &Protein"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCreatClearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private SCreaCode As String
Private UCreaCode As String
Private UProCode As String

Private pPrintToPrinter As String

Public Property Get PrintToPrinter() As String

26790     PrintToPrinter = pPrintToPrinter

End Property

Public Property Let PrintToPrinter(ByVal strNewValue As String)

26800     pPrintToPrinter = strNewValue

End Property

Private Sub ClearAll()

26810     lserumdate = ""
26820     lname(0) = ""
26830     lchart(0) = ""
26840     ldob(0) = ""
26850     lcomment(0) = ""

26860     lurinedate = ""
26870     lname(1) = ""
26880     lchart(1) = ""
26890     ldob(1) = ""
26900     lcomment(1) = ""

26910     lsc = ""
26920     lup = ""
26930     lupc = ""
26940     lcc = ""
26950     luc = ""

26960     lblSerumSID = ""
26970     lblUrineSID = ""

26980     tvolume = ""

End Sub

Private Sub SaveCreat()

          Dim tb As Recordset
          Dim sql As String

26990     On Error GoTo SaveCreat_Error

27000     sql = "Select * from Creatinine where " & _
              "SerumNumber = '" & lblSerumSID & "' " & _
              "and UrineNumber = '" & lblUrineSID & "'"
27010     Set tb = New Recordset
27020     RecOpenServer 0, tb, sql

27030     If tb.EOF Then
27040         tb.AddNew
27050     End If
27060     tb!serumnumber = lblSerumSID
27070     tb!urinenumber = lblUrineSID
27080     tb!urinevolume = tvolume
27090     tb!serumcreat = lsc
27100     tb!urinecreat = luc
27110     tb!urineprol = lup
27120     tb!urinepro24hr = lupc
27130     tb!ccl = lcc
27140     tb!Name = lname(1)
27150     tb!Chart = lchart(1)
27160     If IsDate(ldob(1)) Then tb!DoB = ldob(1)
27170     tb!Comment = Left$(lcomment(1), 30)
27180     tb.Update

27190     Exit Sub

SaveCreat_Error:

          Dim strES As String
          Dim intEL As Integer

27200     intEL = Erl
27210     strES = Err.Description
27220     LogError "fCreatClear", "SaveCreat", intEL, strES, sql


End Sub

Private Function GetCodes() As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim strItem As String
          Dim n As Integer
          Dim strItemCode(1 To 3) As String
          Dim blnRet As Boolean

27230     On Error GoTo GetCodes_Error

27240     For n = 1 To 3
27250         strItemCode(n) = ""
27260         strItem = Choose(n, "BioCodeForCreat", "BioCodeForUCreat", "BioCodeForUProt")

27270         sql = "Select * from Options where " & _
                  "Description = '" & strItem & "'"
27280         Set tb = New Recordset
27290         RecOpenServer 0, tb, sql
27300         If Not tb.EOF Then
27310             strItemCode(n) = Trim$(tb!Contents & "")
27320         End If
27330     Next

27340     SCreaCode = strItemCode(1)
27350     UCreaCode = strItemCode(2)
27360     UProCode = strItemCode(3)

27370     blnRet = True
27380     For n = 1 To 3
27390         If strItemCode(n) = "" Then
27400             blnRet = False
27410         End If
27420     Next
27430     GetCodes = blnRet

27440     Exit Function

GetCodes_Error:

          Dim strES As String
          Dim intEL As Integer

27450     intEL = Erl
27460     strES = Err.Description
27470     LogError "fCreatClear", "GetCodes", intEL, strES, sql


End Function

Private Sub SetCode(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim strItem As String
          Dim strNewCode As String
          Dim strPrompt As String
          Dim strCurrent As String

27480     On Error GoTo SetCode_Error

27490     strItem = Choose(Index + 1, "BioCodeForCreat", "BioCodeForUCreat", "BioCodeForUProt")
27500     strPrompt = Choose(Index + 1, "Serum Creatinine", "Urine Creatinine", "Urine Protein")
27510     strCurrent = Choose(Index + 1, SCreaCode, UCreaCode, UProCode)

27520     strNewCode = iBOX("Enter the Analyser Code " & vbCrLf & _
              "for " & strPrompt & ".", _
              "Enter Code", strCurrent)
          
27530     If Trim$(strNewCode) <> "" Then
27540         sql = "Select * from Options where " & _
                  "Description = '" & strItem & "'"
27550         Set tb = New Recordset
27560         RecOpenServer 0, tb, sql
27570         If tb.EOF Then tb.AddNew
27580         tb!Description = strItem
27590         tb!Contents = strNewCode
27600         tb.Update
27610     End If

27620     GetCodes

27630     Exit Sub

SetCode_Error:

          Dim strES As String
          Dim intEL As Integer

27640     intEL = Erl
27650     strES = Err.Description
27660     LogError "fCreatClear", "SetCode", intEL, strES, sql


End Sub



Private Sub cmdCancel_Click()

27670     Unload Me

End Sub

Private Sub bPrintSerum_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Ward As String
          Dim Clin As String
          Dim GP As String

27680     On Error GoTo bPrintSerum_Click_Error

27690     GetWardClinGP lblSerumSID, Ward, Clin, GP

27700     SaveCreat
        
27710     sql = "Select * from PrintPending where " & _
              "Department = 'T' " & _
              "and SampleID = '" & lblSerumSID & "'"
        
27720     Set tb = New Recordset
27730     RecOpenClient 0, tb, sql
27740     If tb.EOF Then
27750         tb.AddNew
27760     End If
27770     tb!SampleID = lblSerumSID
27780     tb!Ward = Ward
27790     tb!Clinician = Clin
27800     tb!GP = GP
27810     tb!Department = "T"
27820     tb!Initiator = UserName
27830     tb!UsePrinter = pPrintToPrinter
27840     tb.Update

27850     Exit Sub

bPrintSerum_Click_Error:

          Dim strES As String
          Dim intEL As Integer

27860     intEL = Erl
27870     strES = Err.Description
27880     LogError "fCreatClear", "bPrintSerum_Click", intEL, strES, sql

End Sub

Private Sub bprinturine_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Ward As String
          Dim Clin As String
          Dim GP As String

27890     On Error GoTo bprinturine_Click_Error

27900     GetWardClinGP lblUrineSID, Ward, Clin, GP

27910     SaveCreat
        
27920     sql = "Select * from PrintPending where " & _
              "Department = 'R' " & _
              "and SampleID = '" & lblUrineSID & "'"
        
27930     Set tb = New Recordset
27940     RecOpenClient 0, tb, sql
27950     If tb.EOF Then
27960         tb.AddNew
27970     End If
27980     tb!SampleID = lblUrineSID
27990     tb!Ward = Ward
28000     tb!Clinician = Clin
28010     tb!GP = GP
28020     tb!Department = "R"
28030     tb!Initiator = UserName
28040     tb.Update

28050     Exit Sub

bprinturine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

28060     intEL = Erl
28070     strES = Err.Description
28080     LogError "fCreatClear", "bprinturine_Click", intEL, strES, sql


End Sub

Private Sub Calculate()

28090     On Error GoTo Calculate_Error

28100     If Val(lup) <> 0 And Val(tvolume) <> 0 Then
28110         lupc = Format$(Val(lup) * Val(tvolume) / 1000, "0.000")
28120     End If

28130     If Val(luc) <> 0 And Val(lsc) <> 0 And Val(tvolume) <> 0 Then
28140         If Val(luc) > 100 Then
28150             lcc = Format$((Val(luc) * Val(tvolume)) / (Val(lsc) * 1440), "##0")
28160         Else
28170             lcc = Format$((Val(luc) * 1000 * Val(tvolume)) / (Val(lsc) * 1440), "##0")
28180         End If
28190     End If

28200     Exit Sub

Calculate_Error:

          Dim strES As String
          Dim intEL As Integer

28210     intEL = Erl
28220     strES = Err.Description
28230     LogError "fCreatClear", "Calculate", intEL, strES

End Sub

Private Sub cmdRefresh_Click()

28240     ClearAll
28250     FillG

End Sub

Private Sub FillDetailsSerum(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String

28260     On Error GoTo FillDetailsSerum_Error

28270     lblSerumSID = SampleID

28280     sql = "select * from demographics where " & _
              "SampleID = '" & SampleID & "'"
28290     Set tb = New Recordset
28300     RecOpenServer 0, tb, sql
28310     If tb.EOF Then
28320         lserumdate = ""
28330         lname(0) = ""
28340         lchart(0) = ""
28350         ldob(0) = ""
28360         lcomment(0) = ""
28370     Else
28380         lserumdate = tb!Rundate
28390         lchart(0) = tb!Chart & ""
28400         lname(0) = tb!PatName & ""
28410         ldob(0) = tb!DoB & ""
              '  lcomment(0) = tb!biocomment0 & ""
28420     End If

28430     sql = "select * from bioresults where " & _
              "SampleID = '" & SampleID & "' " & _
              "and Code = '" & SCreaCode & "'"
28440     Set tb = New Recordset
28450     RecOpenServer 0, tb, sql
28460     If Not tb.EOF Then
28470         If Not IsNull(tb!Result) Then
28480             lsc = Format$(tb!Result, "####.0")
28490         Else
28500             lsc = ""
28510         End If
28520     Else
28530         lsc = ""
28540     End If

28550     Calculate
28560     SetViewScans lblSerumSID, cmdViewScanSerum
28570     Exit Sub

FillDetailsSerum_Error:

          Dim strES As String
          Dim intEL As Integer

28580     intEL = Erl
28590     strES = Err.Description
28600     LogError "fCreatClear", "FillDetailsSerum", intEL, strES, sql

End Sub

Private Sub cmdSetPrinter_Click()

28610     frmForcePrinter.From = Me
28620     frmForcePrinter.Show 1

28630     If pPrintToPrinter = "Automatic Selection" Then
28640         pPrintToPrinter = ""
28650     End If

28660     If pPrintToPrinter <> "" Then
28670         cmdSetPrinter.BackColor = vbRed
28680         cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
28690     Else
28700         cmdSetPrinter.BackColor = vbButtonFace
28710         pPrintToPrinter = ""
28720         cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
28730     End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdViewScanSerum_Click
' Author    : Masood
' Date      : 02/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewScanSerum_Click()
28740     On Error GoTo cmdViewScanSerum_Click_Error


28750     frmViewScan.SampleID = lblSerumSID
28760     frmViewScan.txtSampleID = lblSerumSID
28770     frmViewScan.Show 1

       
28780     Exit Sub

       
cmdViewScanSerum_Click_Error:

          Dim strES As String
          Dim intEL As Integer

28790     intEL = Erl
28800     strES = Err.Description
28810     LogError "frmCreatClearance", "cmdViewScanSerum_Click", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdViewScanSerumUrine_Click
' Author    : Masood
' Date      : 02/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewScanSerumUrine_Click()

28820     On Error GoTo cmdViewScanSerumUrine_Click_Error


28830     frmViewScan.SampleID = lblUrineSID
28840     frmViewScan.txtSampleID = lblUrineSID
28850     frmViewScan.Show 1

       
28860     Exit Sub

       
cmdViewScanSerumUrine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

28870     intEL = Erl
28880     strES = Err.Description
28890     LogError "frmCreatClearance", "cmdViewScanSerumUrine_Click", intEL, strES

End Sub

Private Sub dtFrom_CloseUp()

28900     cmdRefresh.Visible = True

End Sub


Private Sub dtTo_CloseUp()

28910     cmdRefresh.Visible = True

End Sub


Private Sub Form_Load()

28920     dtFrom = Format$(Now - 30, "dd/mm/yyyy")
28930     dtTo = Format$(Now, "dd/mm/yyyy")

28940     If Not GetCodes() Then
28950         iMsg "Analyte Codes are not set.", vbCritical
28960     End If

End Sub

Private Sub FillDetailsUrine(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String

28970     On Error GoTo FillDetailsUrine_Error

28980     lblUrineSID = SampleID

28990     sql = "select * from demographics where " & _
              "SampleID = '" & SampleID & "'"
29000     Set tb = New Recordset
29010     RecOpenServer 0, tb, sql
29020     If tb.EOF Then
29030         lurinedate = ""
29040         lname(1) = ""
29050         lchart(1) = ""
29060         ldob(1) = ""
29070         lcomment(1) = ""
29080     Else
29090         lurinedate = tb!Rundate
29100         lchart(1) = tb!Chart & ""
29110         lname(1) = tb!PatName & ""
29120         ldob(1) = tb!DoB & ""
              '  lcomment(1) = tb!biocomment0 & ""
29130     End If

29140     sql = "select * from bioresults where " & _
              "SampleID = '" & SampleID & "' " & _
              "and Code = '" & UCreaCode & "'"
29150     Set tb = New Recordset
29160     RecOpenServer 0, tb, sql
29170     If Not tb.EOF Then
29180         If Not IsNull(tb!Result) Then
29190             luc = Format$(tb!Result, "#0.00")
29200         Else
29210             luc = ""
29220         End If
29230     Else
29240         luc = ""
29250     End If

29260     sql = "select * from bioresults where " & _
              "SampleID = '" & SampleID & "' " & _
              "and code = '" & UProCode & "'"
29270     Set tb = New Recordset
29280     RecOpenServer 0, tb, sql
29290     If Not tb.EOF Then
29300         If Not IsNull(tb!Result) Then
29310             lup = Format$(tb!Result, "#0.00")
29320         Else
29330             lup = ""
29340         End If
29350     Else
29360         lup = ""
29370     End If

29380     sql = "Select Result from BioResults where SampleID = '" & SampleID & "' " & "and Code = 'TUV'"
29390     Set tb = New Recordset
29400     RecOpenServer 0, tb, sql
29410     If Not tb.EOF Then
29420         tvolume = tb!Result & ""
29430     Else
29440         tvolume = ""
29450     End If

29460     Calculate
29470     SetViewScans lblUrineSID, cmdViewScanSerumUrine
29480     Exit Sub

FillDetailsUrine_Error:

          Dim strES As String
          Dim intEL As Integer

29490     intEL = Erl
29500     strES = Err.Description
29510     LogError "fCreatClear", "FillDetailsUrine", intEL, strES, sql

End Sub

Private Sub g_Click()

          Dim ySave As Integer
          Dim n As Integer

29520     cmdViewScanSerum.Visible = False
29530     cmdViewScanSerumUrine.Visible = False


29540     If g.MouseRow = 0 Then Exit Sub

29550     If g.TextMatrix(g.row, 0) = "" Then Exit Sub

29560     ySave = g.row

29570     If g.TextMatrix(ySave, 2) <> "" Then
29580         g.Col = 2
29590     Else
29600         g.Col = 3
29610     End If
29620     For n = 1 To g.Rows - 1
29630         g.row = n
29640         g.CellBackColor = 0
29650     Next
29660     g.row = ySave
29670     g.CellBackColor = vbRed
        
29680     If g.Col = 2 Then
29690         FillDetailsSerum g.TextMatrix(ySave, 2)
29700     Else
29710         FillDetailsUrine g.TextMatrix(ySave, 3)
29720     End If
        
End Sub

Private Sub mnuCode_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

29730     On Error GoTo mnuCode_Click_Error

29740     sql = "Select MemberOf from Users where " & _
              "Name = '" & UserName & "' " & _
              "and InUse = 1"
29750     Set tb = New Recordset
29760     RecOpenServer 0, tb, sql
29770     If Not tb.EOF Then
29780         If tb!MemberOf = "Managers" Then
29790             SetCode Index
29800         Else
29810             iMsg "Only Lab Managers can change Codes!", vbCritical
29820         End If
29830     End If

29840     Exit Sub

mnuCode_Click_Error:

          Dim strES As String
          Dim intEL As Integer

29850     intEL = Erl
29860     strES = Err.Description
29870     LogError "fCreatClear", "mnuCode_Click", intEL, strES, sql


End Sub

Private Sub mnuExit_Click()

29880     Unload Me

End Sub

Private Sub tVolume_LostFocus()

29890     Calculate

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim ThisName As String
          Dim Y As Integer
          Dim s As Boolean
          Dim U As Boolean
          Dim n As Integer

29900     On Error GoTo FillG_Error

29910     g.Rows = 2
29920     g.AddItem ""
29930     g.RemoveItem 1

29940     If Trim$(txtName) = "" Then Exit Sub

29950     Screen.MousePointer = vbHourglass

29960     sql = "SELECT DISTINCT D.SampleID, PatName, B.RunDate " & _
              "FROM Demographics AS D, BioResults AS B WHERE " & _
              "D.PatName LIKE '" & AddTicks(txtName) & "%' " & _
              "AND D.SampleID = B.SampleID " & _
              "AND B.RunDate BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "AND '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
              "AND B.Code = '" & SCreaCode & "'"

29970     Set tb = New Recordset
29980     RecOpenServer 0, tb, sql
29990     Do While Not tb.EOF
30000         g.AddItem tb!PatName & vbTab & tb!Rundate & vbTab & tb!SampleID
30010         tb.MoveNext
30020     Loop

30030     sql = "SELECT DISTINCT D.SampleID, PatName, B.RunDate " & _
              "FROM Demographics AS D, BioResults AS B WHERE " & _
              "D.PatName LIKE '" & AddTicks(txtName) & "%' " & _
              "AND D.SampleID = B.SampleID " & _
              "AND B.RunDate BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "AND '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
              "AND (B.Code = '" & UCreaCode & "' or B.Code = '" & UProCode & "')"
30040     Set tb = New Recordset
30050     RecOpenServer 0, tb, sql
30060     Do While Not tb.EOF
30070         g.AddItem tb!PatName & vbTab & tb!Rundate & vbTab & vbTab & tb!SampleID
30080         tb.MoveNext
30090     Loop

30100     If g.Rows > 2 Then
30110         g.RemoveItem 1
30120     End If

30130     For Y = g.Rows - 1 To 1 Step -1
30140         s = False
30150         U = False
30160         ThisName = g.TextMatrix(Y, 0)
30170         If g.TextMatrix(Y, 2) <> "" Then s = True
30180         If g.TextMatrix(Y, 3) <> "" Then U = True
30190         For n = g.Rows - 1 To 1 Step -1
30200             If UCase$(g.TextMatrix(n, 0)) = UCase$(ThisName) Then
30210                 If g.TextMatrix(n, 2) <> "" Then s = True
30220                 If g.TextMatrix(n, 3) <> "" Then U = True
30230             End If
30240             If s And U Then
30250                 Exit For
30260             End If
30270         Next
30280         If Not (s And U) Then
30290             If Y = 1 Then g.AddItem ""
30300             g.RemoveItem Y
30310         End If
30320     Next

30330     Screen.MousePointer = 0

30340     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

30350     intEL = Erl
30360     strES = Err.Description
30370     LogError "fCreatClear", "FillG", intEL, strES, sql


End Sub



