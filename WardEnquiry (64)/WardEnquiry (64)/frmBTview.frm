VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBTview 
   Caption         =   "Blood Transfusion"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraProductHistory 
      Caption         =   "Product History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   60
      TabIndex        =   29
      Top             =   5550
      Width           =   13590
      Begin MSFlexGridLib.MSFlexGrid grdProd 
         Height          =   3465
         Left            =   30
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   300
         Width           =   13500
         _ExtentX        =   23813
         _ExtentY        =   6112
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   $"frmBTview.frx":0000
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
   End
   Begin VB.Frame Frame 
      Caption         =   "Kleihauer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   60
      TabIndex        =   27
      Top             =   2880
      Width           =   13590
      Begin MSFlexGridLib.MSFlexGrid grdKle 
         Height          =   1980
         Left            =   165
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   345
         Width           =   13170
         _ExtentX        =   23230
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   4
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
         FormatString    =   $"frmBTview.frx":00C4
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
   End
   Begin VB.Frame Frame 
      Height          =   2475
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   210
      Width           =   13590
      Begin VB.TextBox txtChart 
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
         Left            =   5145
         MaxLength       =   10
         TabIndex        =   36
         Top             =   210
         Width           =   1155
      End
      Begin VB.TextBox tident 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   8460
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   210
         Width           =   4815
      End
      Begin VB.TextBox tDoB 
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
         Left            =   5145
         MaxLength       =   10
         TabIndex        =   10
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox tAge 
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
         Left            =   5145
         MaxLength       =   4
         TabIndex        =   9
         Top             =   840
         Width           =   1125
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
         Height          =   285
         Left            =   915
         MaxLength       =   50
         TabIndex        =   8
         Top             =   210
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Index           =   0
         Left            =   915
         MaxLength       =   50
         TabIndex        =   7
         Top             =   825
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Index           =   1
         Left            =   915
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1110
         Width           =   3615
      End
      Begin VB.TextBox tMaiden 
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
         Left            =   915
         MaxLength       =   20
         TabIndex        =   5
         Top             =   525
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Index           =   2
         Left            =   915
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1395
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Index           =   3
         Left            =   915
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1695
         Width           =   3615
      End
      Begin VB.TextBox tComment 
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
         Left            =   915
         TabIndex        =   2
         Top             =   2070
         Width           =   12360
      End
      Begin VB.Label lblChartTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MRN"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4740
         TabIndex        =   37
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Sample Date"
         Height          =   255
         Left            =   7470
         TabIndex        =   35
         Top             =   1695
         Width           =   945
      End
      Begin VB.Label lblSampleDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8460
         TabIndex        =   34
         Top             =   1695
         Width           =   1905
      End
      Begin VB.Label lGroup 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   5145
         TabIndex        =   33
         Top             =   1650
         Width           =   1125
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4665
         TabIndex        =   26
         Top             =   1695
         Width           =   435
      End
      Begin VB.Label lblDCTresult 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   8460
         TabIndex        =   25
         Top             =   1110
         Width           =   1905
      End
      Begin VB.Label lblDCT 
         Caption         =   "DCT"
         Height          =   255
         Left            =   8010
         TabIndex        =   24
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "AB Report"
         Height          =   195
         Left            =   7650
         TabIndex        =   23
         Top             =   210
         Width           =   735
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   4695
         TabIndex        =   21
         Top             =   555
         Width           =   450
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   4800
         TabIndex        =   20
         Top             =   870
         Width           =   285
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   4815
         TabIndex        =   19
         Top             =   1170
         Width           =   270
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   420
         TabIndex        =   18
         Top             =   255
         Width           =   420
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Addr 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   465
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   210
         TabIndex        =   16
         Top             =   2085
         Width           =   630
      End
      Begin VB.Label lblMaiden 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "M.Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   540
         Width           =   600
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   735
         TabIndex        =   14
         Top             =   1155
         Width           =   90
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   735
         TabIndex        =   13
         Top             =   1410
         Width           =   90
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   735
         TabIndex        =   12
         Top             =   1710
         Width           =   90
      End
      Begin VB.Label lSex 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Left            =   5130
         TabIndex        =   11
         Top             =   1140
         Width           =   1125
      End
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   12300
      Picture         =   "frmBTview.frx":0185
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9585
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3855
      Left            =   13995
      TabIndex        =   31
      Top             =   5985
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   6800
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
      FormatString    =   "<Chart    |<Date Of Birth     |<Name                           "
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
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   45
      TabIndex        =   32
      Top             =   0
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmBTview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnFormActivated As Boolean

Private pRecentChart As String
Private pRecentPatName As String
Private pRecentDoB As String


Public Property Let RecentChart(ByVal strChart As String)

10  pRecentChart = strChart

End Property

Public Property Let RecentPatName(ByVal strPatName As String)

10  pRecentPatName = strPatName

End Property


Public Property Let RecentDoB(ByVal strDOB As String)

10  pRecentDoB = strDOB

End Property



Private Sub bcancel_Click()
10  Unload Me
End Sub
Private Sub FillPatientDetails()

    Dim sql As String
    Dim tb As Recordset
    Dim n As Integer
    Dim S As String
    Dim ChartList As String
    Dim PatNameList As String
    Dim DoBList As String
    Dim WhereClause As String

10  On Error GoTo FillPatientDetails_Error

20  With grd
30      sql = "Select * from PatientDetails where PatNum='" & pRecentChart & "' and DOB = '" & Format(pRecentDoB, "dd/MMM/yyyy") & "' and Name= '" & AddTicks(pRecentPatName) & "' order by SampleDate desc"
40  End With

50  Set tb = New Recordset
60  RecOpenServerBB 0, tb, sql
70  With tb
80      If Not .EOF Then
90          txtName = !Name & ""
100         tMaiden = !maiden & ""
110         tAddr(0) = !Addr1 & ""
120         tAddr(1) = !Addr2 & ""
130         tAddr(2) = !Addr3 & ""
140         tAddr(3) = !addr4 & ""
150         txtChart = !Patnum & ""
160         tComment = getExternalNotes(txtChart)
170         tComment.ToolTipText = !Comment & ""
180         tDoB = !DoB & ""
190         tAge = IIf(DateDiff("d", Format(!DoB, "dd/MMM/yyyy"), Format(!DateTime, "dd/MMM/yyyy")) > 365.25, (DateDiff("d", Format(!DoB, "dd/MMM/yyyy"), Format(!DateTime, "dd/MMM/yyyy")) \ 365.25), DateDiff("d", Format(!DoB, "dd/MMMM/yyyy"), Format(!DateTime, "dd/MMM/yyyy")))
200         tident = !AIDR & ""
210         If UCase(Trim$(!Sex & "")) = "M" Then
220             lSex.Caption = "Male"
230         ElseIf UCase(Trim$(!Sex & "")) = "F" Then
240             lSex.Caption = "Female"
250         End If
260         lGroup = !fGroup
            'If any DAT result is Positive then overall DCT is Positive otherwise Negative
270         If !DAT0 Or !dat2 Or !dat4 Or !dat6 Or !dat8 Or !dat10 Then
280             lblDCTresult = "Positive"
290         ElseIf !DAT1 Or !dat3 Or !dat5 Or !dat7 Or !dat9 Or !dat11 Then
300             lblDCTresult = "Negative"
310         Else
320             lblDCTresult = ""
330         End If
340         lblSampleDate = Format(!SampleDate & "", "dd/mmm/yyyy hh:mm")
350     End If
360 End With



370 Exit Sub

FillPatientDetails_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmTBView", "FillPatientDetails", intEL, strES


End Sub


Private Sub fillGridKleihauer()

    Dim sql As String
    Dim tb As Recordset
    Dim n As Integer
    Dim S As String
    Dim ChartList As String
    Dim PatNameList As String
    Dim DoBList As String
    Dim WhereClause As String

10  On Error GoTo fillGridKleihauer_Error


20  For n = 1 To grd.Rows - 1
30      With grd
40          sql = "SELECT Kleihauer.SampleID, Kleihauer.Chart, Kleihauer.FetalCells, Kleihauer.DateTime, PatientDetails.name" & _
                " FROM PatientDetails INNER JOIN " & _
                " Kleihauer ON PatientDetails.labnumber = Kleihauer.SampleID AND PatientDetails.patnum = Kleihauer.Chart " & _
                " WHERE     (PatientDetails.name = N'" & AddTicks(grd.TextMatrix(n, 2)) & "') AND (Kleihauer.Chart = N'" & grd.TextMatrix(n, 0) & "')"
50      End With



60      Set tb = New Recordset
70      RecOpenServerBB 0, tb, sql
80      With tb
90          If .EOF Then
100         Else
110             While Not .EOF
120                 S = Format(!DateTime, "dd/MMM/yyyy hh:mm") & vbTab & !SampleID & "" & vbTab & !FetalCells & "" & vbTab & "" & getKleihauerfetalComment(!FetalCells & "")
130                 grdKle.AddItem S
140                 .MoveNext
150                 S = ""
160             Wend
170         End If
180     End With
190 Next
200 If grdKle.Rows > 2 And grdKle.TextMatrix(1, 0) = "" Then
210     grdKle.RemoveItem (1)
220 End If

230 Exit Sub

fillGridKleihauer_Error:

    Dim strES As String
    Dim intEL As Integer

240 intEL = Erl
250 strES = Err.Description
260 LogError "frmTBView", "fillGridKleihauer", intEL, strES


End Sub

Private Function getKleihauerfetalComment(ByVal intFetal As Integer) As String
    Dim strMessageType As String
    Dim sngFMH As Single

10  On Error GoTo getKleihauerfetalComment_Error

20  strMessageType = getFoetalCellWording(intFetal)

30  Select Case strMessageType
    Case "M1": getKleihauerfetalComment = "No fetal cells seen."

40  Case "M2": getKleihauerfetalComment = "<2ml - 1500iu Anti D Is sufficient. No further testing is required."

50  Case "M3": sngFMH = Val(intFetal) * 0.4
60      getKleihauerfetalComment = sngFMH & "ml - 1500iu Anti D is sufficient. Repeat Kleihauer testing is required 72hrs post administration of Anti D."

70  Case "M4":
80      sngFMH = Val(intFetal) * 0.4
90      getKleihauerfetalComment = sngFMH & "ml - Send sample for flow cytometry urgently. Repeat Kleihauer and flow cytometry samples required 72hrs post administration of Anti D."
100 Case "M5":
110     sngFMH = Val(intFetal) * 0.4
120     getKleihauerfetalComment = sngFMH & "ml - Send sample for flow cytometry urgently. Further Anti D required discuss with Consultant Haematologist. Repeat Kleihauer and flow cytometry samples required 72hrs post administration of Anti D."
130 End Select

140 Exit Function

getKleihauerfetalComment_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmBTview", "getKleihauerfetalComment", intEL, strES

End Function

Private Function getFoetalCellWording(ByVal intF As Integer) As String

10  On Error GoTo getFoetalCellWording_Error

20  If intF = "0" Then
30      getFoetalCellWording = "M1"
40  ElseIf intF >= "1" And intF <= "4" Then
50      getFoetalCellWording = "M2"
60  ElseIf intF >= "5" And intF <= "19" Then
70      getFoetalCellWording = "M3"
80  ElseIf intF >= "20" And intF <= "30" Then
90      getFoetalCellWording = "M4"
100 ElseIf intF >= "31" Then
110     getFoetalCellWording = "M5"
120 End If

130 Exit Function

getFoetalCellWording_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmBTview", "getFoetalCellWording", intEL, strES

End Function

Private Sub fillProductHistory()

    Dim sql As String
    Dim tb As Recordset
    Dim n As Integer
    Dim S As String
    Dim UnitNumber As String
    Dim intHoldFor As Integer

10  On Error GoTo fillProductHistory_Error

20  intHoldFor = GetOptionSetting("TransfusionHoldFor", "72", "")

30  For n = 1 To grd.Rows - 1
40      With grd
50          sql = "SELECT  Latest.Number, Latest.ISBT128, Latest.PatID, Latest.PatName, Latest.Event, Latest.GroupRH as Unitgroup, Latest.DateTime, " & _
                  "Latest.DateExpiry, ProductList.Wording as Product" & _
                " FROM Latest INNER JOIN ProductList ON Latest.Barcode = ProductList.BarCode" & _
                " WHERE  (Event IN (N'S', N'X', N'I', N'Y')) and (PatID = N'" & grd.TextMatrix(n, 0) & "') AND (PatName = N'" & AddTicks(grd.TextMatrix(n, 2)) & "') " & _
                " Union All " & _
                " SELECT BatchNumber as number, BatchNumber as ISBT128, chart as PatId, Patname,  eventcode as Event, Unitgroup, RecordDateTime as DateTime ," & _
                " dateexpiry, Product from BatchProducts " & _
                " WHERE  (eventcode IN (N'S', N'X', N'I', N'Y')) and (Chart = N'" & grd.TextMatrix(n, 0) & "') AND (PatName = N'" & AddTicks(grd.TextMatrix(n, 2)) & "')" & _
                " ORDER BY DateTime DESC"
60      End With

70      Set tb = New Recordset
80      RecOpenServerBB 0, tb, sql
90      With tb
100         If .EOF Then
110         Else
120             While Not .EOF
130                 UnitNumber = IIf(Len(!ISBT128 & "") > 0, !ISBT128 & "", !Number & "")
140                 S = !Product & vbTab & Bar2Group(!Unitgroup) & vbTab & UnitNumber & vbTab & _
                        getAvailableForDays(UnitNumber, intHoldFor, !DateExpiry, !Event) & vbTab & _
                        EventCode2Text(!Event) & vbTab & Format(!DateTime, "dd/mm/yy hh:mm")
150                 grdProd.AddItem S
160                 .MoveNext
170                 S = ""
180             Wend
190         End If
200     End With
210 Next
220 If grdProd.Rows > 2 And grdProd.TextMatrix(1, 0) = "" Then
230     grdProd.RemoveItem (1)
240 End If

250 Exit Sub

fillProductHistory_Error:

    Dim strES As String
    Dim intEL As Integer

260 intEL = Erl
270 strES = Err.Description
280 LogError "frmTBView", "fillProductHistory", intEL, strES


End Sub

Private Function getAvailableForDays(ByVal strUnit As String, ByVal intHoldFor As Integer, ByVal dateProdExpiry As Date, ByVal strStatus As String) As String
    Dim dateAvailableTo As Date
    Dim tb As Recordset
    Dim sql As String
    Dim strSampleDate As String

    'Add TransfusionAvailableForDays to patients SAmple date if this new date is greater than Expiry THEN use Expiry date otherwise use this new date/time

10  On Error GoTo getAvailableForDays_Error

20  getAvailableForDays = ""

30  sql = "SELECT LATEST.isbt128, Latest.LabNumber, Latest.event, PatientDetails.sampledate  FROM Latest " & _
          "INNER JOIN PatientDetails " & _
          "ON LATEST.labnumber  = PatientDetails.LabNumber " & _
          "WHERE  (LATEST.ISBT128 = N'" & strUnit & "') and " & _
          "(LATEST.Event IN (N'X', N'I'))"

    'need to join labnumber between Latest and Patientdetails and get SampleDate to use below instead of strSampleDate

40  Set tb = New Recordset
50  RecOpenServerBB 0, tb, sql
60  If Not tb.EOF Then    'IF product a blood product like Red Cells, Platelets etc Not a batch product
70      strSampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")

80      dateAvailableTo = DateAdd("h", intHoldFor, strSampleDate)

90      If dateAvailableTo > dateProdExpiry Then
100         dateAvailableTo = dateProdExpiry
110     End If

120     getAvailableForDays = Format(dateAvailableTo, "dd/mmm/yyyy hh:mm")

130 End If

140 Exit Function

getAvailableForDays_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmBTview", "getAvailableForDays", intEL, strES, sql

End Function

Private Function EventCode2Text(Code As String) As String
    Dim S As String
10  On Error GoTo EventCode2Text_Error

20  Code = UCase(Code)
30  Select Case Code
    Case "I": S = "Issued"
40  Case "X": S = "Cross-matched"
50  Case "S": S = "Transfused"
60  Case "Y": S = "Removed Pending Transfusion"
70  End Select
80  EventCode2Text = S

90  Exit Function

EventCode2Text_Error:
    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmBTview", "EventCode2Text", intEL, strES

End Function


Private Sub Form_Activate()

10  If Not blnFormActivated Then
20      FillPatientDetails
30      fillGridKleihauer
40      fillProductHistory
50      blnFormActivated = True
60  End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
10  blnFormActivated = False
End Sub

Private Sub grdKle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10  On Error GoTo grdKle_MouseMove_Error

20  y = grdKle.MouseCol
30  x = grdKle.MouseRow

40  If grdKle.MouseCol = 3 Then
50      If Trim(grdKle.TextMatrix(x, y)) <> "" Then grdKle.ToolTipText = grdKle.TextMatrix(x, y)
60  Else
70      grdKle.ToolTipText = ""
80  End If

90  Exit Sub

grdKle_MouseMove_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmBTview", "grdKle_MouseMove", intEL, strES

End Sub

Private Function getExternalNotes(ByVal strChart As String) As String
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo getExternalNotes_Error

20  getExternalNotes = ""
30  sql = "Select Notes from ExternalNotes where MRN = '" & strChart & "'"
40  Set tb = New Recordset
50  RecOpenServerBB 0, tb, sql
60  If Not tb.EOF Then
70      getExternalNotes = tb!Notes & ""
80  End If

90  Exit Function

getExternalNotes_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmBTview", "getExternalNotes", intEL, strES, sql

End Function



Private Sub tComment_DblClick()

10  With frmExternal
20      If txtChart <> "" Then
30          .Chart = txtChart
40          .AandE = ""
50      End If
60      .Show 1
70  End With

End Sub
