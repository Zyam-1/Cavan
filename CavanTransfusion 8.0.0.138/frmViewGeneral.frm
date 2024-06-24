VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewGeneral 
   Caption         =   "NetAcquire"
   ClientHeight    =   6585
   ClientLeft      =   90
   ClientTop       =   810
   ClientWidth     =   13275
   DrawWidth       =   10
   Icon            =   "frmViewGeneral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   13275
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   0
      Left            =   660
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   8
      Top             =   270
      Width           =   3555
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   195
         Index           =   0
         Left            =   825
         TabIndex        =   9
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   1
      Left            =   5010
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   7
      Top             =   270
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   2
      Left            =   9420
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   6
      Top             =   270
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   3
      Left            =   660
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   5
      Top             =   2310
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   4
      Left            =   5010
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   4
      Top             =   2310
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   5
      Left            =   9390
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   2310
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   6
      Left            =   660
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   4350
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   7
      Left            =   9390
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   1
      Top             =   4380
      Width           =   3555
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      Height          =   1605
      Index           =   8
      Left            =   5010
      ScaleHeight     =   1545
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   4380
      Width           =   3555
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   660
      TabIndex        =   46
      Top             =   6420
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   12420
      TabIndex        =   45
      Top             =   6000
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   12420
      TabIndex        =   44
      Top             =   3930
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   43
      Top             =   6000
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   42
      Top             =   3930
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   3690
      TabIndex        =   41
      Top             =   5970
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   3690
      TabIndex        =   40
      Top             =   3930
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   12450
      TabIndex        =   39
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   38
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label lValue 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   3690
      TabIndex        =   37
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   36
      Top             =   1860
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   35
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   5010
      TabIndex        =   34
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   7110
      TabIndex        =   33
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   9420
      TabIndex        =   32
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   11520
      TabIndex        =   31
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   30
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   29
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   660
      TabIndex        =   28
      Top             =   5970
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   27
      Top             =   5970
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   5010
      TabIndex        =   26
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   7110
      TabIndex        =   25
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   9390
      TabIndex        =   24
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   11490
      TabIndex        =   23
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   9180
      TabIndex        =   22
      Top             =   2310
      Width           =   195
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9000
      TabIndex        =   21
      Top             =   270
      Width           =   405
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   270
      Width           =   300
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FIB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   4620
      TabIndex        =   19
      Top             =   2310
      Width           =   360
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "APTT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   2310
      Width           =   555
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hgb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   210
      TabIndex        =   17
      Top             =   270
      Width           =   420
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Urea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4350
      Width           =   480
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WCC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   4470
      TabIndex        =   15
      Top             =   4410
      Width           =   495
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Neut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   8910
      TabIndex        =   14
      Top             =   4380
      Width           =   480
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   11490
      TabIndex        =   13
      Top             =   6000
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   9390
      TabIndex        =   12
      Top             =   6000
      Width           =   930
   End
   Begin VB.Label lTo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   7110
      TabIndex        =   11
      Top             =   6000
      Width           =   930
   End
   Begin VB.Label lFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   5010
      TabIndex        =   10
      Top             =   6000
      Width           =   930
   End
End
Attribute VB_Name = "frmViewGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type udtResults
  SampleID As String
  RunDate As String
  Value As String
End Type

Dim LookupSampleIDs() As New Collection
Dim LookupRunDates() As New Collection
Dim LookupValues() As New Collection
Dim LookupX() As New Collection
Dim LookupY() As New Collection
Private Sub Form_Click()

10    Unload Me

End Sub

Private Sub Form_Load()

10    Me.Caption = "NetAcquire - Lab Results for Chart # " & frmxmatch.txtChart

20    DrawGraphs

End Sub

Private Sub FillParameterList(ByRef Results() As udtResults, _
                              ByVal HBC As String, _
                              ByVal Parameter As String)


      Dim sn As Recordset
      Dim tb As Recordset
      Dim snr As Recordset
      Dim sql As String
      Dim X As Integer
      Dim ParameterCode As String
      Dim Chart As String

10    On Error GoTo FillParameterList_Error

20    Chart = frmxmatch.txtChart

30    ReDim Results(0 To 0) As udtResults

40    sql = "select sampleid, runDate from demographics where " & _
            "chart = '" & Chart & "' and "
50    Select Case HBC
        Case "H": sql = sql & "ForHaem = '1'"
60      Case "B": sql = sql & "ForBio = '1'"
70      Case "C": sql = sql & "ForCoag = '1'"
80    End Select
90    sql = sql & " order by RunDate desc"
100   Set sn = New Recordset
110   RecOpenClient 0, sn, sql

120   If Not sn.EOF Then
130     ReDim Results(0 To sn.RecordCount) As udtResults

140     For X = 1 To UBound(Results)
150       With Results(X)
160         .SampleID = sn!SampleID & ""
170         If Not IsNull(sn!RunDate) Then
180           .RunDate = Format(sn!RunDate, "dd/mm/yy")
190         Else
200           .RunDate = ""
210         End If
220         Select Case HBC
              Case "H"
230             sql = "select " & Parameter & " as result from HaemResults where " & _
                      "SampleID = '" & .SampleID & "'"
240           Case "B"
250             sql = "Select Code from BioTestDefinitions where " & _
                      "ShortName = '" & Parameter & "'"
260             Set tb = New Recordset
270             RecOpenServer 0, tb, sql
280             If Not tb.EOF Then
290               ParameterCode = tb!code & ""
300             End If
    
310             sql = "select * from BioResults where " & _
                      "SampleID = '" & .SampleID & "' " & _
                      "and Code = '" & ParameterCode & "'"
320           Case "C"
330             sql = "Select Code from CoagTestDefinitions where " & _
                      "TestName = '" & Parameter & "'"
340             Set tb = New Recordset
350             RecOpenServer 0, tb, sql
360             If Not tb.EOF Then
370               ParameterCode = tb!code & ""
380             End If
390             sql = "select * from CoagResults where " & _
                      "SampleID = '" & .SampleID & "' " & _
                      "and Code = '" & ParameterCode & "'"
400         End Select

410         Set snr = New Recordset
420         RecOpenServer 0, snr, sql

430         If Not snr.EOF Then
440           .Value = snr!Result & ""
450         End If

460         sn.MoveNext
470       End With
480     Next
490   End If

500   Exit Sub

FillParameterList_Error:

      Dim strES As String
      Dim intEL As Integer

510   intEL = Erl
520   strES = Err.Description
530   LogError "frmViewGeneral", "FillParameterList", intEL, strES, sql

End Sub

Private Sub DrawGraphs()

10    ReDim Results(0 To 0) As udtResults

20    ReDim LookupSampleIDs(0 To 8)
30    ReDim LookupRunDates(0 To 8)
40    ReDim LookupValues(0 To 8)
50    ReDim LookupX(0 To 8)
60    ReDim LookupY(0 To 8)

70    FillParameterList Results(), "H", "Hgb"
80    DrawChart Results(), 0, 20

90    FillParameterList Results(), "H", "Plt"
100   DrawChart Results(), 1, 600

110   FillParameterList Results(), "C", "INR"
120   DrawChart Results(), 2, 8

130   FillParameterList Results(), "C", "APTT"
140   DrawChart Results(), 3, 200

150   FillParameterList Results(), "C", "FIB"
160   DrawChart Results(), 4, 8

170   FillParameterList Results(), "B", "K"
180   DrawChart Results(), 5, 9

190   FillParameterList Results(), "B", "URE3"
200   DrawChart Results(), 6, 80

210   FillParameterList Results(), "H", "NeutA"
220   DrawChart Results(), 7, 30

230   FillParameterList Results(), "H", "WBC"
240   DrawChart Results(), 8, 30

End Sub
Private Sub DrawChart(ByRef Results() As udtResults, _
                      ByVal Index As Integer, _
                      ByVal max As Single)

      Dim n As Integer
      Dim Counter As Integer
      Dim LatestDate As String
      Dim EarliestDate As String
      Dim NumberOfDays As Long
      Dim X As Long
      Dim Y As Integer
      Dim PixelsPerDay As Single
      Dim PixelsPerPointY As Single
      Dim FirstDayFilled As Boolean
      Dim gData(1 To 365, 1 To 2) As Variant '(n,1)=rundate, (n,2)=Value
      Dim cVal As Single

10    On Error GoTo DrawChart_Error

20    For n = 0 To UBound(Results)
30      LookupValues(Index).Add CStr(Results(n).Value), CStr(n)
40      LookupRunDates(Index).Add CStr(Results(n).RunDate), CStr(n)
50      LookupSampleIDs(Index).Add CStr(Results(n).SampleID), CStr(n)
60    Next

70    p(Index).Cls
80    p(Index).Picture = LoadPicture("")

90    For n = 1 To 365
100     gData(n, 1) = 0
110     gData(n, 2) = 0
120   Next

130   lFrom(Index) = ""
140   lTo(Index) = ""
150   lValue(Index) = ""
160   p(Index).Picture = LoadPicture("")
170   DoEvents

180   FirstDayFilled = False
190   Counter = 0
200   For X = 1 To UBound(Results)
210     With Results(X)
220       If .Value <> "" Then
230         If Not FirstDayFilled Then
240           FirstDayFilled = True
250           gData(365, 1) = Format$(.RunDate, "dd/mmm/yyyy")
260           LatestDate = Format$(.RunDate, "dd/mmm/yyyy")
270           gData(365, 2) = Val(.Value)
280           lValue(Index) = Val(.Value)
290         Else
300           NumberOfDays = Abs(DateDiff("D", LatestDate, Format$(.RunDate, "dd/mmm/yyyy")))
310           If NumberOfDays < 365 Then
320             gData(365 - NumberOfDays, 1) = .RunDate
330             cVal = Val(.Value)
340             gData(365 - NumberOfDays, 2) = cVal
350             EarliestDate = Format$(.RunDate, "dd/mmm/yyyy")
360           Else
370             Exit For
380           End If
390         End If
400         Counter = Counter + 1
410         If Counter = 15 Then
420           Exit For
430         End If
440       End If
450     End With
460   Next

470   If EarliestDate = "" Or LatestDate = "" Then Exit Sub

480   NumberOfDays = Abs(DateDiff("d", EarliestDate, LatestDate))

490   With p(Index)
500     PixelsPerDay = (.Width - 1060) / NumberOfDays
510     PixelsPerPointY = .Height / max
  
520     X = 580 + (NumberOfDays * PixelsPerDay)
530     Y = .Height - (gData(365, 2) * PixelsPerPointY)
  
540     .ForeColor = vbRed
550     p(Index).Circle (X, Y), 30
560     p(Index).Line (X - 15, Y - 15)-(X + 15, Y + 15), vbRed, BF
570     p(Index).PSet (X, Y)
  
580     For n = 364 To 1 Step -1
590       If gData(n, 1) <> 0 Then
600         NumberOfDays = Abs(DateDiff("d", EarliestDate, Format(gData(n, 1), "dd/mmm/yyyy")))
610         X = 580 + (NumberOfDays * PixelsPerDay)
620         Y = .Height - (gData(n, 2) * PixelsPerPointY)
630         p(Index).Line -(X, Y)
640         p(Index).Line (X - 15, Y - 15)-(X + 15, Y + 15), vbRed, BF
650         p(Index).Circle (X, Y), 30
660         p(Index).PSet (X, Y)
670       End If
680     Next
690   End With

700   lFrom(Index) = Format(EarliestDate, "dd/mm/yyyy")
710   lTo(Index) = Format(LatestDate, "dd/mm/yyyy")

720   If lValue(Index) = "0" Then lValue(Index) = Format(Val(Results(UBound(Results)).Value))
730   If lValue(Index) = "0" Then lValue(Index) = ""

740   Exit Sub

DrawChart_Error:

      Dim strES As String
      Dim intEL As Integer

750   intEL = Erl
760   strES = Err.Description
770   LogError "frmViewGeneral", "DrawChart", intEL, strES

End Sub

Private Sub p_Click(Index As Integer)

10    Unload Me

End Sub

