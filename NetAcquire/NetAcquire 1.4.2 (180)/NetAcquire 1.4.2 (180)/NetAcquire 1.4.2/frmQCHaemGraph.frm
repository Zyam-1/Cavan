VERSION 5.00
Begin VB.Form frmQCHaemGraph 
   Caption         =   "NetAcquire"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   645
      Left            =   8880
      Picture         =   "frmQCHaemGraph.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1620
      Width           =   675
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      FillStyle       =   0  'Solid
      Height          =   3585
      Left            =   180
      ScaleHeight     =   3525
      ScaleWidth      =   7725
      TabIndex        =   0
      Top             =   180
      Width           =   7785
   End
   Begin VB.Label lblMin 
      Caption         =   "Label6"
      Height          =   255
      Left            =   9180
      TabIndex        =   12
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblMax 
      Caption         =   "Label5"
      Height          =   315
      Left            =   9120
      TabIndex        =   11
      Top             =   180
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "-4SD"
      Height          =   195
      Left            =   8010
      TabIndex        =   10
      Top             =   3300
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "-2SD"
      Height          =   195
      Left            =   8010
      TabIndex        =   9
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "+4SD"
      Height          =   195
      Left            =   8010
      TabIndex        =   8
      Top             =   540
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "+2SD"
      Height          =   195
      Left            =   8010
      TabIndex        =   7
      Top             =   1230
      Width           =   405
   End
   Begin VB.Label lblMinus2SD 
      Caption         =   "- 2SD"
      Height          =   195
      Left            =   8400
      TabIndex        =   6
      Top             =   2625
      Width           =   855
   End
   Begin VB.Label lblPlus2SD 
      Caption         =   "+2 SD"
      Height          =   195
      Left            =   8460
      TabIndex        =   5
      Top             =   1215
      Width           =   900
   End
   Begin VB.Label lblMinus4SD 
      Caption         =   "0"
      Height          =   195
      Left            =   8400
      TabIndex        =   3
      Top             =   3300
      Width           =   690
   End
   Begin VB.Label lblMeanVal 
      Height          =   255
      Left            =   8010
      TabIndex        =   2
      Top             =   1890
      Width           =   690
   End
   Begin VB.Label lblPlus4SD 
      Height          =   195
      Left            =   8460
      TabIndex        =   1
      Top             =   540
      Width           =   720
   End
End
Attribute VB_Name = "frmQCHaemGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pRunDates() As String
Private pRunTimes() As String
Private pValues() As String

Private pLotNumber As String

Dim ChartX() As Integer
Dim ChartY() As Integer

Private Sub GetScale()

      Dim sql As String
      Dim tb As Recordset

41380 On Error GoTo GetScale_Error

41390 sql = "SELECT TOP 1 " & _
            "COALESCE(Mean,1) AS Mean, " & _
            "COALESCE(SD1,0.1) AS SD1 " & _
            "FROM HaemControlDefinitions " & _
            "WHERE LotNumber = '" & pLotNumber & "' " & _
            "AND Analyte = '" & pValues(0) & "' " & _
            "ORDER BY DateEntered DESC"
41400 Set tb = New Recordset
41410 RecOpenServer 0, tb, sql
41420 If Not tb.EOF Then
41430   lblMeanVal = tb!mean
41440   lblPlus2SD = tb!mean + (2 * tb!SD1)
41450   lblMinus2SD = tb!mean - (2 * tb!SD1)
41460   lblPlus4SD = tb!mean + (4 * tb!SD1)
41470   lblMinus4SD = tb!mean - (4 * tb!SD1)
41480   lblMax = tb!mean + (5 * tb!SD1)
41490   lblMin = tb!mean - (5 * tb!SD1)
41500 End If

41510 Exit Sub

GetScale_Error:

      Dim strES As String
      Dim intEL As Integer

41520 intEL = Erl
41530 strES = Err.Description
41540 LogError "frmQCHaemGraph", "GetScale", intEL, strES, sql


End Sub

Public Property Let RunDates(ByVal strNewValue As String)

41550 pRunDates = Split(strNewValue, vbTab)

End Property
Private Sub DrawChart()

      Dim n As Integer
      Dim DaysInterval As Long
      Dim X As Integer
      Dim Y As Integer
      Dim PixelsPerDay As Single
      Dim PixelsPerPointY As Single
      Dim FirstDayFilled As Boolean
      Dim LatestDate As String
      Dim EarliestDate As String
      Dim NumberOfDays As Integer
41560 On Error GoTo DrawChart_Error

41570 ReDim ChartX(0 To UBound(pValues)) As Integer
41580 ReDim ChartY(0 To UBound(pValues)) As Integer

41590 LatestDate = pRunDates(1)
41600 EarliestDate = pRunDates(UBound(pRunDates))

41610 pb.Cls
41620 pb.Picture = LoadPicture("")
        
41630 If Val(lblMax) = 0 Then Exit Sub

41640 NumberOfDays = Abs(DateDiff("d", Format(LatestDate, "dd/MMM/yyyy"), Format(EarliestDate, "dd/MMM/yyyy")))
        
41650 If NumberOfDays <> 0 Then
          
41660   FirstDayFilled = False
          
41670   PixelsPerDay = (pb.width / NumberOfDays) * 0.95
          
41680   PixelsPerPointY = pb.height / (Val(lblMax) - Val(lblMin))
          
41690   For n = 1 To UBound(pValues)
41700     If Val(pValues(n)) <> 0 Then
41710       pb.ForeColor = vbGreen
41720       pb.FillColor = vbGreen
41730       DaysInterval = Abs(DateDiff("d", pRunDates(1), Format(pRunDates(n), "dd/mmm/yyyy")))
41740       X = (DaysInterval * PixelsPerDay)
41750       ChartX(n) = X
41760       If pValues(n) > Val(lblPlus4SD) Then
41770         Y = pb.height - ((Val(lblMax) - ((Val(lblMax) - Val(lblPlus4SD)) / 2) - Val(lblMin)) * PixelsPerPointY)
41780         pb.ForeColor = vbBlack
41790         pb.FillColor = vbBlack
41800       ElseIf pValues(n) > Val(lblPlus2SD) Then
41810         Y = pb.height - (Val(pValues(n) - Val(lblMin)) * PixelsPerPointY)
41820         pb.ForeColor = vbRed
41830         pb.FillColor = vbRed
41840       ElseIf pValues(n) < Val(lblMinus4SD) Then
41850         Y = pb.height - (Val(pValues(n) - Val(lblMin)) + (Val(lblMinus4SD) - Val(lblMin)) * PixelsPerPointY)
41860         pb.ForeColor = vbBlack
41870         pb.FillColor = vbBlack
41880       ElseIf pValues(n) < Val(lblMinus2SD) Then
41890         Y = pb.height - (Val(pValues(n) - Val(lblMin)) * PixelsPerPointY)
41900         pb.ForeColor = vbBlue
41910         pb.FillColor = vbBlue
41920       Else
41930         Y = pb.height - (Val(pValues(n) - Val(lblMin)) * PixelsPerPointY)
41940       End If

41950       Debug.Print X, Y
41960       ChartY(n) = Y
41970       pb.Circle (X, Y), 80
41980       pb.PSet (X, Y)
41990     End If
42000   Next

42010   pb.Line (0, pb.height / 2)-(pb.width, pb.height / 2), vbBlack, BF
42020   Y = pb.height - ((Val(lblPlus4SD) - Val(lblMin)) * PixelsPerPointY)
42030   pb.Line (0, Y)-(pb.width, Y), vbBlack, BF
42040   Y = pb.height - ((Val(lblMinus4SD) - Val(lblMin)) * PixelsPerPointY)
42050   pb.Line (0, Y)-(pb.width, Y), vbBlack, BF
42060   Y = pb.height - ((Val(lblPlus2SD) - Val(lblMin)) * PixelsPerPointY)
42070   pb.Line (0, Y)-(pb.width, Y), vbBlack, BF
42080   Y = pb.height - ((Val(lblMinus2SD) - Val(lblMin)) * PixelsPerPointY)
42090   pb.Line (0, Y)-(pb.width, Y), vbBlack, BF
          
42100 End If

42110 Exit Sub

DrawChart_Error:

      Dim strES As String
      Dim intEL As Integer

42120 intEL = Erl
42130 strES = Err.Description
42140 LogError "frmQCHaemGraph", "DrawChart", intEL, strES

End Sub

Public Property Let RunTimes(ByVal strNewValue As String)

42150 pRunTimes = Split(strNewValue, vbTab)

End Property

Public Property Let Values(ByVal strNewValue As String)

42160 pValues = Split(strNewValue, vbTab)

End Property


Private Sub cmdCancel_Click()

42170 Unload Me

End Sub

Private Sub Form_Load()

42180 Me.Caption = "NetAcquire - " & pValues(0) & " Controls"

42190 GetScale
42200 DrawChart

End Sub

Private Sub Pb_Click()

42210 Unload Me

End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim i As Integer
      Dim CurrentDistance As Long
      Dim BestDistance As Long
      Dim BestIndex As Integer

42220 On Error GoTo pbmm

42230 BestIndex = -1
42240 BestDistance = 99999
42250 For i = 1 To UBound(pValues)
42260   CurrentDistance = ((X - ChartX(i)) ^ 2 + (Y - ChartY(i)) ^ 2) ^ (1 / 2)
42270   If i = 0 Or CurrentDistance < BestDistance Then
42280     BestDistance = CurrentDistance
42290     BestIndex = i
42300   End If
42310 Next

42320 If BestIndex <> -1 Then
42330   pb.ToolTipText = Format$(pRunDates(BestIndex), "dd/mmm/yyyy") & " " & Format$(pRunTimes(BestIndex), "HH:nn") & " " & pValues(BestIndex)
42340 End If

42350 Exit Sub

pbmm:
42360 Exit Sub

End Sub



Public Property Let LotNumber(ByVal strNewValue As String)

42370 pLotNumber = strNewValue

End Property
