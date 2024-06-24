VERSION 5.00
Begin VB.Form frmINR 
   Caption         =   "NetAcquire - INR"
   ClientHeight    =   4845
   ClientLeft      =   390
   ClientTop       =   1620
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   12510
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton bCondition 
      Caption         =   "C&ondition"
      Height          =   285
      Left            =   11430
      TabIndex        =   20
      Top             =   420
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "INR Target Range"
      Height          =   1455
      Left            =   10500
      TabIndex        =   16
      Top             =   1440
      Width           =   1635
      Begin VB.CommandButton bTarget 
         Caption         =   "Enter INR Target Range"
         Height          =   585
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lLowerTarget 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lUpperTarget 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   900
         TabIndex        =   18
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   3585
      Left            =   300
      ScaleHeight     =   3525
      ScaleWidth      =   8805
      TabIndex        =   8
      Top             =   810
      Width           =   8865
   End
   Begin VB.TextBox tChart 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   1485
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10740
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lCondition 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9390
      TabIndex        =   15
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label lCurrentDose 
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   4410
      Width           =   750
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Current Warfarin Dose"
      Height          =   195
      Left            =   3090
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lLatest 
      BackColor       =   &H80000018&
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
      Left            =   8190
      TabIndex        =   12
      Top             =   5940
      Width           =   1095
   End
   Begin VB.Label lEarliest 
      BackColor       =   &H80000018&
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
      Left            =   750
      TabIndex        =   11
      Top             =   5820
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         W      a    D r    o  f    s  a   a  r    g  i    e n      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3600
      Left            =   9300
      TabIndex        =   10
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      I  N R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   3585
      Left            =   60
      TabIndex        =   9
      Top             =   810
      Width           =   285
   End
   Begin VB.Label lAddress 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   5940
      TabIndex        =   7
      Top             =   420
      Width           =   3345
   End
   Begin VB.Label lAddress 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   5940
      TabIndex        =   6
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label lDoB 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4290
      TabIndex        =   5
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label lSex 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2190
      TabIndex        =   4
      Top             =   450
      Width           =   1245
   End
   Begin VB.Label lName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2190
      TabIndex        =   3
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chart Number"
      Height          =   195
      Left            =   510
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmINR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gData(1 To 365, 1 To 3) As Variant '(n,1)=rundate, (n,2)=INR, (n,3)=Warfarin

Private LatestSampleID As String

Private LatestINR As String

Private mWard As String

Private ChartChanged As Boolean 'set if keypress on chart number


Private Sub DrawChart()

          Dim tb As Recordset
          Dim sn As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Counter As Integer
          Dim LatestDate As String
          Dim EarliestDate As String
          Dim NumberOfDays As Long
          Dim X As Integer
          Dim Y As Integer
          Dim PixelsPerDay As Single
          Dim PixelsPerPointY As Integer
          Dim FirstDayFilled As Boolean
          Dim CR As CoagResult
          Dim CRs As CoagResults
          Dim CodeForINR As String

7790      On Error GoTo DrawChart_Error

7800      CodeForINR = "044"

7810      pb.Cls
7820      pb.Picture = LoadPicture("")

7830      For n = 1 To 365
7840          gData(n, 1) = 0
7850          gData(n, 2) = 0
7860          gData(n, 3) = 0
7870      Next

7880      pb.Font.Bold = True
7890      For Y = 0 To 6
7900          pb.CurrentY = (pb.height - ((pb.height - 1.5 * pb.TextHeight("W")) / 6) * Y) - 1.5 * TextHeight("W")
7910          pb.CurrentX = 0
7920          pb.ForeColor = vbBlue
7930          pb.Print Format$(Y)
7940          pb.CurrentY = (pb.height - ((pb.height - 1.5 * pb.TextHeight("W")) / 6) * Y) - 1.5 * TextHeight("W")
7950          pb.CurrentX = pb.width - pb.TextWidth("WW")
7960          pb.ForeColor = vbGreen
7970          pb.Print Format$(Y * 2)
7980      Next

7990      sql = "select sampleid, rundate from demographics where " & _
              "chart = '" & tChart & "' " & _
              "order by rundate desc"
        
8000      Set sn = New Recordset
8010      RecOpenClient 0, sn, sql
8020      If sn.EOF Then Exit Sub

8030      FirstDayFilled = False
8040      Counter = 0
8050      Do While Not sn.EOF
8060          sql = "Select * from HaemResults where " & _
                  "SampleID = '" & sn!SampleID & "'"
8070          Set tb = New Recordset
8080          RecOpenClient 0, tb, sql
8090          Set CRs = New CoagResults
8100          Set CRs = CRs.Load(sn!SampleID & "", gDONTCARE, gDONTCARE, "Results")
8110          If Not CRs Is Nothing Then
8120              For Each CR In CRs
8130                  If CR.Code = CodeForINR Then
8140                      If Not FirstDayFilled Then
8150                          FirstDayFilled = True
8160                          gData(365, 1) = Format$(sn!Rundate, "dd/mmm/yyyy")
8170                          gData(365, 2) = Val(CR.Result)
8180                          If Not tb.EOF Then
8190                              gData(365, 3) = Val(tb!Warfarin & "")
8200                              lCurrentDose = tb!Warfarin & ""
8210                          Else
8220                              lCurrentDose = ""
8230                          End If
8240                          LatestSampleID = sn!SampleID & ""
8250                          LatestDate = Format$(sn!Rundate, "dd/mmm/yyyy")
8260                          LatestINR = CR.Result
8270                          lLatest = Format$(LatestDate, "dd/mm/yyyy")
8280                      Else
8290                          NumberOfDays = Abs(DateDiff("D", LatestDate, Format$(sn!Rundate, "dd/mmm/yyyy")))
8300                          If NumberOfDays < 365 Then
8310                              gData(365 - NumberOfDays, 1) = Format$(sn!Rundate, "dd/mmm/yyyy")
8320                              gData(365 - NumberOfDays, 2) = Val(CR.Result)
8330                              If Not tb.EOF Then
8340                                  gData(365 - NumberOfDays, 3) = Val(tb!Warfarin & "")
8350                              End If
8360                              EarliestDate = Format$(sn!Rundate, "dd/mmm/yyyy")
8370                              lEarliest = Format$(sn!Rundate, "dd/mm/yyyy")
8380                          Else
8390                              Exit Do
8400                          End If
8410                      End If
8420                      Counter = Counter + 1
8430                      If Counter = 15 Then
8440                          Exit Do
8450                      End If
8460                  End If
8470              Next
8480          End If
8490          sn.MoveNext
8500      Loop

8510      If EarliestDate = "" Or LatestDate = "" Then Exit Sub

8520      NumberOfDays = Abs(DateDiff("d", EarliestDate, LatestDate))
8530      PixelsPerDay = (pb.width - 1060) / NumberOfDays
8540      PixelsPerPointY = pb.height / 6

8550      X = pb.width - 580
8560      Y = pb.height - (gData(365, 2) * PixelsPerPointY)

8570      pb.ForeColor = vbBlue
8580      pb.Circle (X, Y), 30
8590      pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
8600      pb.PSet (X, Y)

8610      For n = 364 To 1 Step -1
8620          If gData(n, 1) <> 0 Then
8630              NumberOfDays = Abs(DateDiff("d", EarliestDate, gData(n, 1)))
8640              X = 580 + (NumberOfDays * PixelsPerDay)
8650              Y = pb.height - (Val(gData(n, 2)) * PixelsPerPointY)
8660              pb.Line -(X, Y)
8670              pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
8680              pb.Circle (X, Y), 30
8690              pb.PSet (X, Y)
8700          End If
8710      Next

          'Draw Warfarin
8720      NumberOfDays = Abs(DateDiff("d", EarliestDate, LatestDate))
8730      PixelsPerPointY = pb.height / 12

8740      X = pb.width - 580
8750      Y = pb.height - (gData(365, 3) * PixelsPerPointY)

8760      pb.ForeColor = vbGreen
8770      pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbGreen, BF
8780      pb.PSet (X, Y)

8790      For n = 364 To 1 Step -1
8800          If gData(n, 1) <> 0 Then
8810              NumberOfDays = Abs(DateDiff("d", EarliestDate, gData(n, 1)))
8820              X = 480 + (NumberOfDays * PixelsPerDay)
8830              Y = pb.height - (gData(n, 3) * PixelsPerPointY)
8840              pb.Line -(X, Y)
8850              pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbGreen, BF
8860              pb.PSet (X, Y)
8870          End If
8880      Next

8890      sql = "select * from INRHistory where " & _
              "chart = '" & tChart & "'"
8900      Set tb = New Recordset
8910      RecOpenClient 0, tb, sql
8920      If Not tb.EOF Then
8930          PixelsPerPointY = pb.height / 6
8940          pb.ForeColor = vbRed
8950          Y = pb.height - (Val(tb!UpperTarget & "") * PixelsPerPointY)
8960          pb.Line (480, Y)-(pb.width - 580, Y), vbRed
8970          Y = pb.height - (Val(tb!LowerTarget & "") * PixelsPerPointY)
8980          pb.Line (480, Y)-(pb.width - 580, Y), vbRed
8990          lLowerTarget = tb!LowerTarget & ""
9000          lUpperTarget = tb!UpperTarget & ""
9010          lCondition = tb!Condition & ""
9020      End If

9030      pb.ForeColor = vbBlack
9040      pb.Font.Bold = True
9050      pb.CurrentX = pb.TextWidth("I") + 380
9060      pb.CurrentY = pb.height - 1.5 * pb.TextHeight("W")
9070      pb.Print lEarliest
9080      pb.CurrentX = pb.width - pb.TextWidth("ww/ww/www") - 480
9090      pb.CurrentY = pb.height - 1.5 * pb.TextHeight("W")
9100      pb.Print lLatest

9110      Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

9120      intEL = Erl
9130      strES = Err.Description
9140      LogError "fINR", "DrawChart", intEL, strES, sql

End Sub

Public Sub LoadDetails()

          Dim tbPatIF As Recordset
          Dim tbDemog As Recordset
          Dim sql As String
          Dim X As Long

9150      On Error GoTo LoadDetails_Error

9160      pb.Cls

9170      LatestSampleID = ""

9180      tChart = Trim$(tChart)
9190      If tChart = "" Then Exit Sub

9200      sql = "SELECT * FROM PatientIFs WHERE " & _
              "Chart = '" & tChart & "' " & _
              "AND Entity = '" & Entity & "' " & _
              "AND ISNULL(DateTimeAmended ,0) <> 0"
9210      Set tbPatIF = New Recordset
9220      RecOpenServer 0, tbPatIF, sql

9230      sql = "SELECT * FROM Demographics WHERE " & _
              "Chart = '" & tChart & "' " & _
              "AND ISNULL(DateTimeDemographics ,0) <> 0 " & _
              "ORDER BY DateTimeDemographics desc"
9240      Set tbDemog = New Recordset
9250      RecOpenServer 0, tbDemog, sql

9260      If tbPatIF.EOF And tbDemog.EOF Then
9270          lname = ""
9280          lAddress(0) = ""
9290          lAddress(1) = ""
9300          lsex = ""
9310          ldob = ""
9320          lCondition = ""
9330          lLowerTarget = ""
9340          lUpperTarget = ""
9350          lEarliest = ""
9360          lLatest = ""
9370          lCurrentDose = ""
9380      ElseIf tbDemog.EOF Then
9390          With tbPatIF
9400              lname = !PatName & ""
9410              lAddress(0) = !Address0 & ""
9420              lAddress(1) = !Address1 & ""
9430              lsex = !Sex & ""
9440              ldob = !DoB & ""
9450          End With
9460      ElseIf tbPatIF.EOF Then
9470          With tbDemog
9480              lname = !PatName & ""
9490              lAddress(0) = !Addr0 & ""
9500              lAddress(1) = !Addr1 & ""
9510              lsex = !Sex & ""
9520              ldob = !DoB & ""
9530          End With
9540      Else
9550          X = DateDiff("h", tbDemog!DateTimeDemographics, tbPatIF!DateTimeAmended)
9560          If X < 0 Or IsNull(X) Then
9570              With tbDemog
9580                  lname = !PatName & ""
9590                  lAddress(0) = !Addr0 & ""
9600                  lAddress(1) = !Addr1 & ""
9610                  lsex = !Sex & ""
9620                  ldob = !DoB & ""
9630              End With
9640          Else
9650              With tbPatIF
9660                  lname = !PatName & ""
9670                  lAddress(0) = !Address0 & ""
9680                  lAddress(1) = !Address1 & ""
9690                  lsex = !Sex & ""
9700                  ldob = !DoB & ""
9710              End With
9720          End If
9730      End If

9740      DrawChart

9750      Screen.MousePointer = 0

9760      Exit Sub

LoadDetails_Error:

          Dim strES As String
          Dim intEL As Integer

9770      intEL = Erl
9780      strES = Err.Description
9790      LogError "fINR", "LoadDetails", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

9800      Unload Me

End Sub


Private Sub bCondition_Click()

          Dim sql As String
          Dim tb As Recordset

9810      On Error GoTo bCondition_Click_Error

9820      If Trim$(tChart) = "" Then
9830          iMsg "Enter Chart Number", vbCritical
9840          Exit Sub
9850      End If

9860      lCondition = iBOX("Enter Condition", , lCondition)

9870      sql = "Select * from INRHistory where Chart = '" & tChart & "'"
9880      Set tb = New Recordset
9890      RecOpenClient 0, tb, sql
9900      If tb.EOF Then
9910          tb.AddNew
9920          tb!Chart = tChart
9930          tb!LowerTarget = lLowerTarget
9940          tb!UpperTarget = lUpperTarget
9950          tb!Condition = lCondition
9960          tb.Update
9970      Else
9980          sql = "Update INRHistory Set Condition = '" & lCondition & "' where " & _
                  "Chart = '" & tChart & "'"
9990          Cnxn(0).Execute sql
10000     End If

10010     Exit Sub

bCondition_Click_Error:

          Dim strES As String
          Dim intEL As Integer

10020     intEL = Erl
10030     strES = Err.Description
10040     LogError "fINR", "bCondition_Click", intEL, strES, sql

End Sub

Private Sub bPrint_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Ward As String
          Dim Clin As String
          Dim GP As String

10050     On Error GoTo bPrint_Click_Error

10060     GetWardClinGP LatestSampleID, Ward, Clin, GP

10070     sql = "Select * from PrintPending where " & _
              "Department = 'A' " & _
              "and SampleID = '" & LatestSampleID & "'"
10080     Set tb = New Recordset
10090     RecOpenClient 0, tb, sql
10100     If tb.EOF Then
10110         tb.AddNew
10120     End If
10130     tb!SampleID = LatestSampleID
10140     tb!Ward = Ward
10150     tb!Clinician = Clin
10160     tb!GP = GP
10170     tb!Department = "A"
10180     tb!Initiator = UserName
10190     tb!UsePrinter = frmEditAll.PrintToPrinter
10200     tb.Update

10210     sql = "Update CoagResults " & _
              "Set Valid = 1, Printed = 1 where " & _
              "SampleID = '" & LatestSampleID & "'"
10220     Cnxn(0).Execute sql

10230     Unload Me

10240     Exit Sub

bPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

10250     intEL = Erl
10260     strES = Err.Description
10270     LogError "fINR", "bPrint_Click", intEL, strES, sql


End Sub

Private Sub bTarget_Click()

          Dim sql As String
          Dim tb As Recordset

10280     On Error GoTo bTarget_Click_Error

10290     If Trim$(tChart) = "" Then
10300         iMsg "Enter Chart Number", vbCritical
10310         Exit Sub
10320     End If

10330     lLowerTarget = Left$(Format$(Val(iBOX("Enter Lower Target for INR", , lLowerTarget))), 5)
10340     lUpperTarget = Left$(Format$(Val(iBOX("Enter Upper Target for INR", , lUpperTarget))), 5)

10350     sql = "Select * from INRHistory where " & _
              "Chart = '" & tChart & "'"
10360     Set tb = New Recordset
10370     RecOpenClient 0, tb, sql
10380     If tb.EOF Then
10390         tb.AddNew
10400         tb!Chart = tChart
10410         tb!LowerTarget = lLowerTarget
10420         tb!UpperTarget = lUpperTarget
10430         tb!Condition = lCondition
10440         tb.Update
10450     Else
10460         sql = "Update INRHistory Set LowerTarget = '" & lLowerTarget & "', UpperTarget = '" & lUpperTarget & "' where Chart = '" & tChart & "'"
10470         Cnxn(0).Execute sql
10480     End If

10490     Exit Sub

bTarget_Click_Error:

          Dim strES As String
          Dim intEL As Integer

10500     intEL = Erl
10510     strES = Err.Description
10520     LogError "fINR", "bTarget_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10530     ChartChanged = False

End Sub

Private Sub tChart_KeyPress(KeyAscii As Integer)

10540     ChartChanged = True

End Sub

Private Sub tchart_LostFocus()

10550     If ChartChanged Then
10560         LoadDetails
10570     End If

End Sub


Public Property Let Ward(ByVal Ward As String)

10580     mWard = Ward

End Property
