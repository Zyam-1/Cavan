VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBioQC 
   Caption         =   "NetAcquire"
   ClientHeight    =   6585
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   11070
   Begin VB.ListBox lstCode 
      Height          =   2010
      Left            =   1980
      TabIndex        =   30
      Top             =   1770
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plot"
      Height          =   555
      Left            =   2130
      TabIndex        =   22
      Top             =   5730
      Width           =   2955
      Begin VB.OptionButton optShow 
         Caption         =   "Both"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Width           =   645
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Daily Mean"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   6030
      TabIndex        =   16
      Top             =   240
      Width           =   3525
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   326172673
         CurrentDate     =   37649
      End
      Begin ComCtl2.UpDown udPrevious 
         Height          =   315
         Left            =   2760
         TabIndex        =   19
         Top             =   180
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         Value           =   30
         BuddyControl    =   "lblPrevious"
         BuddyDispid     =   196614
         OrigLeft        =   4890
         OrigTop         =   450
         OrigRight       =   5130
         OrigBottom      =   825
         Increment       =   30
         Max             =   600
         Min             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "days"
         Height          =   195
         Left            =   3060
         TabIndex        =   21
         Top             =   210
         Width           =   330
      End
      Begin VB.Label lblPrevious 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         Height          =   315
         Left            =   2340
         TabIndex        =   20
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "and Previous"
         Height          =   195
         Left            =   1380
         TabIndex        =   18
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.PictureBox picDates 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2100
      ScaleHeight     =   615
      ScaleWidth      =   7365
      TabIndex        =   15
      Top             =   4950
      Width           =   7425
   End
   Begin VB.TextBox txt2SD 
      Height          =   285
      Left            =   9600
      TabIndex        =   14
      Top             =   5370
      Width           =   915
   End
   Begin VB.TextBox txtMean 
      Height          =   285
      Left            =   9600
      TabIndex        =   12
      Top             =   2880
      Width           =   915
   End
   Begin VB.TextBox txtLow 
      Height          =   285
      Left            =   9600
      TabIndex        =   10
      Top             =   4650
      Width           =   915
   End
   Begin VB.TextBox txtHigh 
      Height          =   285
      Left            =   9600
      TabIndex        =   9
      Top             =   930
      Width           =   915
   End
   Begin VB.ComboBox cmbControl 
      Height          =   315
      Left            =   2130
      TabIndex        =   6
      Text            =   "cmbControl"
      Top             =   450
      Width           =   2535
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9630
      Picture         =   "frmBioQC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5820
      Width           =   915
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      DrawWidth       =   3
      Height          =   4005
      Left            =   2100
      ScaleHeight     =   3945
      ScaleWidth      =   7365
      TabIndex        =   3
      Top             =   930
      Width           =   7425
   End
   Begin VB.ListBox List1 
      Height          =   5820
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   2
      Top             =   450
      Width           =   1635
   End
   Begin VB.OptionButton oUrine 
      Caption         =   "Urine"
      Height          =   255
      Left            =   1050
      TabIndex        =   1
      Top             =   150
      Width           =   825
   End
   Begin VB.OptionButton oSerum 
      Alignment       =   1  'Right Justify
      Caption         =   "Serum"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Value           =   -1  'True
      Width           =   765
   End
   Begin VB.Label lblDeleteValue 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDeleteValue"
      Height          =   255
      Left            =   7590
      TabIndex        =   29
      Top             =   5820
      Width           =   1245
   End
   Begin VB.Label lblDeleteTime 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDeleteTime"
      Height          =   255
      Left            =   6060
      TabIndex        =   28
      Top             =   6090
      Width           =   1515
   End
   Begin VB.Label lblDeleteParameter 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDeleteParameter"
      Height          =   255
      Left            =   6060
      TabIndex        =   27
      Top             =   5820
      Width           =   1515
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "'<ALT>D' to delete"
      Height          =   525
      Left            =   5190
      TabIndex        =   26
      Top             =   5820
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "2 SD"
      Height          =   195
      Left            =   9840
      TabIndex        =   13
      Top             =   5190
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Mean"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   9870
      TabIndex        =   11
      Top             =   2670
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Low"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9870
      TabIndex        =   8
      Top             =   4440
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "High"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9840
      TabIndex        =   7
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Control"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmBioQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private picMin As Single
Private picMax As Single

Private Type udtDateValue
    strTime As String
    strDate As String
    strValue As String
    lngX As Long
    lngY As Long
End Type

Private DateValues() As udtDateValue

Private fnt As CLogFont

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Sub DrawDates()

          Dim X As Long
          Dim hFont As Long
          Dim intDec As Integer
          Dim n As Integer
         
7830      picDates.AutoRedraw = False

7840      Set fnt = New CLogFont
7850      Set fnt.LogFont = picDates.Font

7860      fnt.Rotation = 90
7870      hFont = SelectObject(picDates.hdc, fnt.Handle)

7880      intDec = Val(lblPrevious) / 31

7890      picDates.Cls

7900      n = Val(lblPrevious)

7910      For X = 0 To picDates.width Step picDates.width / 31
7920          picDates.CurrentX = X
7930          picDates.CurrentY = picDates.height - 100
7940          picDates.Print Format(dtTo - n, "dd MMM")
7950          n = n - intDec
7960      Next

7970      Call SelectObject(picDates.hdc, hFont)
7980      picDates.AutoRedraw = True

End Sub


Private Sub DrawGraph()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleType As String
          Dim Days As Long
          Dim PointsPerDay As Single
          Dim ThisX As Single
          Dim ThisDate As String
          Dim YPoints As Single
          Dim PointsPerY As Single
          Dim Counter As Integer
          Dim StartDate As String
          Dim Diffdays As Integer
          Dim n As Integer
          Dim mean As Single
          Dim Alias As String

7990      On Error GoTo DrawGraph_Error

8000      If List1.ListIndex = -1 Then Exit Sub

8010      Days = Val(lblPrevious)
8020      StartDate = Format$(dtTo - Days, "dd/mmm/yyyy")
8030      PointsPerDay = pb.width / (Days + 1)
8040      ThisX = PointsPerDay / 2

8050      ColourBarVert pb

8060      SampleType = IIf(oSerum, "S", "U")

8070      sql = "Select AliasName from BioQCDefs where ControlName = '" & cmbControl & "'"
8080      Set tb = New Recordset
8090      RecOpenServer 0, tb, sql
8100      If tb.EOF Then Exit Sub
8110      Alias = tb!AliasName & ""

8120      sql = "Delete from BiochemistryQC where Result like '%-%'"
8130      Cnxn(0).Execute sql

8140      sql = "Select RunDate, RunTime, result from BiochemistryQC where " & _
              "SampleType = '" & SampleType & "' " & _
              "and cast(result as float) > " & picMin & " " & _
              "and cast(result as float) < " & picMax & " " & _
              "and AliasName = '" & Alias & "' " & _
              "and Rundate between '" & Format(dtTo - Days, "dd/mmm/yyyy") & "' " & _
              "and '" & Format(dtTo, "dd/mmm/yyyy") & "' " & _
              "and Code = '" & lstCode.List(List1.ListIndex) & "'"
8150      Set tb = New Recordset
8160      RecOpenClient 0, tb, sql

8170      If tb.EOF Then Exit Sub

8180      ReDim DateValues(0 To tb.RecordCount)

8190      YPoints = Abs(picMax - picMin)
8200      If YPoints <= 0 Then Exit Sub

8210      PointsPerY = (pb.height / YPoints)
8220      pb.DrawWidth = 3

8230      Counter = 0
8240      Do While Not tb.EOF
              'If ThisDate = "" Then
8250          ThisDate = Format(tb!Rundate, "dd/mmm/yyyy")
              'End If
              '  If ThisDate <> Format(tb!RunDate, "dd/mmm/yyyy") Then
              '    ThisDate = Format(tb!RunDate, "dd/mmm/yyyy")
8260          Diffdays = DateDiff("D", StartDate, ThisDate)
8270          ThisX = (PointsPerDay * Diffdays) + (PointsPerDay / 2)
              '  End If
8280          If ThisX <= pb.width Then
8290              DateValues(Counter).strTime = tb!RunTime
8300              DateValues(Counter).strDate = ThisDate
8310              DateValues(Counter).strValue = tb!Result & ""
8320              DateValues(Counter).lngX = ThisX
8330              DateValues(Counter).lngY = pb.height - (PointsPerY * (Val(tb!Result & "") - picMin))
8340          End If
8350          If optShow(0) Or optShow(2) Then
8360              pb.PSet (ThisX, pb.height - (PointsPerY * (Val(tb!Result & "") - picMin)))
8370          End If
8380          tb.MoveNext
8390          Counter = Counter + 1
8400      Loop
8410      pb.DrawWidth = 1

8420      If optShow(1) Or optShow(2) Then
        
8430          ThisDate = ""
8440          mean = 0
8450          Counter = 0
8460          For n = 0 To UBound(DateValues)
8470              If ThisDate = DateValues(n).strDate Or ThisDate = "" Then
8480                  ThisDate = DateValues(n).strDate
8490                  mean = mean + Val(DateValues(n).strValue)
8500                  Counter = Counter + 1
8510              Else
8520                  mean = mean / Counter
8530                  ThisX = DateValues(n - 1).lngX
8540                  pb.Circle (ThisX, pb.height - (PointsPerY * (mean - picMin))), 30, vbBlue
8550                  Counter = 1
8560                  mean = Val(DateValues(n).strValue)
8570                  ThisDate = DateValues(n).strDate
8580              End If
8590          Next
8600      End If

8610      Exit Sub

DrawGraph_Error:

          Dim strES As String
          Dim intEL As Integer

8620      intEL = Erl
8630      strES = Err.Description
8640      LogError "fBioQC", "DrawGraph", intEL, strES, sql


End Sub

Private Sub FillLevels()

          Dim tb As Recordset
          Dim sql As String

8650      On Error GoTo FillLevels_Error

8660      If List1.ListIndex = -1 Then
8670          iMsg "Select Parameter"
8680          txtHigh = ""
8690          txtLow = ""
8700          txtMean = ""
8710          txt2SD = ""
8720          Exit Sub
8730      End If

8740      sql = "Select * from BioQCDefs where " & _
              "ControlName = '" & cmbControl & "' " & _
              "and ParameterName = '" & List1 & "'"
8750      Set tb = New Recordset
8760      RecOpenServer 0, tb, sql
8770      If Not tb.EOF Then
8780          txtHigh = tb!High & ""
8790          txtLow = tb!Low & ""
8800          txtMean = tb!mean & ""
8810          txt2SD = tb!SD & ""
8820          picMax = Val(tb!mean & "") + (1.5 * Val(tb!SD & ""))
8830          picMin = Val(tb!mean & "") - (1.5 * Val(tb!SD & ""))
8840      Else
8850          txtHigh = ""
8860          picMax = 9999
8870          txtLow = ""
8880          picMin = 0
8890          txtMean = ""
8900          txt2SD = ""
8910      End If

8920      DrawGraph

8930      Exit Sub

FillLevels_Error:

          Dim strES As String
          Dim intEL As Integer

8940      intEL = Erl
8950      strES = Err.Description
8960      LogError "fBioQC", "FillLevels", intEL, strES, sql


End Sub

Private Function FillList() As Boolean
          'Returns True if the list was modified -
          'either added to or removed from

          Dim tb As Recordset
          Dim sql As String
          Dim SampleType As String
          Dim Found As Boolean
          Dim n As Integer
          Dim strLongName As String

8970      On Error GoTo FillList_Error

8980      FillList = False

8990      SampleType = IIf(oSerum, "S", "U")

9000      sql = "SELECT DISTINCT Q.Code, D.LongName " & _
              "FROM BiochemistryQC Q JOIN BioTestDefinitions D " & _
              "WHERE " & _
              "SampleType = '" & SampleType & "' " & _
              "AND Rundate BETWEEN '" & Format(dtTo - Val(lblPrevious), "dd/mmm/yyyy") & "' " & _
              "AND '" & Format(dtTo, "dd/mmm/yyyy") & "'"
9010      Set tb = New Recordset
9020      RecOpenServer 0, tb, sql
9030      If tb.EOF Then
9040          List1.Clear
9050          lstCode.Clear
9060          FillList = True
9070          Exit Function
9080      End If

9090      Do While Not tb.EOF
        
9100          strLongName = tb!LongName & ""
9110          Found = False
9120          For n = 0 To List1.ListCount - 1
9130              If List1.List(n) = strLongName Then
9140                  Found = True
9150                  Exit For
9160              End If
9170          Next
9180          If Not Found Then
9190              List1.AddItem strLongName
9200              lstCode.AddItem tb!Code & ""
9210              FillList = True
9220          End If
9230          tb.MoveNext

9240      Loop

9250      For n = List1.ListCount - 1 To 0 Step -1
9260          Found = False
9270          tb.MoveFirst
9280          Do While Not tb.EOF
9290              strLongName = tb!LongName & ""
9300              If List1.List(n) = strLongName Then
9310                  Found = True
9320                  Exit Do
9330              End If
9340              tb.MoveNext
9350          Loop
9360          If Not Found Then
9370              List1.RemoveItem n
9380              lstCode.RemoveItem n
9390              FillList = True
9400          End If
9410      Next

9420      Exit Function

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

9430      intEL = Erl
9440      strES = Err.Description
9450      LogError "fBioQC", "FillList", intEL, strES, sql

End Function

Private Function GetAlias() As String

          Dim tb As Recordset
          Dim sql As String

9460      On Error GoTo GetAlias_Error

9470      GetAlias = ""

9480      sql = "SELECT AliasName FROM BioQCDefs WHERE " & _
              "ControlName = '" & cmbControl & "'"
9490      Set tb = New Recordset
9500      RecOpenServer 0, tb, sql
9510      If Not tb.EOF Then
9520          If Trim$(tb!AliasName & "") <> "" Then
9530              GetAlias = tb!AliasName
9540          End If
9550      End If

9560      Exit Function

GetAlias_Error:

          Dim strES As String
          Dim intEL As Integer

9570      intEL = Erl
9580      strES = Err.Description
9590      LogError "fBioQC", "GetAlias", intEL, strES, sql

End Function

Private Sub UpdateValues()

          Dim tb As Recordset
          Dim sql As String
          Dim Alias As String

9600      On Error GoTo UpdateValues_Error

9610      If List1.ListIndex = -1 Then
9620          iMsg "Select Parameter"
9630          txtHigh = ""
9640          txtLow = ""
9650          Exit Sub
9660      End If

9670      Alias = GetAlias()
9680      If Alias = "" Then Exit Sub

9690      sql = "Select * from BioQCDefs where " & _
              "AliasName = '" & Alias & "' " & _
              "and ParameterName = '" & List1 & "'"
9700      Set tb = New Recordset
9710      RecOpenServer 0, tb, sql
9720      If tb.EOF Then
9730          tb.AddNew
9740          tb!ControlName = cmbControl
9750          tb!ParameterName = List1
9760      End If
9770      tb!High = Val(txtHigh)
9780      tb!Low = Val(txtLow)
9790      tb!mean = Val(txtMean)
9800      tb!SD = Val(txt2SD)
9810      tb.Update

9820      If Val(txtLow) <> 0 Then
9830          picMin = Val(txtLow)
9840      Else
9850          picMin = 0
9860      End If

9870      DrawGraph

9880      Exit Sub

UpdateValues_Error:

          Dim strES As String
          Dim intEL As Integer

9890      intEL = Erl
9900      strES = Err.Description
9910      LogError "fBioQC", "UpdateValues", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

9920      Unload Me

End Sub


Private Sub cmbControl_Click()

9930      FillLevels

End Sub


Private Sub dtTo_CloseUp()

9940      DrawDates
9950      If Not FillList() Then
9960          DrawGraph
9970      End If

End Sub

Private Sub Form_Activate()

9980      picDates.Refresh

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim Alias As String
          Dim SampleType As String

9990      On Error GoTo Form_KeyUp_Error

10000     Debug.Print KeyCode, Shift

10010     If KeyCode <> 68 Or Shift <> 4 Then Exit Sub '<ALT> D
10020     If lblDelete.Visible = False Or _
              lblDeleteParameter.Visible = False Or _
              lblDeleteValue.Visible = False Or _
              lblDeleteTime.Visible = False Then Exit Sub

10030     If List1.ListIndex = -1 Then Exit Sub

10040     SampleType = IIf(oSerum, "S", "U")

10050     sql = "Select AliasName from BioQCDefs where ControlName = '" & cmbControl & "'"
10060     Set tb = New Recordset
10070     RecOpenServer 0, tb, sql
10080     If tb.EOF Then Exit Sub
10090     Alias = tb!AliasName & ""

10100     sql = "Delete from BiochemistryQC where " & _
              "SampleType = '" & SampleType & "' " & _
              "and Result = '" & lblDeleteValue & "' " & _
              "and AliasName = '" & Alias & "' " & _
              "and RunTime = '" & Format(lblDeleteTime, "dd/mmm/yyyy hh:mm") & "' " & _
              "and Code = '" & lstCode.List(List1.ListIndex) & "'"
10110     Cnxn(0).Execute sql

10120     Exit Sub

Form_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

10130     intEL = Erl
10140     strES = Err.Description
10150     LogError "fBioQC", "Form_KeyUp", intEL, strES, sql


End Sub


Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

10160     On Error GoTo Form_Load_Error

10170     ReDim DateValues(0 To 0) As udtDateValue

10180     dtTo = Format(Now, "dd/mm/yyyy")
10190     DrawDates

10200     sql = "Select distinct ControlName from BioQCDefs"
10210     Set tb = New Recordset
10220     RecOpenServer 0, tb, sql

10230     cmbControl.Clear

10240     Do While Not tb.EOF
10250         cmbControl.AddItem tb!ControlName & ""
10260         tb.MoveNext
10270     Loop

10280     FillList

10290     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

10300     intEL = Erl
10310     strES = Err.Description
10320     LogError "fBioQC", "Form_Load", intEL, strES, sql


End Sub







Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10330     lblDelete.Visible = False
10340     lblDeleteParameter.Visible = False
10350     lblDeleteValue.Visible = False
10360     lblDeleteTime.Visible = False

End Sub

Private Sub List1_Click()

10370     FillLevels

End Sub


Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10380     lblDelete.Visible = False
10390     lblDeleteParameter.Visible = False
10400     lblDeleteValue.Visible = False
10410     lblDeleteTime.Visible = False

End Sub


Private Sub optShow_Click(Index As Integer)

10420     DrawGraph

End Sub

Private Sub pb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim i As Integer
          Dim CurrentDistance As Long
          Dim BestDistance As Long
          Dim BestIndex As Integer

10430     On Error GoTo pbmm

10440     lblDelete.Visible = False
10450     lblDeleteParameter.Visible = False
10460     lblDeleteValue.Visible = False
10470     lblDeleteTime.Visible = False
10480     If optShow(0) Then
10490         BestIndex = -1
10500         BestDistance = 99999
10510         For i = 0 To UBound(DateValues)
10520             CurrentDistance = ((X - DateValues(i).lngX) ^ 2 + (Y - DateValues(i).lngY) ^ 2) ^ (1 / 2)
10530             If i = 0 Or CurrentDistance < BestDistance Then
10540                 BestDistance = CurrentDistance
10550                 BestIndex = i
10560             End If
10570         Next
        
10580         If BestIndex <> -1 Then
10590             If DateValues(BestIndex).strTime <> "" Then
10600                 pb.ToolTipText = Format$(DateValues(BestIndex).strTime, "dd/MM/yy hh:mm") & " " & DateValues(BestIndex).strValue
        
10610                 lblDeleteParameter.Caption = List1
10620                 lblDeleteTime.Caption = Format$(DateValues(BestIndex).strTime, "dd/MM/yy hh:mm")
10630                 lblDeleteValue.Caption = DateValues(BestIndex).strValue
10640                 lblDelete.Visible = True
10650                 lblDeleteParameter.Visible = True
10660                 lblDeleteValue.Visible = True
10670                 lblDeleteTime.Visible = True
10680             End If
10690         End If
10700     End If

10710     Exit Sub

pbmm:
10720     Exit Sub

End Sub



Private Sub picDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10730     lblDelete.Visible = False
10740     lblDeleteParameter.Visible = False
10750     lblDeleteValue.Visible = False
10760     lblDeleteTime.Visible = False

End Sub


Private Sub txt2SD_KeyUp(KeyCode As Integer, Shift As Integer)

10770     UpdateValues

End Sub

Private Sub txtHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10780     UpdateValues

End Sub

Private Sub txtLow_KeyUp(KeyCode As Integer, Shift As Integer)

10790     UpdateValues

End Sub


Private Sub txtMean_KeyUp(KeyCode As Integer, Shift As Integer)

10800     UpdateValues

End Sub

Private Sub udPrevious_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10810     DrawDates
10820     If Not FillList() Then
10830         DrawGraph
10840     End If

End Sub


