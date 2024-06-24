VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotValidatedPrinted 
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   12060
      Picture         =   "frmNotValidatedPrinted.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3300
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrintSamples 
      Appearance      =   0  'Flat
      Caption         =   "&Print Sample"
      Height          =   975
      Left            =   12060
      Picture         =   "frmNotValidatedPrinted.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5640
      Width           =   1245
   End
   Begin VB.ComboBox cmbSortBy 
      Height          =   315
      ItemData        =   "frmNotValidatedPrinted.frx":0614
      Left            =   8670
      List            =   "frmNotValidatedPrinted.frx":0616
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1500
      Width           =   2145
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   975
      Left            =   12060
      Picture         =   "frmNotValidatedPrinted.frx":0618
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7260
      Width           =   1245
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print List"
      Height          =   975
      Left            =   12060
      Picture         =   "frmNotValidatedPrinted.frx":0922
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2175
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   4455
      Begin VB.OptionButton optValidNotPrint 
         Caption         =   "Validated Not Printed"
         Height          =   195
         Left            =   300
         TabIndex        =   30
         Top             =   1035
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.OptionButton optNotValidatedPrinted 
         Caption         =   "Neither validated nor printed"
         Height          =   315
         Left            =   300
         TabIndex        =   17
         Top             =   1290
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.OptionButton optNotPrinted 
         Caption         =   "Not Printed"
         Height          =   195
         Left            =   2850
         TabIndex        =   16
         Top             =   750
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.OptionButton optNotValidated 
         Caption         =   "Not validated"
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   765
         TabIndex        =   18
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218365953
         CurrentDate     =   37096
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   19
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218365953
         CurrentDate     =   37096
      End
      Begin VB.Label lblTo 
         Caption         =   "To"
         Height          =   195
         Left            =   2430
         TabIndex        =   21
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblFrom 
         Caption         =   "From"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   1635
      Index           =   0
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   3690
      Begin VB.OptionButton optAllDisp 
         Caption         =   "All Disciplines"
         Height          =   255
         Left            =   2100
         TabIndex        =   29
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optBio 
         Caption         =   "Biochemistry"
         Height          =   255
         Left            =   585
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optHaem 
         Caption         =   "Haematology"
         Height          =   255
         Left            =   585
         TabIndex        =   11
         Top             =   615
         Width           =   1485
      End
      Begin VB.OptionButton optCoag 
         Caption         =   "Coagulation"
         Height          =   255
         Left            =   585
         TabIndex        =   10
         Top             =   930
         Width           =   1485
      End
      Begin VB.OptionButton optExt 
         Caption         =   "External"
         Height          =   255
         Left            =   2100
         TabIndex        =   9
         Top             =   930
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.OptionButton optEnd 
         Caption         =   "Endocrinology"
         Height          =   255
         Left            =   585
         TabIndex        =   8
         Top             =   1245
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.OptionButton optImm 
         Caption         =   "Immunology"
         Height          =   255
         Left            =   2100
         TabIndex        =   7
         Top             =   630
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.OptionButton optBG 
         Caption         =   "Blood Gas"
         Height          =   255
         Left            =   2100
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optHisto 
         Caption         =   "Histology"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   615
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optCyto 
         Caption         =   "Cytology"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   930
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optMicro 
         Caption         =   "Microbiology"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optSemen 
         Caption         =   "Semen Analysis"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   1245
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      Caption         =   "Start"
      Height          =   975
      Left            =   12060
      Picture         =   "frmNotValidatedPrinted.frx":0C2C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   900
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6060
      Left            =   240
      TabIndex        =   22
      Top             =   2175
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10689
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmNotValidatedPrinted.frx":0F36
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   12060
      TabIndex        =   32
      Top             =   4275
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Sort By"
      Height          =   255
      Left            =   8700
      TabIndex        =   27
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label lblClick 
      Caption         =   "Please click on the sample id to view details"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   10095
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   8340
      Width           =   480
   End
End
Attribute VB_Name = "frmNotValidatedPrinted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub SetFormCaption()

8530  On Error GoTo SetFormCaption_Error

8540  If optBio.Value = True Then
8550      Me.Caption = "NetAcquire - " & optBio.Caption & " Unvalidated / Not Printed Samples"
8560  ElseIf optHaem.Value = True Then
8570      Me.Caption = "NetAcquire - " & optHaem.Caption & " Unvalidated / Not Printed Samples"
8580  ElseIf optCoag.Value = True Then
8590      Me.Caption = "NetAcquire - " & optCoag.Caption & " Unvalidated / Not Printed Samples"
8600  ElseIf optEnd.Value = True Then
8610      Me.Caption = "NetAcquire - " & optEnd.Caption & " Unvalidated / Not Printed Samples"
8620  ElseIf optImm.Value = True Then
8630      Me.Caption = "NetAcquire - " & optImm.Caption & " Unvalidated / Not Printed Samples"
8640  ElseIf optExt.Value = True Then
8650      Me.Caption = "NetAcquire - " & optExt.Caption & " Unvalidated / Not Printed Samples"
8660  ElseIf optBG.Value = True Then
8670      Me.Caption = "NetAcquire - " & optBG.Caption & " Unvalidated / Not Printed Samples"
8680  ElseIf optHisto.Value = True Then
8690      Me.Caption = "NetAcquire - " & optHisto.Caption & " Unvalidated / Not Printed Samples"
8700  ElseIf optCyto.Value = True Then
8710      Me.Caption = "NetAcquire - " & optCyto.Caption & " Unvalidated / Not Printed Samples"
8720  ElseIf optMicro.Value = True Then
8730      Me.Caption = "NetAcquire - " & optMicro.Caption & " Unvalidated / Not Printed Samples"
8740  ElseIf optSemen.Value = True Then
8750      Me.Caption = "NetAcquire - " & optSemen.Caption & " Unvalidated / Not Printed Samples"
8760  ElseIf optAllDisp.Value = True Then
8770      Me.Caption = "NetAcquire - All Unvalidated / Not Printed Samples"
8780  End If
8790  Exit Sub

SetFormCaption_Error:

      Dim strES As String
      Dim intEL As Integer

8800  intEL = Erl
8810  strES = Err.Description
8820  LogError "frmDaily", "SetFormCaption", intEL, strES

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FillG
' Author    : XPMUser
' Date      : 11/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FillG()

      Dim tb As New Recordset
      Dim s As String
      Dim sql As String
      Dim Asql As String
      Dim Bsql As String
      Dim OldSampleID As String
      Dim NewSampleID As String
      Dim Disc As String
      Dim TestColumn As String
      Dim ResultColumn As String
      Dim TableName As String
      Dim DateColumn As String
      Dim Selection As String
      Dim SortBy As String

      Dim DisCode As String
      Dim DispName As String

8830  On Error GoTo FillG_Error


8840  On Error GoTo FillG_Error

8850  ClearFGrid g

8860  DateColumn = "RunDate"

8870  g.ColWidth(6) = 0
8880  DisCode = ""
8890  DispName = ""
8900  If optBio Then
8910      Disc = "Bio"
8920      TestColumn = "ShortName"
8930      TableName = "BioResults"
8940      DisCode = "B"
8950      DispName = "Biochemistry"
8960  ElseIf optHaem Then
8970      Disc = "Haem"
8980      TestColumn = "AnalyteName"
8990      ResultColumn = "RBC"
9000      TableName = "HaemResults"

9010      DisCode = "H"
9020      DispName = "Haematology"
9030  ElseIf optCoag Then
9040      Disc = "Coag"
9050      TestColumn = "TestName"
9060      TableName = "CoagResults"
9070      DisCode = "D"
9080      DispName = "Coagulation"
9090  ElseIf optExt Then
9100      Disc = "Ext"
9110      TestColumn = "Analyte"
9120      TableName = "ExtResults"
9130      DateColumn = "RetDate"
9140      DisCode = "X"
9150      DispName = "External"
9160  ElseIf optEnd Then
9170      Disc = "End"
9180      TestColumn = "ShortName"
9190      TableName = "EndResults"
9200      DisCode = "E"
9210      DispName = "Endocrinology"
9220  ElseIf optImm Then
9230      Disc = "Imm"
9240      TestColumn = "ShortName"
9250      TableName = "ImmResults"

9260      DisCode = "I"
9270      DispName = "Immunology"
9280  ElseIf optBG Then
9290      Disc = "Bga"
9300      TestColumn = "ShortName"
9310      TableName = "BgaResults"

9320      DisCode = "B"
9330      DispName = "Blood Gas"
9340  ElseIf optCyto Then
9350      Disc = "Cyto"
9360      TestColumn = ""
9370      DisCode = ""
9380      DispName = ""

9390      TestColumn = ""
9400      DispName = "Cytology"
9410  ElseIf optHisto Then
9420      Disc = "Histo"
9430      TestColumn = ""
9440      DisCode = ""
9450      DispName = "Histology"
9460  ElseIf optMicro Then
9470      TableName = "PrintValidLog"
9480      DisCode = ""
9490      DispName = ""
9500  ElseIf optSemen Then
9510      Disc = "Semen"
9520      TestColumn = ""
9530      ResultColumn = "Motility"
9540      TableName = "SemenResults"
9550      DateColumn = "DateTimeOfRecord"
9560      DisCode = ""
9570      DispName = "Semen"
9580  ElseIf optAllDisp Then

9590      FillGAll
9600      Exit Sub
9610  Else
9620      g.Visible = True
9630      iMsg "No Discipline Choosen!"
9640      Exit Sub
9650  End If

9660  If optNotValidated Then
9670      Selection = "AND ISNULL(R.printed,0) = 0"
9680  ElseIf optNotPrinted Then
9690      Selection = "AND ISNULL(R.printed,0) = 0"
9700  ElseIf optValidNotPrint Then
9710      Selection = "AND R.Valid = 1 AND ISNULL(R.printed,0) = 0"
9720  ElseIf optNotValidatedPrinted Then
9730      Selection = "AND R.Valid = 0 AND ISNULL(R.printed,0) = 0"
9740  End If

9750  If optBio Or optCoag Or optEnd Or optImm Or optBG Or optExt Or optSemen Then
9760      sql = "SELECT DISTINCT R.SampleID, D.DateTimeDemographics as EntryTime, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
                "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                "INNER JOIN " & Disc & "TestDefinitions T ON R.Code = T.Code " & _
                "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom.Value, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo.Value + 1, "dd/mmm/yyyy") & "' " & _
                "AND T.Printable = '1' " & _
                Selection & " " & _
                "AND R.SampleID NOT IN (SELECT SampleID FROM PrintPending)"
                
9770  ElseIf optHaem Then
9780      sql = "SELECT DISTINCT R.SampleID, D.DateTimeDemographics as EntryTime, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
                "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy") & "' " & _
                Selection & " " & _
                "AND R.SampleID NOT IN (SELECT SampleID FROM PrintPending)"

9790  End If
9800  SortBy = ""
9810  If cmbSortBy <> "" Then
9820      If cmbSortBy.Text = "Earliest Samples First" Then
9830          SortBy = " ORDER BY EntryTime asc"
9840      Else
9850          SortBy = " ORDER BY D." & Replace(cmbSortBy, "-", ",D.")
9860      End If

9870  End If
9880  sql = sql & " " & SortBy
9890  If optHaem.Value = False Then
9900      sql = sql & " "
9910  End If
9920  Set tb = New Recordset
9930  RecOpenServer 0, tb, sql
9940  While Not tb.EOF
9950      If TableName <> "" And checkAllResultsValid(tb!SampleID, TableName) = True Then
9960          If optHisto Then
9970              s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (sysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H"
9980          ElseIf optCyto Then
9990              s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (sysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "C"
10000         ElseIf optMicro Then
10010             s = Val(tb!SampleID) ' - sysOptMicroOffset(0)
10020         ElseIf optSemen Then
10030             s = Val(tb!SampleID) - sysOptSemenOffset(0)
10040         Else
10050             s = tb!SampleID
10060         End If

10070         s = s & vbTab & _
                  tb!PatName & vbTab & _
                  tb!Chart & vbTab & _
                  tb!GP & vbTab & _
                  tb!Ward & vbTab & _
                  tb!Clinician & "" & vbTab & _
                  DisCode & vbTab & _
                  DispName

10080         g.AddItem s
10090     End If
10100     tb.MoveNext

10110 Wend


10120 FixG g

10130 If g.Rows > 2 Then
10140     lblTotal = "Total samples : " & g.Rows - 1
10150 Else
10160     lblTotal = ""
10170 End If

10180 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

10190 intEL = Erl
10200 strES = Err.Description
10210 LogError "frmDaily", "FillG", intEL, strES, sql


End Sub


Function checkAllResultsValid(SampleID As String, TableName As String) As Boolean
      Dim tb As Recordset
      Dim sql As String

10220 On Error GoTo checkAllResultsValid_Error

10230 sql = " Select * from  " & TableName & " where sampleid = " & SampleID & " and valid=0"

10240 Set tb = New Recordset
10250 RecOpenServer 0, tb, sql
10260 If tb.EOF Then
10270     checkAllResultsValid = True
10280 Else
10290     checkAllResultsValid = False
10300 End If


10310 Exit Function

checkAllResultsValid_Error:
      Dim strES As String
      Dim intEL As Integer

10320 intEL = Erl
10330 strES = Err.Description
10340 LogError "frmNotValidatedPrinted", "checkAllResultsValid", intEL, strES, sql

End Function
'---------------------------------------------------------------------------------------
' Procedure : FillGAll
' Author    : XPMUser
' Date      : 06/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Sub FillGAll()
'Dim TableName As String
'Dim Disc As String
'Dim ResultColumn As String
'Dim TestColumn As String
'Dim Selection As String
'Dim DateColumn As String
'Dim tb As ADODB.Recordset
'Dim SortBy As String
'Dim sql As String
'Dim s As String
'Dim i As Integer
'Dim DisCode As String
'Dim DispName As String
'Dim sqlUnion As String
'On Error GoTo FillGAll_Error
'
'ClearFGrid g
'PrintSamplesVisibilty (1)
'g.ColWidth(6) = 0
'DisCode = ""
'DispName = ""
'sqlUnion = " A UNION ALL "
'For i = 1 To 11
'    DateColumn = "RunDate"
'    If i = 1 Then
'        Disc = "Bio"
'        TestColumn = "ShortName"
'        TableName = "BioResults"
'        DisCode = "B"
'        DispName = "Biochemistry"
'    ElseIf i = 2 Then
'        Disc = "Haem"
'        TestColumn = "AnalyteName"
'        ResultColumn = "RBC"
'        TableName = "HaemResults"
'        DisCode = "H"
'        DispName = "Haematology"
'    ElseIf i = 3 Then
'        Disc = "Coag"
'        TestColumn = "TestName"
'        TableName = "CoagResults"
'        DisCode = "D"
'        DispName = "Coagulation"
'    ElseIf i = 4 Then
'        Disc = "Ext"
'        TestColumn = "Analyte"
'        TableName = "ExtResults"
'        DateColumn = "RetDate"
'        DisCode = "X"
'        DispName = "External"
'    ElseIf i = 5 Then
'        Disc = "End"
'        TestColumn = "ShortName"
'        TableName = "EndResults"
'        DisCode = "E"
'        DispName = "Endocrinology"
'    ElseIf i = 6 Then
'        Disc = "Imm"
'        TestColumn = "ShortName"
'        TableName = "ImmResults"
'        DisCode = "I"
'        DispName = "Immunology"
'    ElseIf i = 7 Then
'        Disc = "Bga"
'        TestColumn = "ShortName"
'        TableName = "BgaResults"
'        DisCode = "B"
'        DispName = "Blood Gas"
'    ElseIf i = 8 Then
'        Disc = "Cyto"   ' NOT EXISTS IN DATABASE
'        TestColumn = ""
'        DispName = "Cytology"
'    ElseIf i = 9 Then
'        Disc = "Histo"  ' NOT EXISTS IN DATABASE
'        TestColumn = ""
'        DispName = "Histology"
'    ElseIf i = 10 Then
'        TableName = "PrintValidLog"
'        DispName = ""
'    ElseIf i = 11 Then
'        Disc = "Semen"
'        TestColumn = ""
'        ResultColumn = "Motility"
'        TableName = "SemenResults"
'        DateColumn = "DateTimeOfRecord"
'        DispName = "Semen"
'    End If
'
'
'    If optNotValidated Then
'        Selection = "AND R.Valid = 0"
'    ElseIf optNotPrinted Then
'        Selection = "AND R.Printed = 0"
'    ElseIf optValidNotPrint Then
'        Selection = "AND R.Valid = 1 AND R.Printed = 0"
'    ElseIf optNotValidatedPrinted Then
'        Selection = "AND R.Valid = 0 AND R.Printed = 0"
'    End If
'    If i <> 4 And i <> 7 And i <> 11 And i <> 8 And i <> 9 And i <> 10 Then
'        If i = 1 Or i = 3 Or i = 5 Or i = 6 Or i = 7 Or i = 4 Or i = 11 Then
'            sql = "SELECT DISTINCT R.SampleID, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
             '                  "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
             '                  "INNER JOIN " & Disc & "TestDefinitions T ON R.Code = T.Code " & _
             '                  "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy") & "' " & _
             '                  "AND T.Printable = '1' " & _
             '                  Selection
'        ElseIf i = 2 Then
'            sql = "SELECT DISTINCT R.SampleID, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
             '                  "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
             '                  "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy") & "' " & _
             '                  Selection
'
'        End If
'        SortBy = ""
'        If cmbSortBy.ListIndex > 0 Then
'            If cmbSortBy <> "" Then
'                If InStr(1, cmbSortBy, "-") > 0 Then
'                    SortBy = " ORDER BY D." & Replace(cmbSortBy, "-", ",D.")
'                Else
'                    SortBy = " ORDER BY  D." & cmbSortBy
'                End If
'            End If
'            sql = sql & " " & SortBy
'        ElseIf optHaem.Value = False Then
'            'sql = sql & " ORDER BY D.DateTimeDemographics "
'            sql = sql & " "
'        End If
'
'
''        sqlUnion = sqlUnion & sql
'
'        Set tb = New Recordset
'        RecOpenServer 0, tb, sql
'        '          If tb.State = 0 Then Exit Sub
'        While Not tb.EOF
'            If optHisto Then
'                s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (sysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H"
'            ElseIf optCyto Then
'                s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (sysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "C"
'            ElseIf optMicro Then
'                s = Val(tb!SampleID) - sysOptMicroOffset(0)
'            ElseIf optSemen Then
'                s = Val(tb!SampleID) - sysOptSemenOffset(0)
'            Else
'                s = tb!SampleID
'            End If
'
'            s = s & vbTab & _
             '                tb!PatName & vbTab & _
             '                tb!Chart & vbTab & _
             '                tb!GP & vbTab & _
             '                tb!Ward & vbTab & _
             '                tb!Clinician & "" & vbTab & _
             '                DisCode & vbTab & _
             '                DispName
'
'            g.AddItem s
'            tb.MoveNext
'
'        Wend
'    End If
'Next i
'
'
'
'
'
'
'
'
'FixG g
'
'If g.Rows > 2 Then
'    lblTotal = "Total samples : " & g.Rows - 1
'Else
'    lblTotal = ""
'End If
'
'
'Exit Sub
'
'
'FillGAll_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmNotValidatedPrinted", "FillGAll", intEL, strES, sql
'End Sub

Sub FillGAll()
      Dim TableName As String
      Dim Disc As String
      Dim ResultColumn As String
      Dim TestColumn As String
      Dim Selection As String
      Dim DateColumn As String
      Dim tb As ADODB.Recordset
      Dim SortBy As String
      Dim sql As String
      Dim s As String
      Dim i As Integer
      Dim DisCode As String
      Dim DispName As String
      Dim sqlUnion As String
      Dim DeptClms As String

10350 On Error GoTo FillGAll_Error

10360 ClearFGrid g
10370 PrintSamplesVisibilty (1)
10380 g.ColWidth(6) = 0

10390 DisCode = ""
10400 DispName = ""

10410 sqlUnion = " Dummy "


10420 For i = 1 To 11
10430     DateColumn = "RunDate"
10440     If i = 1 Then
10450         Disc = "Bio"
10460         TestColumn = "ShortName"
10470         TableName = "BioResults"
10480         DisCode = "B"
10490         DispName = "Biochemistry"
10500     ElseIf i = 2 Then
10510         Disc = "Haem"
10520         TestColumn = "AnalyteName"
10530         ResultColumn = "RBC"
10540         TableName = "HaemResults"
10550         DisCode = "H"
10560         DispName = "Haematology"
10570     ElseIf i = 3 Then
10580         Disc = "Coag"
10590         TestColumn = "TestName"
10600         TableName = "CoagResults"
10610         DisCode = "D"
10620         DispName = "Coagulation"
10630     ElseIf i = 4 Then
10640         Disc = "Ext"
10650         TestColumn = "Analyte"
10660         TableName = "ExtResults"
10670         DateColumn = "RetDate"
10680         DisCode = "X"
10690         DispName = "External"
10700     ElseIf i = 5 Then
10710         Disc = "End"
10720         TestColumn = "ShortName"
10730         TableName = "EndResults"
10740         DisCode = "E"
10750         DispName = "Endocrinology"
10760     ElseIf i = 6 Then
10770         Disc = "Imm"
10780         TestColumn = "ShortName"
10790         TableName = "ImmResults"
10800         DisCode = "I"
10810         DispName = "Immunology"
10820     ElseIf i = 7 Then
10830         Disc = "Bga"
10840         TestColumn = "ShortName"
10850         TableName = "BgaResults"
10860         DisCode = "B"
10870         DispName = "Blood Gas"
10880     ElseIf i = 8 Then
10890         Disc = "Cyto"   ' NOT EXISTS IN DATABASE
10900         TestColumn = ""
10910         DispName = "Cytology"
10920     ElseIf i = 9 Then
10930         Disc = "Histo"  ' NOT EXISTS IN DATABASE
10940         TestColumn = ""
10950         DispName = "Histology"
10960     ElseIf i = 10 Then
10970         TableName = "PrintValidLog"
10980         DispName = ""
10990     ElseIf i = 11 Then
11000         Disc = "Semen"
11010         TestColumn = ""
11020         ResultColumn = "Motility"
11030         TableName = "SemenResults"
11040         DateColumn = "DateTimeOfRecord"
11050         DispName = "Semen"
11060     End If


11070     If optNotValidated Then
11080         Selection = "AND R.Valid = 0"
11090     ElseIf optNotPrinted Then
11100         Selection = "AND ISNULL(R.printed,0) = 0"
11110     ElseIf optValidNotPrint Then
11120         Selection = "AND R.Valid = 1 AND ISNULL(R.printed,0) = 0"
11130     ElseIf optNotValidatedPrinted Then
11140         Selection = "AND R.Valid = 0 AND ISNULL(R.printed,0) = 0"
11150     End If

11160     DeptClms = ",'" & DisCode & "' AS DisCode,'" & DispName & "' AS DispName"

11170     If i <> 4 And i <> 7 And i <> 11 And i <> 8 And i <> 9 And i <> 10 Then
11180         If i = 1 Or i = 3 Or i = 5 Or i = 6 Or i = 7 Or i = 4 Or i = 11 Then
11190             sql = "SELECT DISTINCT R.SampleID, D.DateTimeDemographics as EntryTime, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician " & DeptClms & " FROM " & TableName & " R " & _
                        "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                        "INNER JOIN " & Disc & "TestDefinitions T ON R.Code = T.Code " & _
                        "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & IIf(i = 1, Format$(dtTo.Value + 1, "dd/mmm/yyyy"), Format$(dtTo.Value, "dd/mmm/yyyy")) & "' " & _
                        "AND T.Printable = '1' " & _
                        Selection & " " & _
                        "AND R.SampleID NOT IN (SELECT SampleID FROM PrintPending)"
11200         ElseIf i = 2 Then
11210             sql = "SELECT DISTINCT R.SampleID, D.DateTimeDemographics as EntryTime, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician " & DeptClms & " FROM " & TableName & " R " & _
                        "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                        "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy") & "' " & _
                        Selection & " " & _
                        "AND R.SampleID NOT IN (SELECT SampleID FROM PrintPending)"

11220         End If

11230         sqlUnion = sqlUnion & " UNION ALL " & sql


11240     End If
11250 Next i
11260 SortBy = ""
11270 If cmbSortBy.Text = "Earliest Samples First" Then
11280     SortBy = " ORDER BY EntryTime asc"    '& Replace(cmbSortBy, "-", ",D.")
11290 Else
11300     SortBy = " ORDER BY D." & Replace(cmbSortBy, "-", ",D.")
11310 End If

11320 sqlUnion = sqlUnion & " " & SortBy

11330 sqlUnion = Replace(sqlUnion, "Dummy  UNION ALL ", "")

11340 Set tb = New Recordset
11350 RecOpenServer 0, tb, sqlUnion
      '          If tb.State = 0 Then Exit Sub
11360 While Not tb.EOF
          '-----
11370     If tb!DispName = "Biochemistry" Then
11380         TableName = "BioResults"
11390     ElseIf tb!DispName = "Haematology" Then
11400         TableName = "HaemResults"
11410     ElseIf tb!DispName = "Coagulation" Then
11420         TableName = "CoagResults"
11430     End If
11440     If checkAllResultsValid(tb!SampleID, TableName) = True Then
              '=====
              '
11450         If optHisto Then
11460             s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (sysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H"
11470         ElseIf optCyto Then
11480             s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (sysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "C"
11490         ElseIf optMicro Then
11500             s = Val(tb!SampleID) ' - sysOptMicroOffset(0)
11510         ElseIf optSemen Then
11520             s = Val(tb!SampleID) - sysOptSemenOffset(0)
11530         Else
11540             s = tb!SampleID
11550         End If

11560         s = s & vbTab & _
                  tb!PatName & vbTab & _
                  tb!Chart & vbTab & _
                  tb!GP & vbTab & _
                  tb!Ward & vbTab & _
                  tb!Clinician & "" & vbTab & _
                  tb!DisCode & vbTab & _
                  tb!DispName

11570         g.AddItem s
11580     End If
11590     tb.MoveNext

11600 Wend

11610 FixG g

11620 If g.Rows > 2 Then
11630     lblTotal = "Total samples : " & g.Rows - 1
11640 Else
11650     lblTotal = ""
11660 End If


11670 Exit Sub


FillGAll_Error:

      Dim strES As String
      Dim intEL As Integer

11680 intEL = Erl
11690 strES = Err.Description
11700 LogError "frmNotValidatedPrinted", "FillGAll", intEL, strES, sql
End Sub

Private Sub bcancel_Click()

11710 On Error GoTo bCancel_Click_Error

11720 Unload Me

11730 Exit Sub

bCancel_Click_Error:
      Dim strES As String
      Dim intEL As Integer

11740 intEL = Erl
11750 strES = Err.Description
11760 LogError "frmNotValidatedPrinted", "bcancel_Click", intEL, strES

End Sub

Private Sub bPrint_Click()

11770 On Error GoTo bPrint_Click_Error

      Dim Y As Long
      Dim X As Long
      Dim sql As String
      Dim sn As New Recordset


11780 If UserHasAuthority(UserMemberOf, "MainBatchPrinting") = False Then
11790     iMsg "You do not have authority for batch Printing " & vbCrLf & "Please contact system administrator"
11800     Exit Sub
11810 End If

11820 Printer.Orientation = vbPRORLandscape
11830 Printer.Font.Name = "Courier New"
11840 PrintText FormatString("Unvalidated / Not Printed Samples List", 99, , AlignCenter), 10, True, , , , True
11850 PrintText FormatString("From " & Format(dtFrom, "dd/mmm/yyyy") & " to " & Format(dtTo, "dd/mmm/yyyy"), 99, , AlignCenter), 10, True, , , , True
11860 PrintText String(107, "-"), , , , , , True



11870 For Y = 0 To g.Rows - 1


11880     PrintText FormatString(g.TextMatrix(Y, 0), 10, "|"), 9, IIf(Y = 0, True, False)   'sample id
11890     PrintText FormatString(g.TextMatrix(Y, 1), 30, "|"), 9, IIf(Y = 0, True, False)     'patient name
11900     PrintText FormatString(g.TextMatrix(Y, 2), 10, "|"), 9, IIf(Y = 0, True, False)     'test name
11910     PrintText FormatString(g.TextMatrix(Y, 3), 20, "|"), 9, IIf(Y = 0, True, False)      'gp
11920     PrintText FormatString(g.TextMatrix(Y, 4), 20, "|"), 9, IIf(Y = 0, True, False)  'ward
11930     PrintText FormatString(g.TextMatrix(Y, 5), 20), 9, IIf(Y = 0, True, False), , , , True     'result
          'PrintText FormatString(g.TextMatrix(Y, 6), 10), 9, IIf(Y = 0, True, False), , , , True    'return date

11940     If Y = 0 Then PrintText String(107, "-"), , , , , , True
11950 Next

11960 Printer.EndDoc

11970 Exit Sub

bPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

11980 intEL = Erl
11990 strES = Err.Description
12000 LogError "frmNotValidatedPrinted", "bprint_Click", intEL, strES

End Sub


Private Sub cmdRefresh_Click()
12010 FillG

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdPrintSamples_Click
' Author    : XPMUser
' Date      : 22/Oct/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function SavePrintInhibitNotValidated(ByVal SampleID As String, ByVal Dept As String) As Boolean
      'Returns True if there is something to print

      Dim sql As String
      Dim Y As Integer
      Dim Discipline As String
      Dim g As MSFlexGrid

12020 On Error GoTo SavePrintInhibitNotValidated_Error

12030 Discipline = Dept


12040 sql = "DELETE FROM PrintInhibit WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Discipline = '" & Discipline & "' " & _
            "INSERT INTO PrintInhibit " & _
            "(SampleID, Discipline, Parameter) " & _
            "SELECT DISTINCT '" & SampleID & "', '" & Discipline & "', D.ShortName " & _
            "FROM " & Discipline & "TestDefinitions D " & _
            "JOIN " & Discipline & "Results R " & _
            "ON D.Code = R.Code " & _
            "WHERE R.SampleID = '" & SampleID & "' AND R.Valid <> 1"
12050 Cnxn(0).Execute sql


12060 SavePrintInhibitNotValidated = True

12070 Exit Function

SavePrintInhibitNotValidated_Error:

      Dim strES As String
      Dim intEL As Integer

12080 intEL = Erl
12090 strES = Err.Description
12100 LogError "frmNotValidatedNotPrinted", "SavePrintInhibitNotValidated", intEL, strES, sql

End Function

Private Sub cmdPrintSamples_Click()
12110 On Error GoTo cmdPrintSamples_Click_Error
      Dim DeptSel As String
      Dim i As Integer

12120 With g
12130     For i = 1 To .Rows - 1
12140         If .TextMatrix(i, 6) <> "" Then
                            
12150             Call PrintThis(.TextMatrix(i, 0), .TextMatrix(i, 1), "SEX", .TextMatrix(i, 4), .TextMatrix(i, 3), .TextMatrix(i, 5), .TextMatrix(i, 6))
12160         End If
12170     Next i
12180     MsgBox .Rows - 1 & " Samples sent for printing"
12190 End With

12200 If optAllDisp Then
12210     FillGAll
12220 Else
12230     FillG
12240 End If

12250 Exit Sub

cmdPrintSamples_Click_Error:

      Dim strES As String
      Dim intEL As Integer

12260 intEL = Erl
12270 strES = Err.Description
12280 LogError "frmNotValidatedPrinted", "cmdPrintSamples_Click", intEL, strES
End Sub
Private Sub PrintThis(pSampleID As String, pPatName As String, pSex As String, pWard As String, pGP As String, pClinican As String, pDept As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Discipline As String

12290 On Error GoTo PrintThis_Error

12300 Select Case pDept
          Case "B": Discipline = "Bio"
12310     Case "C": Discipline = "Coag"
12320     Case Else: Discipline = ""
12330 End Select

12340 SavePrintInhibitNotValidated pSampleID, Discipline

12350 sql = "Select * from PrintPending where " & _
            "Department = '" & pDept & "' " & _
            "and SampleID = '" & pSampleID & "'"
12360 Set tb = New Recordset
12370 RecOpenClient 0, tb, sql
12380 If tb.EOF Then
12390     tb.AddNew
12400 End If
12410 tb!SampleID = pSampleID
12420 tb!Ward = pWard
12430 tb!Clinician = pClinican
12440 tb!GP = pGP
12450 tb!Department = pDept
12460 tb!Initiator = GetValidatorUser(pSampleID, pDept) 'UserName
12470 tb!UsePrinter = "" 'pPrintToPrinter

12480 tb.Update
12490 Exit Sub

PrintThis_Error:

      Dim strES As String
      Dim intEL As Integer

12500 intEL = Erl
12510 strES = Err.Description
12520 LogError "frmEditMicrobiology", "PrintThis", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : XPMUser
' Date      : 11/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

12530 On Error GoTo Form_Load_Error
12540 dtFrom = Format$(Now - 7, "dd/mm/yyyy")
12550 dtTo = Format$(Now, "dd/mm/yyyy")
      'optEnd.Enabled = SysOptDeptEnd(0)
      'optImm.Enabled = SysOptDeptImm(0)
      'optBG.Enabled = SysOptDeptBga(0)
      'optMicro.Enabled = False 'SysOptDeptMicro(0)
      'optSemen.Enabled = SysOptDeptSemen(0)
      'optHisto.Enabled = SysOptDeptHisto(0)
      'optCyto.Enabled = SysOptDeptCyto(0)
12560 SetFormCaption
12570 cmbSortBy.Clear
12580 cmbSortBy.AddItem ("Ward-GP-Clinician")
12590 cmbSortBy.AddItem ("Earliest Samples First")
12600 cmbSortBy.AddItem ("GP-Ward-Clinician")
12610 cmbSortBy.AddItem ("GP-Clinician-Ward")
12620 cmbSortBy.AddItem ("Ward-Clinician-GP")
12630 cmbSortBy.AddItem ("Clinician-Ward-GP")
12640 cmbSortBy.AddItem ("Clinician-GP-Ward")
12650 cmbSortBy.ListIndex = 0
12660 PrintSamplesVisibilty (0)
12670 g.ColWidth(6) = 0
12680 g.ColWidth(7) = 1100
12690 optAllDisp.Value = True
12700 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

12710 intEL = Erl
12720 strES = Err.Description
12730 LogError "frmNotValidatedPrinted", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
          
12740 On Error GoTo g_Click_Error
12750 If g.MouseRow = 0 Then
12760     If SortOrder Then
12770         g.Sort = flexSortGenericAscending
12780     Else
12790         g.Sort = flexSortGenericDescending
12800     End If
12810     SortOrder = Not SortOrder
12820 Else
12830     If g.MouseCol = 0 Then
12840         If UserName = "" Then
12850             iMsg "Please logon to system to view sample"
12860         Else

12870             With frmEditAll
12880                 .txtSampleID = g.TextMatrix(g.MouseRow, g.MouseCol)
12890                 Unload Me
                      
12900                 .Show 1
12910             End With
12920         End If
12930     End If
12940 End If

12950 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

12960 intEL = Erl
12970 strES = Err.Description
12980 LogError "frmDaily", "g_Click", intEL, strES

End Sub

Private Sub optAllDisp_Click()
12990 On Error GoTo optAllDisp_Click_Error

13000 SetFormCaption
13010 FillGAll

13020 Exit Sub

optAllDisp_Click_Error:
      Dim strES As String
      Dim intEL As Integer

13030 intEL = Erl
13040 strES = Err.Description
13050 LogError "frmNotValidatedPrinted", "optAllDisp_Click", intEL, strES

End Sub

Private Sub optBG_Click()
13060 On Error GoTo optBG_Click_Error
13070 SetFormCaption
13080 FillG
13090 Exit Sub

optBG_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13100 intEL = Erl
13110 strES = Err.Description
13120 LogError "frmDaily", "optBG_Click", intEL, strES

End Sub

Private Sub optBio_Click()
13130 On Error GoTo optBio_Click_Error

13140 SetFormCaption
13150 FillG

13160 Exit Sub

optBio_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13170 intEL = Erl
13180 strES = Err.Description
13190 LogError "frmDaily", "optBio_Click", intEL, strES

End Sub

Private Sub optCoag_Click()
13200 On Error GoTo optCoag_Click_Error

13210 SetFormCaption
13220 FillG


13230 Exit Sub

optCoag_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13240 intEL = Erl
13250 strES = Err.Description
13260 LogError "frmDaily", "optCoag_Click", intEL, strES

End Sub

Private Sub optCyto_Click()
13270 On Error GoTo optCyto_Click_Error

13280 SetFormCaption
13290 FillG

13300 Exit Sub

optCyto_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13310 intEL = Erl
13320 strES = Err.Description
13330 LogError "frmDaily", "optCyto_Click", intEL, strES

End Sub

Private Sub optEnd_Click()
13340 On Error GoTo optEnd_Click_Error

13350 SetFormCaption
13360 FillG


13370 Exit Sub

optEnd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13380 intEL = Erl
13390 strES = Err.Description
13400 LogError "frmDaily", "optEnd_Click", intEL, strES

End Sub

Private Sub optExt_Click()

13410 On Error GoTo optExt_Click_Error

13420 SetFormCaption
13430 FillG


13440 Exit Sub

optExt_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13450 intEL = Erl
13460 strES = Err.Description
13470 LogError "frmDaily", "optExt_Click", intEL, strES

End Sub

Private Sub optHaem_Click()
13480 On Error GoTo optHaem_Click_Error

13490 SetFormCaption
13500 FillG


13510 Exit Sub

optHaem_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13520 intEL = Erl
13530 strES = Err.Description
13540 LogError "frmDaily", "optHaem_Click", intEL, strES

End Sub

Private Sub optHisto_Click()
13550 On Error GoTo optHisto_Click_Error

13560 SetFormCaption
13570 FillG


13580 Exit Sub

optHisto_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13590 intEL = Erl
13600 strES = Err.Description
13610 LogError "frmDaily", "optHisto_Click", intEL, strES

End Sub

Private Sub optImm_Click()
13620 On Error GoTo optImm_Click_Error

13630 SetFormCaption
13640 FillG


13650 Exit Sub

optImm_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13660 intEL = Erl
13670 strES = Err.Description
13680 LogError "frmDaily", "optImm_Click", intEL, strES

End Sub

Private Sub optMicro_Click()
13690 On Error GoTo optMicro_Click_Error

13700 SetFormCaption
13710 FillG

13720 Exit Sub

optMicro_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13730 intEL = Erl
13740 strES = Err.Description
13750 LogError "frmDaily", "optMicro_Click", intEL, strES

End Sub

Private Sub optNotPrinted_Click()

13760 On Error GoTo optNotPrinted_Click_Error

13770 If optExt Then
13780     iMsg "Only unvalidated samples can be searched for externals"
13790     optNotValidated.Value = True
13800 End If
13810     PrintSamplesVisibilty (1)
13820 Exit Sub

optNotPrinted_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13830 intEL = Erl
13840 strES = Err.Description
13850 LogError "frmNotValidatedPrinted", "optNotPrinted_Click", intEL, strES

End Sub

Private Sub optNotValidated_Click()
13860 On Error GoTo optNotValidated_Click_Error

13870 PrintSamplesVisibilty (0)

13880 Exit Sub

optNotValidated_Click_Error:
      Dim strES As String
      Dim intEL As Integer

13890 intEL = Erl
13900 strES = Err.Description
13910 LogError "frmNotValidatedPrinted", "optNotValidated_Click", intEL, strES

End Sub

Private Sub optNotValidatedPrinted_Click()

13920 On Error GoTo optNotValidatedPrinted_Click_Error
13930 If optExt Then
13940     iMsg "Only unvalidated samples can be searched for externals"
13950     optNotValidated.Value = True
13960 End If

13970 PrintSamplesVisibilty (0)
13980 Exit Sub

optNotValidatedPrinted_Click_Error:

      Dim strES As String
      Dim intEL As Integer

13990 intEL = Erl
14000 strES = Err.Description
14010 LogError "frmNotValidatedPrinted", "optNotValidatedPrinted_Click", intEL, strES

End Sub

Private Sub optSemen_Click()
14020 On Error GoTo optSemen_Click_Error

14030 SetFormCaption
14040 FillG


14050 Exit Sub

optSemen_Click_Error:

      Dim strES As String
      Dim intEL As Integer

14060 intEL = Erl
14070 strES = Err.Description
14080 LogError "frmDaily", "optSemen_Click", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintSamplesVisibilty
' Author    : XPMUser
' Date      : 22/Oct/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintSamplesVisibilty(ShowButton As Integer)
14090 On Error GoTo PrintSamplesVisibilty_Error


14100 cmdPrintSamples.Visible = False
14110 If ShowButton = 1 Then
14120     If UserName <> "" Then
14130         cmdPrintSamples.Visible = True
14140     End If
14150 End If


14160 Exit Sub


PrintSamplesVisibilty_Error:

      Dim strES As String
      Dim intEL As Integer

14170 intEL = Erl
14180 strES = Err.Description
14190 LogError "frmNotValidatedPrinted", "PrintSamplesVisibilty", intEL, strES

End Sub

Private Sub optValidNotPrint_Click()
14200 On Error GoTo optValidNotPrint_Click_Error

14210 PrintSamplesVisibilty (1)

14220 Exit Sub

optValidNotPrint_Click_Error:
      Dim strES As String
      Dim intEL As Integer

14230 intEL = Erl
14240 strES = Err.Description
14250 LogError "frmNotValidatedPrinted", "optValidNotPrint_Click", intEL, strES

End Sub
Private Sub cmdXL_Click()

14260 On Error GoTo cmdXL_Click_Error

14270 ExportFlexGrid g, Me

14280 Exit Sub

cmdXL_Click_Error:
      Dim strES As String
      Dim intEL As Integer

14290 intEL = Erl
14300 strES = Err.Description
14310 LogError "frmNotValidatedPrinted", "cmdXL_Click", intEL, strES

End Sub
