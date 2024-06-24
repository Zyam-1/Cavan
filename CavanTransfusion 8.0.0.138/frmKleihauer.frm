VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmKleihauer 
   Caption         =   "NetAcquire - Transfusion - Kleihauer"
   ClientHeight    =   6435
   ClientLeft      =   405
   ClientTop       =   675
   ClientWidth     =   10770
   Icon            =   "frmKleihauer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10770
   Begin VB.Frame Frame3 
      Caption         =   "Negative Control"
      Height          =   1215
      Left            =   8670
      TabIndex        =   56
      Top             =   3750
      Width           =   1725
      Begin VB.OptionButton optNegPass 
         Alignment       =   1  'Right Justify
         Caption         =   "Pass"
         Height          =   255
         Left            =   720
         TabIndex        =   59
         Top             =   300
         Width           =   645
      End
      Begin VB.OptionButton optNegFail 
         Alignment       =   1  'Right Justify
         Caption         =   "Fail"
         Height          =   255
         Left            =   810
         TabIndex        =   58
         Top             =   570
         Width           =   555
      End
      Begin VB.OptionButton optNegNotChecked 
         Alignment       =   1  'Right Justify
         Caption         =   "Not Checked"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Positive Control"
      Height          =   1215
      Left            =   6720
      TabIndex        =   52
      Top             =   3750
      Width           =   1725
      Begin VB.OptionButton optPosNotChecked 
         Alignment       =   1  'Right Justify
         Caption         =   "Not Checked"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optPosFail 
         Alignment       =   1  'Right Justify
         Caption         =   "Fail"
         Height          =   255
         Left            =   810
         TabIndex        =   54
         Top             =   570
         Width           =   555
      End
      Begin VB.OptionButton optPosPass 
         Alignment       =   1  'Right Justify
         Caption         =   "Pass"
         Height          =   255
         Left            =   720
         TabIndex        =   53
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.TextBox txtChart 
      Height          =   285
      Left            =   4950
      TabIndex        =   49
      Top             =   330
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   765
      Left            =   6780
      Picture         =   "frmKleihauer.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5340
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   765
      Left            =   8010
      Picture         =   "frmKleihauer.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5355
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   9270
      Picture         =   "frmKleihauer.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5340
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   6720
      TabIndex        =   23
      Top             =   1800
      Width           =   3675
      Begin ComCtl2.UpDown udFetal 
         Height          =   225
         Left            =   2730
         TabIndex        =   47
         Top             =   330
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         _Version        =   327681
         BuddyControl    =   "txtFetal"
         BuddyDispid     =   196622
         OrigLeft        =   3060
         OrigTop         =   270
         OrigRight       =   3405
         OrigBottom      =   510
         Max             =   200
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtFetal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1770
         TabIndex        =   25
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblFMHReport 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estimated FMH is 999 mls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   1185
         Width           =   3135
      End
      Begin VB.Label lblRh 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1770
         TabIndex        =   27
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Patients Rh status"
         Height          =   195
         Left            =   420
         TabIndex        =   26
         Top             =   810
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Number of Fetal Cells"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.TextBox txtSampleID 
      Height          =   345
      Left            =   1410
      TabIndex        =   1
      Top             =   300
      Width           =   1755
   End
   Begin ComCtl2.UpDown udLabNum 
      Height          =   255
      Left            =   3210
      TabIndex        =   2
      Top             =   360
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtSampleID"
      BuddyDispid     =   196627
      OrigLeft        =   3480
      OrigTop         =   360
      OrigRight       =   4470
      OrigBottom      =   615
      Max             =   99999
      Min             =   1
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   1020
      TabIndex        =   60
      Top             =   6210
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblReceivedDate 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8790
      TabIndex        =   61
      Top             =   615
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblSampleDate 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8790
      TabIndex        =   51
      Top             =   315
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sample Date"
      Height          =   195
      Left            =   8820
      TabIndex        =   50
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblSampleComment 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   46
      Top             =   4680
      Width           =   4785
   End
   Begin VB.Label lblComment 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   45
      Top             =   2580
      Width           =   4785
   End
   Begin VB.Label lblAddr 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1410
      TabIndex        =   44
      Top             =   2280
      Width           =   4785
   End
   Begin VB.Label lblAddr 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1410
      TabIndex        =   43
      Top             =   1980
      Width           =   4785
   End
   Begin VB.Label lblAddr 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   42
      Top             =   1680
      Width           =   4785
   End
   Begin VB.Label lblAddr 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1410
      TabIndex        =   41
      Top             =   1380
      Width           =   4785
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   40
      Top             =   1080
      Width           =   4785
   End
   Begin VB.Label lblAge 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8010
      TabIndex        =   39
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label lblDoB 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6720
      TabIndex        =   38
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label lblTypenex 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6720
      TabIndex        =   37
      Top             =   330
      Width           =   1155
   End
   Begin VB.Label lblSpecial 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   36
      Top             =   4380
      Width           =   4785
   End
   Begin VB.Label lblProcedure 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   35
      Top             =   4080
      Width           =   4785
   End
   Begin VB.Label lblConditions 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   34
      Top             =   3780
      Width           =   4785
   End
   Begin VB.Label lblGP 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   33
      Top             =   3480
      Width           =   4785
   End
   Begin VB.Label lblClinician 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   32
      Top             =   3180
      Width           =   4785
   End
   Begin VB.Label lblWard 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1410
      TabIndex        =   31
      Top             =   2880
      Width           =   4785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Comment"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   4710
      Width           =   1230
   End
   Begin VB.Label lblKnownAntibody 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1410
      TabIndex        =   21
      Top             =   780
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   5040
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   930
      TabIndex        =   19
      Top             =   1140
      Width           =   420
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Ward"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   33
      Left            =   960
      TabIndex        =   18
      Top             =   2940
      Width           =   390
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Clin"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   31
      Left            =   1095
      TabIndex        =   17
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Cond"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   975
      TabIndex        =   16
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Proc"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   29
      Left            =   1020
      TabIndex        =   15
      Top             =   4140
      Width           =   330
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Addr 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   28
      Left            =   885
      TabIndex        =   14
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Spec"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   27
      Left            =   975
      TabIndex        =   13
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Remark"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   26
      Left            =   795
      TabIndex        =   12
      Top             =   2610
      Width           =   555
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   1260
      TabIndex        =   11
      Top             =   1740
      Width           =   90
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "3"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   1260
      TabIndex        =   10
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "4"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   1260
      TabIndex        =   9
      Top             =   2340
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "GP"
      Height          =   195
      Left            =   1125
      TabIndex        =   8
      Top             =   3540
      Width           =   225
   End
   Begin VB.Label Label13 
      Caption         =   "TYPENEX"
      Height          =   195
      Left            =   6750
      TabIndex        =   7
      Top             =   120
      Width           =   765
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "D.o.B."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   6750
      TabIndex        =   6
      Top             =   960
      Width           =   450
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Age"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   8010
      TabIndex        =   5
      Top             =   960
      Width           =   285
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   9270
      TabIndex        =   4
      Top             =   960
      Width           =   270
   End
   Begin VB.Label lblSex 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9240
      TabIndex        =   3
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmKleihauer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CalculateFMH()

      Dim sngFMH As Single

10    lblFMHReport.Caption = ""

20    If lblRh.Caption <> "" Then
30      If Trim$(txtFetal) <> "" Then
40        sngFMH = Val(txtFetal) * 0.4
50        If lblRh.Caption = "Positive" Then
60          lblFMHReport = "Estimated FMH is " & Format$(sngFMH, "##0.0") & " mls"
70        ElseIf lblRh.Caption = "Negative" Then
'80          If Val(txtFetal) < 30 Then
'90            lblFMHReport = "Estimated FMH is < 12 mls"
'100         ElseIf Val(txtFetal) = 30 Then
'110           lblFMHReport = "Estimated FMH is = 12 mls"
'120         Else
80            lblFMHReport = "Estimated FMH is " & Format$(sngFMH, "##0.0") & " mls"
'140         End If
90        End If
100     End If
110   End If

End Sub

Private Sub ClearDetails()

10    lblKnownAntibody.Visible = False
20    txtChart.Text = ""
30    lblKnownAntibody.Caption = ""
40    lblTypenex.Caption = ""
50    lblName.Caption = ""
60    lblSex.Caption = ""
70    lblAddr(0).Caption = ""
80    lblAddr(1).Caption = ""
90    lblAddr(2).Caption = ""
100   lblAddr(3).Caption = ""
110   lblDoB.Caption = ""
120   lblAge.Caption = ""
130   lblWard.Caption = ""
140   lblGP.Caption = ""
150   lblClinician.Caption = ""
160   lblConditions.Caption = ""
170   lblProcedure.Caption = ""
180   lblSpecial.Caption = ""
190   lblComment.Caption = ""
200   lblSampleComment.Caption = ""
210   txtFetal.Text = ""
220   lblRh.Caption = ""
230   lblFMHReport.Caption = ""

End Sub

Private Sub Save()
  
      Dim sql As String
      Dim pos As String
      Dim Neg As String

10    On Error GoTo Save_Error

20    sql = "IF NOT EXISTS " & _
      "   (SELECT * from Kleihauer where Sampleid = '" & txtSampleID.Text & "')" & _
      "   INSERT INTO Kleihauer " & _
            "( SampleID, Chart, FetalCells, Rh, Report, Operator, DateTime ) VALUES " & _
            "('" & txtSampleID.Text & "', " & _
            " '" & txtChart.Text & "', " & _
            " '" & Val(txtFetal.Text) & "', " & _
            " '" & lblRh.Caption & "', " & _
            " '" & lblFMHReport.Caption & "', " & _
            " '" & UserName & "', " & _
            " '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' )" & _
      "ELSE " & _
      "UPDATE Kleihauer SET Chart = '" & txtChart.Text & "' , FetalCells = '" & Val(txtFetal.Text) & "' , Rh = '" & lblRh.Caption & "', Report = '" & lblFMHReport.Caption & "' ," & _
      "Operator = '" & UserName & "' , DateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' where Sampleid = '" & txtSampleID.Text & "'"
      
30    CnxnBB(0).Execute sql

40    pos = Switch(optPosPass, "P", optPosFail, "F", optPosNotChecked, "N")
50    Neg = Switch(optNegPass, "P", optNegFail, "F", optNegNotChecked, "N")

60    sql = "INSERT INTO KleihauerQC " & _
              "( SampleID, Rhesus, Positive, Negative, Operator, DateTime ) VALUES " & _
              "('" & txtSampleID.Text & "', " & _
              " '" & lblRh & "', " & _
              " '" & pos & "', " & _
              " '" & Neg & "', " & _
              " '" & UserName & "', " & _
              " '" & Format(Now, "yyyyMMdd HH:mm:ss") & "' )"
70    CnxnBB(0).Execute sql

80    iMsg "Kleihauer saved!"

90    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmKleihauer", "Save", intEL, strES, sql
  
End Sub

Private Sub LoadKleihauer()
  
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo LoadKleihauer_Error

20    sql = "Select * from Kleihauer where " & _
            "SampleID = '" & txtSampleID & "' " & _
            "order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      txtFetal.Text = CStr(tb!FetalCells)
70      lblRh.Caption = tb!Rh & ""
80      lblFMHReport.Caption = tb!Report & ""
90    End If

100   Exit Sub

LoadKleihauer_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmKleihauer", "LoadKleihauer", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPrint_Click()

10    On Error GoTo cmdPrint_Click_Error

20    If Trim$(lblName.Caption) = "" Or lblFMHReport.Caption = "" Then
30      iMsg "Nothing to Print!", vbExclamation
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60    End If

70    If Trim$(txtChart.Text) = "" Then
80      iMsg "Must Have Chart Number!", vbExclamation
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   If Not QueryRhOK() Then
130     iMsg "Rh Status does not match historical record!", vbCritical
140     If TimedOut Then Unload Me: Exit Sub
150     Exit Sub
160   End If
 
170   Save
180   PrintKleihauerFormCavan txtSampleID, txtFetal, getFoetalCellWording(txtFetal), lblSampleDate, lblReceivedDate

190   Exit Sub

cmdPrint_Click_Error:

 Dim strES As String
 Dim intEL As Integer

200    intEL = Erl
210    strES = Err.Description
220    LogError "frmKleihauer", "cmdPrint_Click", intEL, strES

End Sub

Private Function getFoetalCellWording(ByVal intF As Integer) As String

10    If intF = "0" Then
20        getFoetalCellWording = "M1"
30    ElseIf intF >= "1" And intF <= "4" Then
40        getFoetalCellWording = "M2"
50    ElseIf intF >= "5" And intF <= "19" Then
60        getFoetalCellWording = "M3"
70    ElseIf intF >= "20" And intF <= "30" Then
80        getFoetalCellWording = "M4"
90    ElseIf intF >= "31" Then
100       getFoetalCellWording = "M5"
110   End If

End Function

Private Function QueryRhOK() As Boolean

      Dim sql As String
      Dim tb As Recordset
      Dim Rh As String

10    On Error GoTo QueryRhOK_Error

20    QueryRhOK = False

30    sql = "select fGroup from PatientDetails where " & _
            "PatNum = '" & txtChart.Text & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If tb.EOF Then
70      QueryRhOK = True
80    Else
90      Rh = ""
100     Do While Not tb.EOF
110       If Trim$(tb!fGroup & "") <> "" Then
120         If InStr(tb!fGroup, "Pos") <> 0 Then
130           Rh = "Positive"
140         ElseIf InStr(tb!fGroup, "Neg") <> 0 Then
150           Rh = "Negative"
160         End If
170       End If
180       tb.MoveNext
190     Loop
  
200     If lblRh.Caption = Rh Then
210       QueryRhOK = True
220     End If

230   End If

240   Exit Function

QueryRhOK_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmKleihauer", "QueryRhOK", intEL, strES, sql


End Function

Public Sub CheckKleihauerInDb(ByVal Cx As Connection)

      Dim sql As String
      Dim tbExists As Recordset

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist "
      'if it has a record then the table does exist.

10    On Error GoTo CheckKleihauerInDb_Error

20    sql = "SELECT Name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = 'Kleihauer'"
30    Set tbExists = New Recordset
40    Set tbExists = Cx.Execute(sql)

50    If tbExists.EOF Then 'There is no table  in database
60      sql = "CREATE TABLE Kleihauer " & _
              "( SampleID  nvarchar(50) collate SQL_Latin1_General_CP1_CI_AS, " & _
              "  Chart  nvarchar(50) collate SQL_Latin1_General_CP1_CI_AS, " & _
              "  FetalCells  int, " & _
              "  Rh nvarchar(10), " & _
              "  Report nvarchar(50), " & _
              "  Operator nvarchar(50), " & _
              "  DateTime  datetime )"
70      Cx.Execute sql
80    End If

90    Exit Sub

CheckKleihauerInDb_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmKleihauer", "CheckKleihauerInDb", intEL, strES, sql

End Sub
Private Sub cmdSave_Click()

10    If Trim$(lblName.Caption) = "" Or lblFMHReport.Caption = "" Then
20      iMsg "Nothing to Save!", vbExclamation
30      If TimedOut Then Unload Me: Exit Sub
40      Exit Sub
50    End If

60    If Trim$(txtChart.Text) = "" Then
70      iMsg "Must Have Chart Number!", vbExclamation
80      If TimedOut Then Unload Me: Exit Sub
90      Exit Sub
100   End If

      'If Not QueryRhOK() Then
      '  iMsg "Rh Status does not match historical record!", vbCritical
      '  If TimedOut Then Unload Me: Exit Sub
      '  Exit Sub
      'End If
  
110   Save

End Sub

Private Sub Form_Load()

10    lblFMHReport.Caption = ""

20    CheckKleihauerInDb CnxnBB(0)

End Sub


Private Sub LoadLabNumber()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadLabNumber_Error

  
20    ClearDetails

30    sql = "select * from patientdetails where " & _
            "labnumber = '" & txtSampleID & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70      lblSampleDate = Format$(tb!SampleDate, "dd/MM/yyyy hh:mm")
80      lblReceivedDate = Format$(tb!DateReceived, "dd/MM/yyyy hh:mm")
90      txtChart.Text = tb!Patnum & ""
100     If Trim$(txtChart.Text) <> "" Then
110       lblKnownAntibody.Caption = CheckPreviousABScreen(txtChart.Text)
120       If lblKnownAntibody.Caption <> "" Then
130         lblKnownAntibody.Visible = True
140       End If
150     End If
160     lblTypenex = tb!Typenex & ""
170     lblName = tb!Name & ""
180     Select Case Left$(tb!Sex & "", 1)
          Case "M": lblSex = "Male"
190       Case "F": lblSex = "Female"
200       Case "U": lblSex = "Unknown"
210       Case Else: lblSex = ""
220     End Select
230     lblAddr(0) = tb!Addr1 & ""
240     lblAddr(1) = tb!Addr2 & ""
250     lblAddr(2) = tb!Addr3 & ""
260     lblAddr(3) = tb!addr4 & ""
270     If Not IsNull(tb!DoB) Then
280       lblDoB = Format(tb!DoB, "dd/mm/yyyy")
290     Else
300       lblDoB = ""
310     End If
320     If Trim$(tb!Age & "") <> "" Then
330       lblAge = tb!Age
340     Else
350       lblAge = CalcAge(lblDoB)
360     End If
370     lblWard = tb!Ward & ""
380     lblGP = tb!GP & ""
390     lblClinician = tb!Clinician & ""
400     lblConditions = tb!Conditions & ""
410     lblProcedure = tb!Procedure & ""
420     lblSpecial = tb!specialprod & ""
430     lblComment = StripComment(tb!Comment & "")
440     lblSampleComment = tb!SampleComment & ""
  
450   End If

460   LoadKleihauer

470   Exit Sub

LoadLabNumber_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmKleihauer", "LoadLabNumber", intEL, strES, sql

End Sub

Private Sub LoadChart()

      Dim tb As Recordset
      Dim sql As String
      Dim strChart As String

10    On Error GoTo LoadChart_Error

20    strChart = txtChart.Text
30    ClearDetails
40    txtChart.Text = strChart

50    If txtChart.Text = "" Then Exit Sub

60    sql = "Select * from patientdetails where " & _
            "labnumber in ( " & _
            "  Select top 1 SampleID from Kleihauer where " & _
            "  Chart = '" & strChart & "' " & _
            "  order by datetime desc )"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql
90    If Not tb.EOF Then
  
100     txtSampleID.Text = tb!LabNumber & ""
110     lblKnownAntibody.Caption = CheckPreviousABScreen(txtChart.Text)
120     If lblKnownAntibody.Caption <> "" Then
130       lblKnownAntibody.Visible = True
140     End If
150     lblTypenex = tb!Typenex & ""
160     lblName = tb!Name & ""
170     Select Case Left$(tb!Sex & "", 1)
          Case "M": lblSex = "Male"
180       Case "F": lblSex = "Female"
190       Case "U": lblSex = "Unknown"
200       Case Else: lblSex = ""
210     End Select
220     lblAddr(0) = tb!Addr1 & ""
230     lblAddr(1) = tb!Addr2 & ""
240     lblAddr(2) = tb!Addr3 & ""
250     lblAddr(3) = tb!addr4 & ""
260     If Not IsNull(tb!DoB) Then
270       lblDoB = Format(tb!DoB, "dd/mm/yyyy")
280     Else
290       lblDoB = ""
300     End If
310     If Trim$(tb!Age & "") <> "" Then
320       lblAge = tb!Age
330     Else
340       lblAge = CalcAge(lblDoB)
350     End If
360     lblWard = tb!Ward & ""
370     lblGP = tb!GP & ""
380     lblClinician = tb!Clinician & ""
390     lblConditions = tb!Conditions & ""
400     lblProcedure = tb!Procedure & ""
410     lblSpecial = tb!specialprod & ""
420     lblComment = StripComment(tb!Comment & "")
430     lblSampleComment = tb!SampleComment & ""
  
440   End If

450   LoadKleihauer

460   Exit Sub

LoadChart_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmKleihauer", "LoadChart", intEL, strES, sql

End Sub


Private Sub lblRh_Click()

10    Select Case lblRh.Caption
        Case "": lblRh.Caption = "Negative"
20      Case "Negative": lblRh.Caption = "Positive"
30      Case Else: lblRh.Caption = ""
40    End Select

50    CalculateFMH

End Sub

Private Sub txtChart_LostFocus()

10    LoadChart

End Sub


Private Sub txtFetal_KeyPress(KeyAscii As Integer)
10    KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtFetal_LostFocus()

10    CalculateFMH

End Sub


Private Sub txtSampleID_LostFocus()

10    LoadLabNumber

End Sub


Private Sub udFetal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    CalculateFMH

End Sub


Private Sub udLabNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    LoadLabNumber

End Sub


