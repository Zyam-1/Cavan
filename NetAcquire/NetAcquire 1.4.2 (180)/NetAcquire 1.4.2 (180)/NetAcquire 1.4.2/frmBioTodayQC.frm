VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBioTodayQC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Validate QC"
   ClientHeight    =   7830
   ClientLeft      =   420
   ClientTop       =   525
   ClientWidth     =   9585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7830
   ScaleWidth      =   9585
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6720
      Picture         =   "frmBioTodayQC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   90
      Width           =   1155
   End
   Begin VB.ComboBox cmbControl 
      Height          =   315
      Left            =   3840
      TabIndex        =   6
      Text            =   "cmbControl"
      Top             =   420
      Width           =   2535
   End
   Begin VB.OptionButton optSampleType 
      Alignment       =   1  'Right Justify
      Caption         =   "Urine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2460
      TabIndex        =   4
      Top             =   390
      Width           =   675
   End
   Begin VB.OptionButton optSampleType 
      Alignment       =   1  'Right Justify
      Caption         =   "Serum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2370
      TabIndex        =   3
      Top             =   150
      Value           =   -1  'True
      Width           =   765
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   90
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   220856321
      CurrentDate     =   37082
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6885
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12144
      _Version        =   393216
      Cols            =   9
      FixedCols       =   2
      RowHeightMin    =   400
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      FormatString    =   $"frmBioTodayQC.frx":030A
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
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8220
      Picture         =   "frmBioTodayQC.frx":03B4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1155
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      Height          =   345
      Index           =   0
      Left            =   6360
      ScaleHeight     =   285
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   6390
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     Click on Parameter       to Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   525
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1845
   End
End
Attribute VB_Name = "frmBioTodayQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

11860     Unload Me

End Sub

Private Sub FillParameters()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim SampleType As String
          Dim intLine As Integer
          Dim LongName As String
          Dim Alias As String

11870     On Error GoTo FillParameters_Error

11880     If Trim$(cmbControl) = "" Then Exit Sub

11890     If g.TextMatrix(1, 0) <> "" Then
11900         For n = g.Rows - 1 To 1 Step -1
11910             Unload pic(n)
11920         Next
11930     End If

11940     g.Visible = False
11950     g.Rows = 2
11960     g.AddItem ""
11970     g.RemoveItem 1

11980     SampleType = IIf(optSampleType(0), "S", "U")
11990     sql = "Select AliasName from BioQCDefs where ControlName = '" & cmbControl & "'"
12000     Set tb = New Recordset
12010     RecOpenServer 0, tb, sql
12020     If tb.EOF Then Exit Sub
12030     Alias = tb!AliasName & ""

          'ColourBar pic(0), 0, 1, Result

12040     intLine = 1
12050     sql = "Select Q.code, Q.result, Q.runtime, D.LongName, B.* from " & _
              "BiochemistryQC as Q, Biotestdefinitions as D, BioQCDefs as B where " & _
              "Q.Code = D.Code " & _
              "and B.ParameterName = D.LongName " & _
              "and B.ControlName = '" & cmbControl & "' " & _
              "and Q.SampleType = '" & SampleType & "' " & _
              "and Q.RunDate = '" & Format(dt, "dd/mmm/yyyy") & "' " & _
              "and Q.AliasName = '" & Alias & "' " & _
              "order by D.printpriority, Q.RunTime"
12060     Set tb = New Recordset
12070     RecOpenClient 0, tb, sql
12080     Do While Not tb.EOF
        
12090         g.AddItem tb!LongName & vbTab & _
                  Format$(tb!RunTime, "hh:mm") & vbTab & _
                  Format$(tb!Result & "", "0.0#") & vbTab & _
                  vbTab & _
                  tb!mean & vbTab & _
                  tb!Low & vbTab & _
                  tb!High & vbTab & _
                  tb!SD & vbTab & _
                  tb!Code

12100         Load pic(intLine)
12110         pic(intLine).ScaleMode = 0
12120         pic(intLine).ScaleWidth = 1024
12130         ColourBar pic(intLine), Val(tb!Low), Val(tb!High), Val(tb!Result)
12140         g.Col = 3
12150         g.row = g.Rows - 1
12160         Set g.CellPicture = pic(intLine).Image
12170         g.CellPictureAlignment = flexAlignCenterCenter
12180         g.CellAlignment = flexAlignRightTop
12190         If Val(tb!Result) < Val(tb!Low) Or Val(tb!Result) > Val(tb!High) Then
12200             g.Col = 2
12210             g.CellBackColor = vbRed
12220         End If
12230         If tb!LongName = g.TextMatrix(g.row - 1, 0) Then
12240             g.Col = 0
12250             g.CellBackColor = vbMagenta
12260             g.CellForeColor = 1
12270             g.CellFontBold = True
12280             g.row = g.row - 1
12290             g.CellBackColor = vbMagenta
12300             g.CellForeColor = 1
12310             g.CellFontBold = True
12320         End If
12330         intLine = intLine + 1
12340         tb.MoveNext
12350     Loop

12360     If g.Rows > 2 Then
12370         g.RemoveItem 1
12380     End If
12390     g.Visible = True

12400     Exit Sub

FillParameters_Error:

          Dim strES As String
          Dim intEL As Integer

12410     intEL = Erl
12420     strES = Err.Description
12430     LogError "fBioTodayQC", "FillParameters", intEL, strES, sql

End Sub

Private Sub cmdRefresh_Click()

12440     FillParameters

End Sub


Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

12450     On Error GoTo Form_Load_Error

12460     dt = Format$(Now, "dd/mm/yyyy")

12470     g.ColWidth(8) = 0

12480     sql = "Select distinct ControlName from BioQCDefs"
12490     Set tb = New Recordset
12500     RecOpenServer 0, tb, sql

12510     cmbControl.Clear

12520     Do While Not tb.EOF
12530         cmbControl.AddItem tb!ControlName & ""
12540         tb.MoveNext
12550     Loop

12560     FillParameters

12570     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

12580     intEL = Erl
12590     strES = Err.Description
12600     LogError "fBioTodayQC", "Form_Load", intEL, strES, sql


End Sub

Private Sub g_Click()

          Dim TestName As String
          Dim s As String
          Dim sql As String
          Dim NewValue As String
          Dim tb As Recordset
          Dim Alias As String

12610     On Error GoTo g_Click_Error

12620     If g.MouseRow = 0 Then Exit Sub

12630     TestName = g.TextMatrix(g.row, 0)
12640     If TestName = "" Then Exit Sub

12650     sql = "Select AliasName from BioQCDefs where ControlName = '" & cmbControl & "'"
12660     Set tb = New Recordset
12670     RecOpenServer 0, tb, sql
12680     If tb.EOF Then Exit Sub
12690     Alias = tb!AliasName & ""

12700     g.Col = 0
12710     Select Case g.MouseCol
              Case 0:
12720             If g.CellBackColor = vbMagenta Then
12730                 s = "Do you want to remove references to " & _
                          TestName & " (" & cmbControl & ") ran on " & dt & _
                          " at " & g.TextMatrix(g.row, 1) & ". " & _
                          "You will not be able to undo these changes.  " & vbCrLf & _
                          "Do you wish to continue?"
12740                 If iMsg(s, vbQuestion + vbYesNo, , vbRed) <> vbYes Then
12750                     Exit Sub
12760                 End If
          
12770                 sql = "delete from BiochemistryQC where AliasName = '" & Alias & "' " & _
                          "and Code = '" & g.TextMatrix(g.row, 8) & "' " & _
                          "and runtime = " & _
                          "'" & Format(dt, "dd/mmm/yyyy") & " " & g.TextMatrix(g.row, 1) & "'"
12780                 Cnxn(0).Execute sql
12790             Else
12800                 s = "By pressing 'Delete', you will remove all references to " & _
                          TestName & " (" & cmbControl & ") on " & dt & _
                          " from the control file. " & _
                          "   You  will also remove all " & TestName & " results " & _
                          "ran on " & dt & ". " & vbCrLf & _
                          "You will not be able to undo these changes.  " & vbCrLf & _
                          "Do you wish to continue?"
12810                 If iMsg(s, vbQuestion + vbYesNo, , vbRed) <> vbYes Then
12820                     Exit Sub
12830                 End If
          
12840                 sql = "delete from BiochemistryQC where AliasName = '" & Alias & "' " & _
                          "and Code = '" & g.TextMatrix(g.row, 8) & "' " & _
                          "and rundate = " & _
                          "'" & Format(dt, "dd/mmm/yyyy") & "'"
12850                 Cnxn(0).Execute sql
          
12860                 sql = "delete from bioResults where " & _
                          "Code = '" & g.TextMatrix(g.row, 8) & "' " & _
                          "and RunDate = " & _
                          "'" & Format(dt, "dd/mmm/yyyy") & "'"
12870                 Cnxn(0).Execute sql
12880             End If
          
12890             FillParameters

12900         Case 1, 2, 3:
12910         Case 4, 5, 6, 7:
12920             s = Choose(g.Col - 3, "Mean", "Low", "High", "1 SD") & _
                      " Value for " & TestName & "?"
12930             NewValue = iBOX(s, , g.TextMatrix(g.row, g.Col))
12940             g.TextMatrix(g.row, g.Col) = Format$(Val(NewValue))
          
12950             sql = "Select * from BioQCDefs where " & _
                      "ControlName = '" & cmbControl & "' " & _
                      "and ParameterName = '" & TestName & "'"
12960             Set tb = New Recordset
12970             RecOpenServer 0, tb, sql
12980             If tb.EOF Then
12990                 tb.AddNew
13000                 tb!ControlName = cmbControl
13010                 tb!ParameterName = TestName
13020                 tb!AliasName = Alias
13030             End If
13040             tb(Choose(g.Col - 3, "Mean", "Low", "High", "SD")) = Val(NewValue)
13050             tb.Update

13060     End Select

13070     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

13080     intEL = Erl
13090     strES = Err.Description
13100     LogError "fBioTodayQC", "g_Click", intEL, strES, sql

End Sub


