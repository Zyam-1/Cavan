VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRemoveBioTest 
   Caption         =   "NetAcquire"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMarkAllAs 
      Caption         =   "Mark All as Not Printed"
      Height          =   645
      Left            =   8070
      TabIndex        =   12
      Top             =   90
      Width           =   1335
   End
   Begin VB.CommandButton cmdMarkAs 
      Caption         =   "Mark as Not Printed"
      Height          =   645
      Left            =   6540
      TabIndex        =   11
      Top             =   90
      Width           =   1335
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   525
      Left            =   7560
      Picture         =   "frmRemoveBioTest.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Analyte from Results"
      Height          =   525
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   3255
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Remove All Results"
      Height          =   525
      Left            =   180
      TabIndex        =   4
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   6540
      Picture         =   "frmRemoveBioTest.frx":030A
      TabIndex        =   3
      Top             =   870
      Width           =   2865
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "UnDelete selected Analyte(s)"
      Enabled         =   0   'False
      Height          =   525
      Left            =   3480
      TabIndex        =   2
      Top             =   4680
      Width           =   2385
   End
   Begin VB.CommandButton cmdViewArchive 
      Caption         =   "No Deleted Results"
      Height          =   525
      Left            =   3930
      TabIndex        =   0
      Top             =   840
      Width           =   2085
   End
   Begin MSFlexGridLib.MSFlexGrid gArchive 
      Height          =   2715
      Left            =   540
      TabIndex        =   1
      Top             =   1710
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4789
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   "<Analyte              |^Type    |<Raw Result   |<Run Date/Time             |<Date/Time Archived |<Archived By      |<Code"
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6180
      TabIndex        =   10
      Top             =   4770
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Analyte(s) to be retrieved"
      Height          =   255
      Left            =   570
      TabIndex        =   8
      Top             =   4920
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   780
      Picture         =   "frmRemoveBioTest.frx":0974
      Top             =   4470
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   3930
      TabIndex        =   7
      Top             =   300
      Width           =   735
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4710
      TabIndex        =   6
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frmRemoveBioTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private pAnalyteCode As String
Private pAnalyte As String

Private pSampleID As Long

Private SortOrder As Boolean

Private pDiscipline As String

Public Property Get Analyte() As String

43510 Analyte = pAnalyte

End Property

Public Property Get AnalyteCode() As String

43520 AnalyteCode = pAnalyteCode

End Property


Public Property Let Analyte(ByVal strNewValue As String)

43530 pAnalyte = strNewValue

End Property
Public Property Let AnalyteCode(ByVal strNewValue As String)

43540 pAnalyteCode = strNewValue

End Property

Private Sub LoadIt()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

43550 On Error GoTo Form_Load_Error

43560 sql = "SELECT COALESCE(Printed, 0) Printed FROM BioResults " & _
            "WHERE SampleID = '" & pSampleID & "' " & _
            "AND Code = '" & pAnalyteCode & "'"
43570 Set tb = New Recordset
43580 RecOpenServer 0, tb, sql
43590 If Not tb.EOF Then
43600   If tb!Printed = 0 Then
43610     cmdMarkAs.Caption = "Mark as Printed"
43620   Else
43630     cmdMarkAs.Caption = "Mark as not Printed"
43640   End If
43650   cmdMarkAs.Visible = True
43660 Else
43670   cmdMarkAs.Visible = False
43680 End If

43690 cmdRemove.Caption = "Remove " & pAnalyte & " from Results."
43700 lblSampleID.Caption = pSampleID

43710 gArchive.ColWidth(6) = 0

43720 Me.height = 2040

43730 sql = "SELECT A.SampleType, A.Result, A.RunTime, A.Code, " & _
            "A.ArchiveDateTime, A.ArchivedBy, D.LongName " & _
            "FROM " & pDiscipline & "ResultsAudit A " & _
            "JOIN " & pDiscipline & "TestDefinitions D  " & _
            "ON A.Code = D.Code " & _
            "WHERE SampleID = '" & pSampleID & "' " & _
            "ORDER BY D.PrintPriority"
43740 Set tb = New Recordset
43750 RecOpenServer 0, tb, sql
43760 If Not tb.EOF Then
43770   cmdViewArchive.Enabled = True
43780   cmdViewArchive.Caption = "View Deleted Results"
43790   With gArchive
43800     .Rows = 2
43810     .AddItem ""
43820     .RemoveItem 1
43830   End With

43840   Do While Not tb.EOF
43850     s = tb!LongName & vbTab & _
              tb!SampleType & vbTab & _
              tb!Result & vbTab & _
              Format$(tb!RunTime, "dd/MM/yy HH:mm") & vbTab & _
              Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss") & vbTab & _
              tb!ArchivedBy & vbTab & _
              tb!Code & ""
43860     gArchive.AddItem s
43870     tb.MoveNext
43880   Loop
43890   If gArchive.Rows > 2 Then
43900     gArchive.RemoveItem 1
43910   End If
43920 Else
43930   cmdViewArchive.Enabled = False
43940   cmdViewArchive.Caption = "No Deleted Results"
43950 End If

43960 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

43970 intEL = Erl
43980 strES = Err.Description
43990 LogError "frmRemoveBioTest", "LoadIt", intEL, strES, sql

End Sub

Public Property Let SampleID(ByVal lngNewValue As Long)

44000 pSampleID = lngNewValue

End Property

Private Sub cmdCancel_Click()

44010 pAnalyte = ""
44020 pAnalyteCode = ""
44030 Me.Hide

End Sub


Private Sub cmdClearAll_Click()

44040 If iMsg("Clear All Results?" & vbCrLf & "You will not be able to undo this action!", vbQuestion + vbYesNo, "Confirm", vbRed, 12) = vbYes Then
44050   If UCase(iBOX("Enter your password", , , True)) = UCase$(TechnicianPassFor(UserName)) Then
44060     pAnalyte = "All"
44070   Else
44080     pAnalyte = ""
44090     pAnalyteCode = ""
44100   End If
44110 Else
44120   pAnalyte = ""
44130   pAnalyteCode = ""
44140 End If

44150 Me.Hide

End Sub


Private Sub cmdMarkAllAs_Click()

      Dim sql As String
      Dim NewValue As Integer

44160 NewValue = IIf(cmdMarkAllAs.Caption = "Mark All as Printed", 1, 0)
44170 sql = "UPDATE BioResults " & _
            "SET Printed = '" & NewValue & "' " & _
            "WHERE SampleID = '" & pSampleID & "' "
44180 Cnxn(0).Execute sql

44190 If NewValue = 1 Then
44200   cmdMarkAllAs.Caption = "Mark All as not Printed"
44210 Else
44220   cmdMarkAllAs.Caption = "Mark All as Printed"
44230 End If

44240 LoadIt

End Sub

Private Sub cmdMarkAs_Click()

      Dim sql As String
      Dim NewValue As Integer

44250 On Error GoTo cmdMarkAs_Click_Error

44260 NewValue = IIf(cmdMarkAs.Caption = "Mark as Printed", 1, 0)
44270 sql = "UPDATE BioResults " & _
            "SET Printed = '" & NewValue & "' " & _
            "WHERE SampleID = '" & pSampleID & "' " & _
            "AND Code = '" & pAnalyteCode & "'"
44280 Cnxn(0).Execute sql

44290 LoadIt

44300 Exit Sub

cmdMarkAs_Click_Error:

      Dim strES As String
      Dim intEL As Integer

44310 intEL = Erl
44320 strES = Err.Description
44330 LogError "frmRemoveBioTest", "cmdMarkAs_Click", intEL, strES, sql

End Sub

Private Sub cmdRemove_Click()

44340 Me.Hide

End Sub


Private Sub cmdRetrieve_Click()

      Dim sql As String
      Dim ADT As String
      Dim Code As String
      Dim n As Integer

44350 On Error GoTo cmdRetrieve_Click_Error

44360 gArchive.Col = 0
44370 For n = 1 To gArchive.Rows - 1
44380   gArchive.row = n
44390   If gArchive.CellBackColor = vbYellow Then
44400     Code = gArchive.TextMatrix(n, 6)
44410     ADT = Format$(gArchive.TextMatrix(n, 4), "dd/MMM/yyyy HH:nn:ss")
       
44420 sql = "DELETE FROM " & pDiscipline & "Results WHERE Code = '" & Code & "' AND SampleID = '" & lblSampleID & "'"
44430 Cnxn(0).Execute sql

44440     sql = "INSERT INTO " & pDiscipline & "Results " & _
               "(SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
               " Operator, Flags, Units, SampleType, Analyser, Faxed, Authorised) " & _
               " SELECT SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
               " Operator, Flags, Units, SampleType, Analyser, Faxed, Authorised " & _
               " FROM " & pDiscipline & "ResultsAudit WHERE " & _
               " Code = '" & Code & "' " & _
               " AND SampleID = '" & lblSampleID & "' " & _
               " AND ABS(DATEDIFF(Second, ArchiveDateTime,'" & ADT & "')) < 2"
44450     Cnxn(0).Execute sql
44460   End If
44470 Next

44480 pAnalyte = ""
44490 pAnalyteCode = ""
44500 Me.Hide

44510 Exit Sub

cmdRetrieve_Click_Error:

      Dim strES As String
      Dim intEL As Integer

44520 intEL = Erl
44530 strES = Err.Description
44540 LogError "frmRemoveBioTest", "cmdRetrieve_Click", intEL, strES, sql

End Sub

Private Sub cmdViewArchive_Click()
        
44550 If cmdViewArchive.Caption = "View Deleted Results" Then
44560   cmdViewArchive.Caption = "Hide Deleted Results"
44570   Me.height = 6090
44580 Else
44590   cmdViewArchive.Caption = "View Deleted Results"
44600   Me.height = 2040
44610 End If

End Sub


Private Sub cmdXL_Click()

44620 ExportFlexGrid gArchive, Me

End Sub

Private Sub Form_Load()

44630 LoadIt

End Sub


Private Sub gArchive_Click()

      Dim n As Integer
      Dim ySave As Integer
      Dim Code As String

44640 On Error GoTo gArchive_Click_Error

44650 With gArchive

44660   If .MouseRow = 0 Then
44670     If InStr(.TextMatrix(0, .Col), "Date") <> 0 Then
44680       .Sort = 9
44690     Else
44700       If SortOrder Then
44710         .Sort = flexSortGenericAscending
44720       Else
44730         .Sort = flexSortGenericDescending
44740       End If
44750     End If
44760     SortOrder = Not SortOrder
44770     Exit Sub
44780   End If

44790   .Col = 0
44800   .row = .MouseRow
        
44810   If .CellBackColor = vbYellow Then
44820     .CellBackColor = 0
44830   Else
44840     Code = .TextMatrix(.row, 6)
44850     .CellBackColor = vbYellow
44860     ySave = .row
44870     For n = 1 To .Rows - 1
44880       If n <> ySave And .TextMatrix(n, 0) = .TextMatrix(ySave, 0) Then
44890         .row = n
44900         .CellBackColor = 0
44910       End If
44920     Next
44930   End If
        
44940   cmdRetrieve.Enabled = False
44950   For n = 1 To .Rows - 1
44960     .row = n
44970     If .CellBackColor = vbYellow Then
44980       cmdRetrieve.Enabled = True
44990       Exit For
45000     End If
45010   Next
        
45020 End With

45030 Exit Sub

gArchive_Click_Error:

      Dim strES As String
      Dim intEL As Integer

45040 intEL = Erl
45050 strES = Err.Description
45060 LogError "frmRemoveBioTest", "gArchive_Click", intEL, strES

End Sub


Private Sub gArchive_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

45070 With gArchive
45080   If Not IsDate(.TextMatrix(Row1, .Col)) Then
45090     Cmp = 0
45100     Exit Sub
45110   End If

45120   If Not IsDate(.TextMatrix(Row2, .Col)) Then
45130     Cmp = 0
45140     Exit Sub
45150   End If

45160   d1 = Format(.TextMatrix(Row1, .Col), "dd/mmm/yyyy hh:mm:ss")
45170   d2 = Format(.TextMatrix(Row2, .Col), "dd/mmm/yyyy hh:mm:ss")

45180   If SortOrder Then
45190     Cmp = Sgn(DateDiff("s", d1, d2))
45200   Else
45210     Cmp = Sgn(DateDiff("s", d2, d1))
45220   End If
45230 End With

End Sub



Public Property Let Discipline(ByVal sNewValue As String)

45240 pDiscipline = sNewValue

End Property
