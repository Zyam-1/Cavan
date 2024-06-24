VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestKnownToAnalyser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Active Analyser Requests"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnalysers 
      Caption         =   "Analyser List"
      Height          =   1155
      Left            =   5910
      Picture         =   "frmTestKnownToAnalyser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   1155
      HelpContextID   =   10026
      Left            =   5910
      Picture         =   "frmTestKnownToAnalyser.frx":0F72
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6060
      Width           =   1065
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   5910
      Picture         =   "frmTestKnownToAnalyser.frx":1E3C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3870
      Width           =   1065
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   5700
      TabIndex        =   5
      Text            =   "cmbCategory"
      Top             =   1020
      Width           =   1485
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   5700
      TabIndex        =   2
      Text            =   "cmbHospital"
      Top             =   1560
      Width           =   1485
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      HelpContextID   =   10010
      Left            =   5700
      TabIndex        =   1
      Text            =   "cmbSampleType"
      Top             =   480
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid gKnown 
      Height          =   6945
      HelpContextID   =   10090
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   12250
      _Version        =   393216
      Cols            =   4
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Long Name                 |<Short Name |^Code    |^Known To Analyser"
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
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   5910
      TabIndex        =   9
      Top             =   5100
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   6120
      TabIndex        =   6
      Top             =   810
      Width           =   630
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   30
      Picture         =   "frmTestKnownToAnalyser.frx":2146
      Top             =   180
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   30
      Picture         =   "frmTestKnownToAnalyser.frx":241C
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   6150
      TabIndex        =   4
      Top             =   1350
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Type"
      Height          =   195
      Left            =   5970
      TabIndex        =   3
      Top             =   270
      Width           =   930
   End
End
Attribute VB_Name = "frmTestKnownToAnalyser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String

Private pSampleType As String

Private AnalyserList() As String
Private Sub FillgKnown()

      Dim tb As Recordset
      Dim sql As String
      Dim strS As String
      Dim TopRow As Integer

11580 On Error GoTo FillgKnown_Error

11590 If cmbHospital = "" Then
11600   cmbHospital = HospName(0)
11610 End If
11620 If cmbCategory = "" Then
11630   cmbCategory = "Human"
11640 End If

11650 With gKnown
        
11660   TopRow = .TopRow
        
11670   .Rows = 2
11680   .AddItem ""
11690   .RemoveItem 1
11700   .Visible = False
        
        '<Long Name  |<Short Name |^Code |^Known    (0 to 3)
        
11710   sql = "SELECT LongName, ShortName, Code, MAX(COALESCE(KnownToAnalyser, 0)) Known, MAX(PrintPriority) P, Analyser " & _
              "FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE SampleType = '" & pSampleType & "' " & _
              "AND Category = '" & cmbCategory & "' " & _
              "AND Hospital = '" & cmbHospital & "' " & _
              "GROUP BY Code, LongName, ShortName, Analyser " & _
              "ORDER BY P"
11720   Set tb = New Recordset
11730   RecOpenServer 0, tb, sql
11740   Do While Not tb.EOF
11750     strS = tb!LongName & vbTab & _
                 tb!ShortName & vbTab & _
                 tb!Code & vbTab & _
                 tb!Analyser & ""

11760     .AddItem strS
11770     tb.MoveNext
11780   Loop
        
11790   If .Rows > 2 Then
11800     .RemoveItem 1
11810   End If
11820   .Visible = True
        
11830   .MergeCells = flexMergeRestrictColumns
11840   .MergeCol(0) = True
11850   .MergeCol(1) = True

11860   .TopRow = TopRow
        
11870 End With

11880 Exit Sub

FillgKnown_Error:

      Dim strES As String
      Dim intEL As Integer

11890 intEL = Erl
11900 strES = Err.Description
11910 LogError "frmTestKnownToAnalyser", "FillgKnown", intEL, strES, sql


End Sub

Private Sub GetAnalyserList()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

11920 On Error GoTo GetAnalyserList_Error

11930 sql = "SELECT Text FROM Lists " & _
            "WHERE ListType = 'Analyser' " & _
            "ORDER BY ListOrder"
11940 Set tb = New Recordset
11950 RecOpenServer 0, tb, sql
11960 ReDim AnalyserList(1 To 1) As String
11970 AnalyserList(1) = ""
11980 n = 2
11990 Do While Not tb.EOF
12000   ReDim Preserve AnalyserList(1 To n) As String
12010   AnalyserList(n) = tb!Text
12020   n = n + 1
12030   tb.MoveNext
12040 Loop

12050 Exit Sub

GetAnalyserList_Error:

      Dim strES As String
      Dim intEL As Integer

12060 intEL = Erl
12070 strES = Err.Description
12080 LogError "frmTestKnownToAnalyser", "GetAnalyserList", intEL, strES, sql

End Sub

Private Sub cmbCategory_Click()

12090 FillgKnown

End Sub

Private Sub cmbHospital_Click()

12100 FillgKnown

End Sub

Private Sub cmbSampleType_Click()

12110 pSampleType = ListCodeFor("ST", cmbSampleType)

12120 FillgKnown

End Sub


Private Sub cmdAnalysers_Click()

12130 With frmListsGeneric
12140   .ListType = "Analyser"
12150   .ListTypeName = "Analyser"
12160   .ListTypeNames = "Analysers"
12170   .Show 1
12180 End With

12190 GetAnalyserList

End Sub

Private Sub cmdCancel_Click()

12200 Unload Me

End Sub

Private Sub cmdXL_Click()

12210 ExportFlexGrid gKnown, Me

End Sub

Private Sub Form_Load()

12220 FillSampleTypes
12230 FillCategories
12240 FillHospitals
12250 GetAnalyserList

12260 FillgKnown

End Sub

Private Sub FillCategories()

      Dim sql As String
      Dim tb As Recordset

12270 On Error GoTo FillCategories_Error

12280 sql = "Select Cat from Categorys " & _
            "order by ListOrder"
12290 Set tb = New Recordset
12300 RecOpenServer 0, tb, sql

12310 cmbCategory.Clear
12320 Do While Not tb.EOF
12330   cmbCategory.AddItem tb!Cat & ""
12340   tb.MoveNext
12350 Loop

12360 If cmbCategory.ListCount > 0 Then
12370   cmbCategory.ListIndex = 0
12380 End If

12390 Exit Sub

FillCategories_Error:

      Dim strES As String
      Dim intEL As Integer

12400 intEL = Erl
12410 strES = Err.Description
12420 LogError "frmTestKnownToAnalyser", "FillCategories", intEL, strES, sql


End Sub


Private Sub FillSampleTypes()

      Dim sql As String
      Dim tb As Recordset

12430 On Error GoTo FillSampleTypes_Error

12440 sql = "Select Text from Lists where " & _
            "ListType = 'ST' " & _
            "order by ListOrder"
12450 Set tb = New Recordset
12460 RecOpenServer 0, tb, sql

12470 cmbSampleType.Clear
12480 Do While Not tb.EOF
12490   cmbSampleType.AddItem tb!Text & ""
12500   tb.MoveNext
12510 Loop

12520 cmbSampleType = "Serum"
12530 pSampleType = ListCodeFor("ST", "Serum")

12540 Exit Sub

FillSampleTypes_Error:

      Dim strES As String
      Dim intEL As Integer

12550 intEL = Erl
12560 strES = Err.Description
12570 LogError "frmTestKnownToAnalyser", "FillSampleTypes", intEL, strES, sql


End Sub

Private Sub FillHospitals()

      Dim sql As String
      Dim tb As Recordset

12580 On Error GoTo FillHospitals_Error

12590 sql = "Select Text from Lists where " & _
            "ListType = 'HO' " & _
            "order by ListOrder"
12600 Set tb = New Recordset
12610 RecOpenServer 0, tb, sql

12620 cmbHospital.Clear
12630 Do While Not tb.EOF
12640   cmbHospital.AddItem tb!Text & ""
12650   tb.MoveNext
12660 Loop

12670 cmbHospital.Text = HospName(0)

12680 Exit Sub

FillHospitals_Error:

      Dim strES As String
      Dim intEL As Integer

12690 intEL = Erl
12700 strES = Err.Description
12710 LogError "frmTestKnownToAnalyser", "FillHospitals", intEL, strES, sql


End Sub


Private Sub gKnown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim sql As String
      Dim ThisIndex As Integer
      Dim n As Integer

      '<Long Name  |<Short Name|^Code |^Analyser    (0 to 3)

12720 On Error GoTo gKnown_MouseUp_Error

12730 With gKnown
12740   If .MouseRow = 0 Or .MouseCol <> 3 Then Exit Sub

12750   ThisIndex = 1
12760   .row = .MouseRow

12770   For n = 1 To UBound(AnalyserList)
12780     If .TextMatrix(.row, 3) = AnalyserList(n) Then
12790       ThisIndex = n
12800       Exit For
12810     End If
12820   Next
12830   If ThisIndex < UBound(AnalyserList) Then
12840     .TextMatrix(.row, 3) = AnalyserList(ThisIndex + 1)
12850   Else
12860     .TextMatrix(.row, 3) = AnalyserList(1)
12870   End If
        
12880   sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
              "SET Analyser = '" & .TextMatrix(.row, 3) & "' " & _
              "WHERE LongName = '" & .TextMatrix(.row, 0) & "' " & _
              "AND Code = '" & .TextMatrix(.row, 2) & "' " & _
              "AND SampleType = '" & pSampleType & "'"
12890   Cnxn(0).Execute sql
         
12900 End With

12910 FillgKnown

12920 Exit Sub

gKnown_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

12930 intEL = Erl
12940 strES = Err.Description
12950 LogError "frmTestKnownToAnalyser", "gKnown_MouseUp", intEL, strES, sql

End Sub




Public Property Let Discipline(ByVal sNewValue As String)

12960 pDiscipline = sNewValue

End Property
Public Property Let SampleType(ByVal sNewValue As String)

12970 pSampleType = sNewValue

End Property

