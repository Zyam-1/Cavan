VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestUnitsPrecision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Units and Precision"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   1155
      HelpContextID   =   10026
      Left            =   7230
      Picture         =   "frmTestUnitsPrecision.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6030
      Width           =   1065
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   7230
      Picture         =   "frmTestUnitsPrecision.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2100
      Width           =   1065
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   7020
      TabIndex        =   5
      Text            =   "cmbCategory"
      Top             =   900
      Width           =   1485
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   7020
      TabIndex        =   2
      Text            =   "cmbHospital"
      Top             =   1440
      Width           =   1485
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      HelpContextID   =   10010
      Left            =   7020
      TabIndex        =   1
      Text            =   "cmbSampleType"
      Top             =   360
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7005
      HelpContextID   =   10090
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   5
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
      FormatString    =   "<Long Name                 |<Short Name |^Units               |^Dec. Places |^Code"
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
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   7230
      TabIndex        =   9
      Top             =   3330
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   7440
      TabIndex        =   6
      Top             =   690
      Width           =   630
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestUnitsPrecision.frx":11D4
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestUnitsPrecision.frx":14AA
      Top             =   300
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   7470
      TabIndex        =   4
      Top             =   1230
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Type"
      Height          =   195
      Left            =   7290
      TabIndex        =   3
      Top             =   150
      Width           =   930
   End
End
Attribute VB_Name = "frmTestUnitsPrecision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String

Private pSampleType As String


Private strUnits() As String

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim strS As String

22090 On Error GoTo FillG_Error

22100 If cmbHospital = "" Then
22110   cmbHospital = HospName(0)
22120 End If
22130 If cmbCategory = "" Then
22140   cmbCategory = "Human"
22150 End If

22160 With g
22170   .Rows = 2
22180   .AddItem ""
22190   .RemoveItem 1
22200   .Visible = False

      '<Long Name  |<Short Name |^Units|^Dec.Places|^Code    (0 to 4)

22210   sql = "SELECT LongName, ShortName, Code, Units, DP, MAX(PrintPriority) P " & _
              "FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE SampleType = '" & pSampleType & "' " & _
              "AND Category = '" & cmbCategory & "' " & _
              "AND Hospital = '" & cmbHospital & "' " & _
              "" & _
              "GROUP BY Code, LongName, ShortName, Units, DP " & _
              "ORDER BY P"
22220   Set tb = New Recordset
22230   RecOpenServer 0, tb, sql
22240   Do While Not tb.EOF
22250     strS = tb!LongName & vbTab & _
                 tb!ShortName & vbTab & _
                 tb!Units & vbTab & _
                 tb!DP & vbTab & _
                 tb!Code & ""
22260     .AddItem strS
22270     tb.MoveNext
22280   Loop

22290   If .Rows > 2 Then
22300     .RemoveItem 1
22310   End If
22320   .Visible = True

22330   .MergeCells = flexMergeRestrictColumns
22340   .MergeCol(0) = True
22350   .MergeCol(1) = True
22360 End With

22370 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

22380 intEL = Erl
22390 strES = Err.Description
22400 LogError "frmTestUnitsPrecision", "Fillg", intEL, strES, sql

End Sub

Private Sub cmbCategory_Click()

22410 FillG

End Sub

Private Sub cmbHospital_Click()

22420 FillG

End Sub

Private Sub cmbSampleType_Click()

22430 pSampleType = ListCodeFor("ST", cmbSampleType)

22440 FillG

End Sub


Private Sub cmdCancel_Click()

22450 Unload Me

End Sub

Private Sub cmdXL_Click()

22460 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

22470 FillSampleTypes
22480 FillCategories
22490 FillHospitals

22500 FillG

End Sub


Private Sub FillCategories()

      Dim sql As String
      Dim tb As Recordset

22510 On Error GoTo FillCategories_Error

22520 sql = "Select Cat from Categorys " & _
            "order by ListOrder"
22530 Set tb = New Recordset
22540 RecOpenServer 0, tb, sql

22550 cmbCategory.Clear
22560 Do While Not tb.EOF
22570   cmbCategory.AddItem tb!Cat & ""
22580   tb.MoveNext
22590 Loop

22600 If cmbCategory.ListCount > 0 Then
22610   cmbCategory.ListIndex = 0
22620 End If

22630 Exit Sub

FillCategories_Error:

      Dim strES As String
      Dim intEL As Integer

22640 intEL = Erl
22650 strES = Err.Description
22660 LogError "frmTestUnitsPrecision", "FillCategories", intEL, strES, sql


End Sub


Private Sub FillSampleTypes()

      Dim sql As String
      Dim tb As Recordset

22670 On Error GoTo FillSampleTypes_Error

22680 sql = "Select Text from Lists where " & _
            "ListType = 'ST' " & _
            "order by ListOrder"
22690 Set tb = New Recordset
22700 RecOpenServer 0, tb, sql

22710 cmbSampleType.Clear
22720 Do While Not tb.EOF
22730   cmbSampleType.AddItem tb!Text & ""
22740   tb.MoveNext
22750 Loop

22760 If pSampleType <> "" Then
22770   cmbSampleType = ListTextFor("ST", pSampleType)
22780 Else
22790   pSampleType = "S"
22800   cmbSampleType = "Serum"
22810 End If

22820 Exit Sub

FillSampleTypes_Error:

      Dim strES As String
      Dim intEL As Integer

22830 intEL = Erl
22840 strES = Err.Description
22850 LogError "frmTestUnitsPrecision", "FillSampleTypes", intEL, strES, sql


End Sub

Private Sub FillHospitals()

      Dim sql As String
      Dim tb As Recordset

22860 On Error GoTo FillHospitals_Error

22870 sql = "Select Text from Lists where " & _
            "ListType = 'HO' " & _
            "order by ListOrder"
22880 Set tb = New Recordset
22890 RecOpenServer 0, tb, sql

22900 cmbHospital.Clear
22910 Do While Not tb.EOF
22920   cmbHospital.AddItem tb!Text & ""
22930   tb.MoveNext
22940 Loop

22950 cmbHospital.Text = HospName(0)

22960 Exit Sub

FillHospitals_Error:

      Dim strES As String
      Dim intEL As Integer

22970 intEL = Erl
22980 strES = Err.Description
22990 LogError "frmTestUnitsPrecision", "FillHospitals", intEL, strES, sql


End Sub


Private Sub FillStrUnits()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

23000 On Error GoTo FillStrUnits_Error

23010 sql = "Select Text from Lists where " & _
            "ListType = 'UN' " & _
            "order by ListOrder"
23020 Set tb = New Recordset
23030 RecOpenServer 0, tb, sql

23040 ReDim strUnits(0 To 0)
23050 strUnits(0) = ""
23060 n = 1

23070 Do While Not tb.EOF

23080   ReDim Preserve strUnits(0 To n)
23090   strUnits(n) = tb!Text & ""
        
23100   n = n + 1
        
23110   tb.MoveNext
        
23120 Loop

23130 Exit Sub

FillStrUnits_Error:

      Dim strES As String
      Dim intEL As Integer

23140 intEL = Erl
23150 strES = Err.Description
23160 LogError "frmTestUnitsPrecision", "FillStrUnits", intEL, strES, sql

End Sub


Public Property Let Discipline(ByVal sNewValue As String)

23170 pDiscipline = sNewValue

End Property
Public Property Let SampleType(ByVal sNewValue As String)

23180 pSampleType = sNewValue

End Property

Private Sub g_Click()

      Dim f As Form

      '<Long Name  |<Short Name (0 to 1)
      '^Units  |^Dec.Pl         (2 to 3)
      '^Code                    (4)

23190 On Error GoTo g_Click_Error

23200 Select Case g.MouseCol
        
        Case 2
23210     FillStrUnits
23220     Set f = New fcdrDBox
23230     With f
23240       .Options = strUnits
23250       .Prompt = "Enter Units for " & g.TextMatrix(g.row, 0)
23260       .Show 1
23270       g = .ReturnValue
23280     End With
23290     Unload f
23300     Set f = Nothing
23310     SaveCodes "Units", g.TextMatrix(g.row, 2)
        
23320   Case 3 'Decimal Places
23330     Select Case g.TextMatrix(g.row, 3)
            Case "0": g.TextMatrix(g.row, 3) = "1"
23340       Case "1": g.TextMatrix(g.row, 3) = "2"
23350       Case "2": g.TextMatrix(g.row, 3) = "3"
23360       Case Else: g.TextMatrix(g.row, 3) = "0"
23370     End Select
23380     SaveCodes "DP", g.TextMatrix(g.row, 3)
          
23390 End Select

23400 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

23410 intEL = Erl
23420 strES = Err.Description
23430 LogError "frmTestUnitsPrecision", "g_Click", intEL, strES

End Sub

Private Sub SaveCodes(ByVal FieldName As String, _
                      ByVal FieldValue As String)

      Dim sql As String

23440 On Error GoTo SaveCodes_Error

23450 If cmbHospital = "" Then
23460   cmbHospital = HospName(0)
23470 End If
23480 If cmbCategory = "" Then
23490   cmbCategory = "Human"
23500 End If

23510 sql = "Update " & pDiscipline & "TestDefinitions " & _
            "Set " & FieldName & " = '" & FieldValue & "' " & _
            "WHERE Code = '" & g.TextMatrix(g.row, 4) & "' " & _
            "AND SampleType = '" & pSampleType & "' " & _
            "AND Category = '" & cmbCategory & "' " & _
            "AND Hospital = '" & cmbHospital & "'"
23520 Cnxn(0).Execute sql

23530 Exit Sub

SaveCodes_Error:

      Dim strES As String
      Dim intEL As Integer

23540 intEL = Erl
23550 strES = Err.Description
23560 LogError "frmTestUnitsPrecision", "SaveCodes", intEL, strES, sql


End Sub




