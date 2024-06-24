VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestInUse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   9120
      Picture         =   "frmTestInUse.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4590
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   1155
      HelpContextID   =   10026
      Left            =   9120
      Picture         =   "frmTestInUse.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6060
      Width           =   1065
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   9120
      Picture         =   "frmTestInUse.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2100
      Width           =   1065
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   8910
      TabIndex        =   5
      Text            =   "cmbCategory"
      Top             =   900
      Width           =   1485
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   8910
      TabIndex        =   2
      Text            =   "cmbHospital"
      Top             =   1440
      Width           =   1485
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      HelpContextID   =   10010
      Left            =   8910
      TabIndex        =   1
      Text            =   "cmbSampleType"
      Top             =   360
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid gInUse 
      Height          =   7005
      HelpContextID   =   10090
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   6
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
      FormatString    =   "<Long Name                 |<Short Name |^Code    |^In Use      |^Healthlink    |^Accredited"
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
      Left            =   9120
      TabIndex        =   9
      Top             =   3330
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   9330
      TabIndex        =   6
      Top             =   690
      Width           =   630
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestInUse.frx":209E
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestInUse.frx":2374
      Top             =   300
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   9360
      TabIndex        =   4
      Top             =   1230
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Type"
      Height          =   195
      Left            =   9180
      TabIndex        =   3
      Top             =   150
      Width           =   930
   End
End
Attribute VB_Name = "frmTestInUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String

Private pSampleType As String

Private Sub FillgInUse()

      Dim tb As Recordset
      Dim sql As String
      Dim strS As String

10140 On Error GoTo FillgInUse_Error
10150 cmdSave.Enabled = False
10160 If cmbHospital = "" Then
10170     cmbHospital = HospName(0)
10180 End If
10190 If cmbCategory = "" Then
10200     cmbCategory = "Human"
10210 End If

10220 gInUse.Rows = 2
10230 gInUse.AddItem ""
10240 gInUse.RemoveItem 1
10250 gInUse.Visible = False

      '<Long Name  |<Short Name |^Code |^In Use    (0 to 3)

10260 sql = "SELECT LongName, ShortName, Code, MAX(COALESCE(InUse, 0)) InUse, MAX(PrintPriority) P ,ISNULL(HealthLink,0) as Healthlink,ISNULL(Accredited,0) as Accredited " & _
            "FROM " & pDiscipline & "TestDefinitions " & _
            "WHERE SampleType = '" & pSampleType & "' " & _
            "AND Category = '" & cmbCategory & "' " & _
            "AND Hospital = '" & cmbHospital & "' " & _
            "GROUP BY Code, LongName, ShortName,HealthLink,Accredited " & _
            "ORDER BY P"
10270 Set tb = New Recordset
10280 RecOpenServer 0, tb, sql
10290 Do While Not tb.EOF
10300     strS = tb!LongName & vbTab & _
                 tb!ShortName & vbTab & _
                 tb!Code & "" & vbTab

10310     gInUse.AddItem strS

10320     gInUse.row = gInUse.Rows - 1
10330     gInUse.Col = 3
10340     Set gInUse.CellPicture = IIf(tb!InUse, imgSquareTick.Picture, imgSquareCross.Picture)
10350     gInUse.CellPictureAlignment = flexAlignCenterCenter

10360     gInUse.row = gInUse.Rows - 1
10370     gInUse.Col = 4
10380     Set gInUse.CellPicture = IIf((tb!HealthLink = 1), imgSquareTick.Picture, imgSquareCross.Picture)
10390     gInUse.CellPictureAlignment = flexAlignCenterCenter

10400     gInUse.row = gInUse.Rows - 1
10410     gInUse.Col = 5
10420     Set gInUse.CellPicture = IIf((tb!Accredited = 1), imgSquareTick.Picture, imgSquareCross.Picture)
10430     gInUse.CellPictureAlignment = flexAlignCenterCenter
10440     tb.MoveNext
10450 Loop

10460 If gInUse.Rows > 2 Then
10470     gInUse.RemoveItem 1
10480 End If
10490 gInUse.Visible = True

10500 gInUse.MergeCells = flexMergeRestrictColumns
10510 gInUse.MergeCol(0) = True
10520 gInUse.MergeCol(1) = True

10530 Exit Sub

FillgInUse_Error:

      Dim strES As String
      Dim intEL As Integer

10540 intEL = Erl
10550 strES = Err.Description
10560 LogError "frmBioRanges", "FillgInUse", intEL, strES, sql

End Sub

Private Sub cmbCategory_Click()

10570 FillgInUse

End Sub

Private Sub cmbHospital_Click()

10580 FillgInUse

End Sub

Private Sub cmbSampleType_Click()

10590 pSampleType = ListCodeFor("ST", cmbSampleType)

10600 FillgInUse

End Sub


Private Sub cmdCancel_Click()

10610 Unload Me

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSave_Click
' Author    : XPMUser
' Date      : 17/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSave_Click()
      Dim sql As String
      Dim i As Integer
10620 On Error GoTo cmdSave_Click_Error

10630 Cnxn(0).Execute ("UPDATE " & pDiscipline & "TestDefinitions SET InUse = 0 , Healthlink = 0 , Accredited = 0 WHERE SampleType ='" & pSampleType & "'")
10640 With gInUse
10650     For i = 1 To .Rows - 1
10660         .row = i
10670         .Col = 3
10680         If .CellPicture = imgSquareTick.Picture Then
10690             sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
                        " SET InUse =  1 " & _
                        " WHERE LongName = '" & .TextMatrix(i, 0) & "' " & _
                        " AND SampleType = '" & pSampleType & "'"
10700             Cnxn(0).Execute sql
10710         End If
10720         .Col = 4
10730         If .CellPicture = imgSquareTick.Picture Then
10740             sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
                        " SET Healthlink  =  1 " & _
                        " WHERE LongName = '" & .TextMatrix(i, 0) & "' " & _
                        " AND SampleType = '" & pSampleType & "'"
10750             Cnxn(0).Execute sql
10760         End If

10770         .Col = 5
10780         If .CellPicture = imgSquareTick.Picture Then
10790             sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
                        " SET Accredited  =  1 " & _
                        " WHERE LongName = '" & .TextMatrix(i, 0) & "' " & _
                        " AND SampleType = '" & pSampleType & "'"
10800             Cnxn(0).Execute sql
10810         End If
10820     Next i
10830 End With

10840 FillgInUse

10850 Exit Sub


cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

10860 intEL = Erl
10870 strES = Err.Description
10880 LogError "frmTestInUse", "cmdSave_Click", intEL, strES, sql
End Sub

Private Sub cmdXL_Click()

10890 ExportFlexGrid gInUse, Me

End Sub


Private Sub Form_Load()

10900 FillSampleTypes
10910 FillCategories
10920 FillHospitals

10930 FillgInUse

End Sub


Private Sub FillCategories()

      Dim sql As String
      Dim tb As Recordset

10940 On Error GoTo FillCategories_Error

10950 sql = "Select Cat from Categorys " & _
            "order by ListOrder"
10960 Set tb = New Recordset
10970 RecOpenServer 0, tb, sql

10980 cmbCategory.Clear
10990 Do While Not tb.EOF
11000   cmbCategory.AddItem tb!Cat & ""
11010   tb.MoveNext
11020 Loop

11030 If cmbCategory.ListCount > 0 Then
11040   cmbCategory.ListIndex = 0
11050 End If

11060 Exit Sub

FillCategories_Error:

      Dim strES As String
      Dim intEL As Integer

11070 intEL = Erl
11080 strES = Err.Description
11090 LogError "frmTestInUse", "FillCategories", intEL, strES, sql


End Sub


Private Sub FillSampleTypes()

      Dim sql As String
      Dim tb As Recordset

11100 On Error GoTo FillSampleTypes_Error

11110 sql = "Select Text from Lists where " & _
            "ListType = 'ST' " & _
            "order by ListOrder"
11120 Set tb = New Recordset
11130 RecOpenServer 0, tb, sql

11140 cmbSampleType.Clear
11150 Do While Not tb.EOF
11160   cmbSampleType.AddItem tb!Text & ""
11170   tb.MoveNext
11180 Loop

11190 If pSampleType <> "" Then
11200   cmbSampleType = ListTextFor("ST", pSampleType)
11210 Else
11220   pSampleType = "S"
11230   cmbSampleType = "Serum"
11240 End If

11250 Exit Sub

FillSampleTypes_Error:

      Dim strES As String
      Dim intEL As Integer

11260 intEL = Erl
11270 strES = Err.Description
11280 LogError "frmTestInUse", "FillSampleTypes", intEL, strES, sql


End Sub

Private Sub FillHospitals()

      Dim sql As String
      Dim tb As Recordset

11290 On Error GoTo FillHospitals_Error

11300 sql = "Select Text from Lists where " & _
            "ListType = 'HO' " & _
            "order by ListOrder"
11310 Set tb = New Recordset
11320 RecOpenServer 0, tb, sql

11330 cmbHospital.Clear
11340 Do While Not tb.EOF
11350   cmbHospital.AddItem tb!Text & ""
11360   tb.MoveNext
11370 Loop

11380 cmbHospital.Text = HospName(0)

11390 Exit Sub

FillHospitals_Error:

      Dim strES As String
      Dim intEL As Integer

11400 intEL = Erl
11410 strES = Err.Description
11420 LogError "frmTestInUse", "FillHospitals", intEL, strES, sql


End Sub


Private Sub gInUse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      'Dim sql As String
      '
11430 On Error GoTo gInUse_MouseUp_Error

11440 With gInUse
11450     .row = .RowSel
11460     .Col = .ColSel
11470     If (.ColSel = 5 Or .ColSel = 4 Or .ColSel = 3) And pDiscipline = "Bio" Then
11480         cmdSave.Enabled = True
11490         Set .CellPicture = IIf((.CellPicture = imgSquareCross.Picture), imgSquareTick.Picture, imgSquareCross.Picture)
11500     End If
11510 End With


      '
      ''<Long Name  |<Short Name|^Code |^In Use    (0 to 3)
      '
      'With gInUse
      '  If .ColSel = 4 And pDiscipline = "Bio" Then
      '    .row = .RowSel
      '    .Col = 4
      '      sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
             '          " SET Healthlink =  " & IIf((.CellPicture = imgSquareCross.Picture), 1, 0) & _
             '          " WHERE LongName = '" & .TextMatrix(.row, 0) & "' " & _
             '          " AND SampleType = '" & pSampleType & "'"
      '    Cnxn(0).Execute Sql
      '    DoEvents
      '    FillgInUse
      '    Exit Sub
      '  End If
      '
      '  If .MouseRow = 0 Or .MouseCol <> 3 Then Exit Sub
      '
      '  .row = .MouseRow
      '  .Col = 3
      '
      '  If .CellPicture = imgSquareCross.Picture Then
      '
      '    sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
           '          "SET InUse = 0 " & _
           '          "WHERE LongName = '" & .TextMatrix(.row, 0) & "' " & _
           '          "AND SampleType = '" & pSampleType & "'"
      '    Cnxn(0).Execute Sql
      '
      '    sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
           '          "SET InUse = 1 " & _
           '          "WHERE Code = '" & .TextMatrix(.row, 2) & "' " & _
           '          "AND SampleType = '" & pSampleType & "'"
      '    Cnxn(0).Execute Sql
      '
      '  Else
      '
      '    sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
           '          "SET InUse = 0 " & _
           '          "WHERE Code = '" & .TextMatrix(.row, 2) & "' " & _
           '          "AND SampleType = '" & pSampleType & "'"
      '    Cnxn(0).Execute Sql
      '
      '  End If
      '
      'End With
      '
      'FillgInUse

11520 Exit Sub

gInUse_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

11530 intEL = Erl
11540 strES = Err.Description
11550 LogError "frmBioRanges", "gInUse_MouseUp", intEL, strES

End Sub




Public Property Let Discipline(ByVal sNewValue As String)

11560 pDiscipline = sNewValue

End Property
Public Property Let SampleType(ByVal sNewValue As String)

11570 pSampleType = sNewValue

End Property

