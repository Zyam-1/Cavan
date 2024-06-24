VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioMultiNormal 
   Caption         =   "NetAcquire - Biochemistry Normal Ranges"
   ClientHeight    =   7395
   ClientLeft      =   180
   ClientTop       =   480
   ClientWidth     =   12210
   ControlBox      =   0   'False
   HelpContextID   =   10017
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   12210
   Begin VB.Frame Frame4 
      Caption         =   "Discipline"
      Height          =   1005
      Left            =   9810
      TabIndex        =   22
      Top             =   630
      Width           =   2265
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   24
         Top             =   480
         Width           =   1275
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   23
         Top             =   720
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Legend"
      Height          =   2595
      Left            =   9840
      TabIndex        =   18
      Top             =   3510
      Width           =   2235
      Begin VB.Label Label6 
         Caption         =   "Historical Ranges are present.   Click to view."
         Height          =   585
         Left            =   390
         TabIndex        =   21
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   120
         Picture         =   "frmBioMultiNormal.frx":0000
         Top             =   2070
         Width           =   210
      End
      Begin VB.Label Label5 
         Caption         =   "Age Related Ranges are being displayed. Click to collapse."
         Height          =   675
         Left            =   390
         TabIndex        =   20
         Top             =   1170
         Width           =   1515
      End
      Begin VB.Image imgSquareMinus 
         Height          =   225
         Left            =   120
         Picture         =   "frmBioMultiNormal.frx":02D6
         Top             =   1380
         Width           =   210
      End
      Begin VB.Label Label4 
         Caption         =   "Age Related Ranges are present.        Click to Expand."
         Height          =   615
         Left            =   390
         TabIndex        =   19
         Top             =   300
         Width           =   1485
      End
      Begin VB.Image imgSquarePlus 
         Height          =   225
         Left            =   120
         Picture         =   "frmBioMultiNormal.frx":03D0
         Top             =   480
         Width           =   210
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Names"
      Height          =   915
      Left            =   9840
      TabIndex        =   15
      Top             =   2460
      Width           =   2235
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Long Names"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Short Names"
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   16
         Top             =   510
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ranges"
      Height          =   615
      Left            =   9840
      TabIndex        =   12
      Top             =   1740
      Width           =   2235
      Begin VB.OptionButton optRange 
         Caption         =   "Flags"
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   300
         Width           =   675
      End
      Begin VB.OptionButton optRange 
         Alignment       =   1  'Right Justify
         Caption         =   "Normals"
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   3900
      TabIndex        =   11
      Text            =   "cmbHospital"
      Top             =   300
      Width           =   1455
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   2100
      TabIndex        =   9
      Text            =   "cmbCategory"
      Top             =   300
      Width           =   1575
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      HelpContextID   =   10010
      Left            =   330
      TabIndex        =   6
      Text            =   "cmbSampleType"
      Top             =   300
      Width           =   1575
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   855
      HelpContextID   =   10130
      Left            =   8220
      Picture         =   "frmBioMultiNormal.frx":04CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "bprint"
      Top             =   6330
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   855
      HelpContextID   =   10026
      Left            =   10290
      Picture         =   "frmBioMultiNormal.frx":0B34
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6330
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   855
      HelpContextID   =   10120
      Left            =   6690
      Picture         =   "frmBioMultiNormal.frx":119E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6330
      Width           =   1455
   End
   Begin VB.CommandButton cmdListAll 
      Caption         =   "&List All Analytes"
      Height          =   855
      HelpContextID   =   10110
      Left            =   3270
      Picture         =   "frmBioMultiNormal.frx":1808
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6330
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddAgeSpecific 
      Caption         =   "&Add Parameter Age Specific Range"
      Height          =   855
      HelpContextID   =   10100
      Left            =   330
      Picture         =   "frmBioMultiNormal.frx":1C4A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6330
      Visible         =   0   'False
      Width           =   2685
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   5385
      HelpContextID   =   10090
      Left            =   330
      TabIndex        =   0
      Top             =   720
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9499
      _Version        =   393216
      Cols            =   25
      FixedCols       =   4
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
      FormatString    =   $"frmBioMultiNormal.frx":208C
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   3930
      TabIndex        =   10
      Top             =   90
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   90
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Type"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   90
      Width           =   930
   End
End
Attribute VB_Name = "frmBioMultiNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ShowAllAgeRanges As String


Private Sub FillSampleTypes()

Dim sql As String
Dim tb As Recordset

sql = "Select Text from Lists where " & _
      "ListType = 'ST' " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql

cmbSampleType.Clear
Do While Not tb.EOF

  cmbSampleType.AddItem tb!Text & ""
  
  tb.MoveNext
  
Loop

End Sub
Private Sub FillHospitals()

Dim sql As String
Dim tb As Recordset

sql = "Select Text from Lists where " & _
      "ListType = 'HO' " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql

cmbHospital.Clear
Do While Not tb.EOF

  cmbHospital.AddItem tb!Text & ""
  
  tb.MoveNext
  
Loop

End Sub
Private Sub FillCategories()

Dim sql As String
Dim tb As Recordset

sql = "Select Cat from Categorys " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql

cmbCategory.Clear
Do While Not tb.EOF

  cmbCategory.AddItem tb!Cat & ""
  
  tb.MoveNext
  
Loop

End Sub


Private Sub cmbCategory_Click()

FillG

End Sub

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbHospital_Click()

FillG

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbSampleType_Click()

FillG

End Sub

Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub Form_Activate()

Dim intN As Integer

If optLongOrShort(0) Then
  gBio.ColWidth(0) = 1500
  gBio.ColWidth(1) = 0
Else
  gBio.ColWidth(0) = 0
  gBio.ColWidth(1) = 1500
End If
' "<Long Name             |<Short Name   |^Age From |^Age To   "               '0 to 3
' "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
' "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
' "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
' "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
' "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
' "|<RowNumber|<ModifiedFlag"                                                  '24 to 25

If optRange(0) Then
  For intN = 4 To 7
    gBio.ColWidth(intN) = 720
  Next
  For intN = 8 To 15
    gBio.ColWidth(intN) = 0
  Next
Else
  For intN = 4 To 7
    gBio.ColWidth(intN) = 0
  Next
  For intN = 8 To 11
    gBio.ColWidth(intN) = 720
  Next
  For intN = 12 To 15
    gBio.ColWidth(intN) = 0
  Next
End If

For intN = 19 To 25
  gBio.ColWidth(intN) = 0
Next

End Sub

Private Sub Form_Load()

Dim intN As Integer
Dim strS As String
Dim Discipline As String

FillSampleTypes
FillCategories
FillHospitals

strS = "<Long Name             |<Short Name   |^Age From |^Age To   "                      '0 to 3
strS = strS & "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
strS = strS & "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
strS = strS & "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
strS = strS & "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
strS = strS & "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
strS = strS & "|<RowNumber|<ModifiedFlag"                                                  '24 to 25
gBio.FormatString = strS

gBio.RowHeight(0) = 660
For intN = 4 To 11
  gBio.ColWidth(intN) = 720
Next
'For intN = 19 To 25
'  gBio.ColWidth(intN) = 0
'Next

For intN = 0 To 2
  If optDiscipline(intN) Then
    Discipline = Left$(optDiscipline(intN).Caption, 3)
  End If
Next
CheckDisciplineActive Discipline

FillG

End Sub

Private Sub gbio_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim s As String

If cmdListAll.Visible Then
  iMsg "Historical View is not updateable.", vbExclamation
  Exit Sub
End If

If gBio.MouseRow = 0 Then
  cmdAddAgeSpecific.Visible = False
  Exit Sub
End If

cmdAddAgeSpecific.Caption = "&Add " & gBio.TextMatrix(gBio.Row, 0) & " Age Specific Range"
cmdAddAgeSpecific.Visible = True
'strS = "<Long Name             |<Short Name   |^Age From |^Age To   "                      '0 to 3
'strS = strS & "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
'strS = strS & "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
'strS = strS & "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
'strS = strS & "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
'strS = strS & "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
'strS = strS & "|<RowNumber|<ModifiedFlag"                                                  '24 to 25

Select Case gBio.MouseCol
    
  Case 0
    gBio.Col = 0
    If gBio.CellPicture = imgSquarePlus.Picture Then
      ShowAllAgeRanges = gBio.TextMatrix(gBio.MouseRow, 0)
    Else
      ShowAllAgeRanges = ""
    End If
    FillG gBio.TopRow
    
  Case 2, 3
    iMsg "Ages are not Editable!", vbExclamation
  
  Case 4 To 15
    s = "Enter " & gBio.TextMatrix(0, gBio.Col) & " Range for " & gBio.TextMatrix(gBio.Row, 0)
    gBio = iBOX(s, , gBio.TextMatrix(gBio.Row, gBio.Col))
    gBio.TextMatrix(gBio.Row, 17) = Format$(Now, "dd/mm/yyyy")
    cmdSave.Enabled = True
    gBio.TextMatrix(gBio.Row, 25) = "Yes"
  Case 16
    If cmdSave.Enabled Then
      iMsg "Details must be saved before viewing history.", vbExclamation
    Else
      FillGHistory gBio.TextMatrix(gBio.Row, 0)
    End If
    
  Case 17, 18
    iMsg "Active From/To Dates are not Editable!", vbExclamation
  
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdSave.Enabled Then
  If iMsg("Exit without Saving?", vbQuestion + vbYesNo) = vbNo Then
    Cancel = True
  End If
End If

End Sub



Private Sub cmdsave_Click()

SaveDetails

End Sub


Private Sub cmdListAll_Click()

FillG

End Sub


Private Sub cmdcancel_Click()

Unload Me

End Sub



Private Sub cmdAddAgeSpecific_Click()
    
Dim n As Integer
Dim s As String
Dim strSampleType As String

If cmbHospital = "" Then
  cmbHospital = HospName(0)
End If
If cmbCategory = "" Then
  cmbCategory = "Human"
End If
If cmbSampleType = "" Then
  cmbSampleType = "Serum"
End If
strSampleType = ListCodeFor("ST", cmbSampleType)

'Caption = "&Add PT Age Specific Range"
n = InStr(cmdAddAgeSpecific.Caption, " Age")
s = Mid$(cmdAddAgeSpecific.Caption, 6, n - 6)

With frmBioMultiAddAge
  .Analyte = s
  .SampleType = strSampleType
  .Hospital = cmbHospital
  .Category = cmbCategory
  .Show 1
End With
FillG

End Sub



Private Sub bPrint_Click()

Dim n As Integer
Dim s As String
Dim lngDP As Long

Printer.Print "Biochemistry Normal Ranges"
Printer.Print
Printer.Print
Printer.Print "Parameter   Male Low High    Female Low High  Plausible Low High  Units"
Printer.Print
'0 Analyte
'1 Age From
'2 Age To
'3 MaleLow
'4 MaleHigh
'5 FemaleLow
'6 FemaleHigh
'7 FlagMaleLow
'8 FlagMaleHigh
'9 FlagFemaleLow
'10 FlagFemaleHigh
'11 PlausibleLow
'12 PlausibleHigh
'13 Code
'14 Units
'15 Dec.Pl
'16 Active From
'17 Active To
'18 Displayable
'19 Printable
'20 Delta %
'21 Previous

For n = 1 To gBio.Rows - 1

  lngDP = gBio.TextMatrix(n, 15)
  
  s = Left$(gBio.TextMatrix(n, 0) & Space$(12), 12)
  s = s & Left$(FormatNumber(gBio.TextMatrix(n, 3), lngDP) & Space$(9), 9)
  s = s & Left$(FormatNumber(gBio.TextMatrix(n, 4), lngDP) & Space$(9), 9)
  s = s & Left$(FormatNumber(gBio.TextMatrix(n, 5), lngDP) & Space$(9), 9)
  s = s & Left$(FormatNumber(gBio.TextMatrix(n, 6), lngDP) & Space$(9), 9)
  s = s & Left$(FormatNumber(gBio.TextMatrix(n, 11), lngDP) & Space$(9), 9)
  s = s & Left$(FormatNumber(gBio.TextMatrix(n, 12), lngDP) & Space$(9), 9)
  s = s & gBio.TextMatrix(n, 14)
      
  Printer.Print s
  
Next

Printer.EndDoc

End Sub



Private Sub SaveDetails()

Dim intN As Integer
Dim tb As Recordset
Dim tbOrig As Recordset
Dim sql As String
Dim strSampleType As String
Dim Discipline As String

If cmbHospital = "" Then
  cmbHospital = HospName(0)
End If
If cmbCategory = "" Then
  cmbCategory = "Human"
End If
If cmbSampleType = "" Then
  cmbSampleType = "Serum"
End If
strSampleType = ListCodeFor("ST", cmbSampleType)

For intN = 0 To 2
  If optDiscipline(intN) Then
    Discipline = Left$(optDiscipline(intN).Caption, 3)
  End If
Next

'strS = "<Long Name             |<Short Name   |^Age From |^Age To   "                      '0 to 3
'strS = strS & "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
'strS = strS & "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
'strS = strS & "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
'strS = strS & "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
'strS = strS & "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
'strS = strS & "|<RowNumber|<ModifiedFlag"                                                  '24 to 25


For intN = 1 To gBio.Rows - 1
  If gBio.TextMatrix(intN, 25) = "Yes" Then 'row has been modified
    'set ActiveToDate to yesterday
    sql = "Select * from " & Discipline & "TestDefinitions where " & _
          "LongName = '" & gBio.TextMatrix(intN, 0) & "' " & _
          "and SampleType = '" & strSampleType & "' " & _
          "and Category = '" & cmbCategory & "' " & _
          "and Hospital = '" & cmbHospital & "' " & _
          "and AgeFromDays = '" & gBio.TextMatrix(intN, 21) & "' " & _
          "and ActiveToDate = '" & Format$(gBio.TextMatrix(intN, 18), "dd/mmm/yyyy") & "' " & _
          "order by ActiveToDate desc"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    Set tbOrig = New Recordset
    RecOpenServer 0, tbOrig, sql
    If Not tb.EOF Then
      tb!ActiveToDate = Format$(Now - 1, "dd/mmm/yyyy")
      tb.Update
    End If
    tb.AddNew
    tb!LongName = gBio.TextMatrix(intN, 0)
    tb!ShortName = gBio.TextMatrix(intN, 1)
    tb!MaleLow = Val(gBio.TextMatrix(intN, 4))
    tb!MaleHigh = Val(gBio.TextMatrix(intN, 5))
    tb!FemaleLow = Val(gBio.TextMatrix(intN, 6))
    tb!FemaleHigh = Val(gBio.TextMatrix(intN, 7))
    tb!FlagMaleLow = Val(gBio.TextMatrix(intN, 8))
    tb!FlagMaleHigh = Val(gBio.TextMatrix(intN, 9))
    tb!FlagFemaleLow = Val(gBio.TextMatrix(intN, 10))
    tb!FlagFemaleHigh = Val(gBio.TextMatrix(intN, 11))
    tb!PlausibleLow = Val(gBio.TextMatrix(intN, 12))
    tb!PlausibleHigh = Val(gBio.TextMatrix(intN, 13))
    tb!AutoValLow = Val(gBio.TextMatrix(intN, 14))
    tb!AutoValHigh = Val(gBio.TextMatrix(intN, 15))
    
    tb!Code = tbOrig!Code & ""
    tb!BarCode = tbOrig!BarCode & ""
    tb!ImmunoCode = tbOrig!ImmunoCode & ""
    tb!Units = tbOrig!Units & ""
    tb!DP = tbOrig!DP
    
    tb!ActiveFromDate = Format$(Now, "dd/mmm/yyyy")
    tb!ActiveToDate = Format$(Now, "dd/mmm/yyyy")
    
    tb!AgeFromDays = gBio.TextMatrix(intN, 21)
    tb!AgeToDays = gBio.TextMatrix(intN, 22)
    tb!PrintPriority = tbOrig!PrintPriority
      
    tb!KnownToAnalyser = tbOrig!KnownToAnalyser
    tb!Analyser = tbOrig!Analyser
    tb!InUse = tbOrig!InUse
    tb!SplitList = tbOrig!SplitList
    tb!EOD = tbOrig!EOD
    
    tb!Hospital = cmbHospital
    tb!Category = cmbCategory
    tb!SampleType = strSampleType
    
    tb!h = tbOrig!h
    tb!s = tbOrig!s
    tb!l = tbOrig!l
    tb!o = tbOrig!o
    tb!g = tbOrig!g
    tb!J = tbOrig!J
    
    tb!DoDelta = tbOrig!DoDelta
    tb!DeltaLimit = tbOrig!DeltaLimit
    
    tb!Printable = tbOrig!Printable
    tb.Update
  End If
Next

cmdSave.Enabled = False

CheckDisciplineActive Discipline
FillG

End Sub
Private Sub FillGHistory(ByVal TestName As String)

Dim tb As Recordset
Dim sql As String
Dim s As String
Dim intN As Integer
Dim Discipline As String

gBio.Rows = 2
gBio.AddItem ""
gBio.RemoveItem 1

For intN = 0 To 2
  If optDiscipline(intN) Then
    Discipline = Left$(optDiscipline(intN).Caption, 3)
  End If
Next
' "<Long Name             |<Short Name   |^Age From |^Age To   "               '0 to 3
' "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
' "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
' "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
' "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
' "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
' "|<RowNumber|<ModifiedFlag"                                                  '24 to 25

sql = "Select * from " & Discipline & "TestDefinitions where " & _
      "LongName = '" & TestName & "' " & _
      "order by ActiveToDate desc, AgeFromDays asc"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  s = tb!LongName & vbTab & tb!ShortName & vbTab & _
      vbTab & vbTab & _
      tb!MaleLow & vbTab & tb!MaleHigh & vbTab & _
      tb!FemaleLow & vbTab & tb!FemaleHigh & vbTab & _
      tb!FlagMaleLow & vbTab & tb!FlagMaleHigh & vbTab & _
      tb!FlagFemaleLow & vbTab & tb!FlagFemaleHigh & vbTab & _
      tb!PlausibleLow & vbTab & tb!PlausibleHigh & vbTab & _
      tb!AutoValLow & vbTab & tb!AutoValHigh & vbTab & _
      vbTab & tb!ActiveFromDate & vbTab & tb!ActiveToDate & vbTab & _
      tb!Code & vbTab & tb!DP & vbTab & _
      tb!AgeFromDays & vbTab & _
      tb!AgeToDays & vbTab & _
      tb!PrintPriority
  gBio.AddItem s
  tb.MoveNext
Loop
If gBio.Rows > 2 Then gBio.RemoveItem 1

AdjustAgeView

cmdListAll.Visible = True

End Sub

Private Sub AdjustAgeView()

Dim intN As Integer
' "<Long Name             |<Short Name   |^Age From |^Age To   "               '0 to 3
' "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
' "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
' "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
' "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
' "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
' "|<RowNumber|<ModifiedFlag"                                                  '24 to 25

For intN = 1 To gBio.Rows - 1
  If IsNumeric(gBio.TextMatrix(intN, 21)) Then
    gBio.TextMatrix(intN, 2) = dmyFromCount(gBio.TextMatrix(intN, 21))
  End If
  If IsNumeric(gBio.TextMatrix(intN, 22)) Then
    gBio.TextMatrix(intN, 3) = dmyFromCount(gBio.TextMatrix(intN, 22))
  End If
Next
  
End Sub





Private Sub FillG(Optional ByVal RowTop As Long)

Dim sql As String
Dim tb As Recordset
Dim strS As String
Dim intN As Integer
Dim blnFound As Boolean
Dim blnShowThis As Boolean
Dim strSampleType As String
Dim Discipline As String

If cmbHospital = "" Then
  cmbHospital = HospName(0)
End If
If cmbCategory = "" Then
  cmbCategory = "Human"
End If
If cmbSampleType = "" Then
  cmbSampleType = "Serum"
End If
strSampleType = ListCodeFor("ST", cmbSampleType)

For intN = 0 To 2
  If optDiscipline(intN) Then
    Discipline = Left$(optDiscipline(intN).Caption, 3)
  End If
Next
' "<Long Name             |<Short Name   |^Age From |^Age To   "               '0 to 3
' "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
' "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
' "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
' "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
' "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
' "|<RowNumber|<ModifiedFlag"                                                  '24 to 25

gBio.Visible = False
gBio.Rows = 2
gBio.AddItem ""
gBio.RemoveItem 1

sql = "Select * from " & Discipline & "TestDefinitions where " & _
      "SampleType = '" & strSampleType & "' " & _
      "and Category = '" & cmbCategory & "' " & _
      "and Hospital = '" & cmbHospital & "' " & _
      "order by PrintPriority , ActiveToDate desc, AgeFromDays"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  blnShowThis = False
  blnFound = False
  For intN = 1 To gBio.Rows - 1
    'Is there historical record?
    If gBio.TextMatrix(intN, 0) = tb!LongName And Val(gBio.TextMatrix(intN, 21)) = tb!AgeFromDays Then
      blnFound = True
      blnShowThis = False
      gBio.Col = 16 'Previous
      Set gBio.CellPicture = imgSquareTick.Picture
      gBio.CellPictureAlignment = flexAlignCenterCenter
      Exit For
    End If
  Next
  
  If Not blnFound Then
    If ShowAllAgeRanges <> tb!LongName Then  'is other ages
      For intN = 1 To gBio.Rows - 1
        If gBio.TextMatrix(intN, 0) = tb!LongName Then
          blnFound = True
          gBio.Col = 0
          gBio.Row = intN
          Set gBio.CellPicture = imgSquarePlus.Picture
          gBio.CellPictureAlignment = flexAlignRightCenter
          Exit For
        End If
      Next
    Else
      blnShowThis = False
      For intN = 1 To gBio.Rows - 1
        If gBio.TextMatrix(intN, 0) = tb!LongName Then
          blnFound = True
          If ShowAllAgeRanges = tb!LongName Then
            blnShowThis = True
            Exit For
          End If
        End If
      Next
    End If
  End If
  If Not blnFound Or blnShowThis Then
    strS = tb!LongName & vbTab & tb!ShortName & vbTab & _
           vbTab & vbTab & _
           tb!MaleLow & vbTab & tb!MaleHigh & vbTab & _
           tb!FemaleLow & vbTab & tb!FemaleHigh & vbTab & _
           tb!FlagMaleLow & vbTab & tb!FlagMaleHigh & vbTab & _
           tb!FlagFemaleLow & vbTab & tb!FlagFemaleHigh & vbTab & _
           tb!PlausibleLow & vbTab & tb!PlausibleHigh & vbTab & _
           tb!AutoValLow & vbTab & tb!AutoValHigh & vbTab & vbTab
    If tb!ActiveFromDate <> "01/01/1990" Then
      strS = strS & tb!ActiveFromDate
    End If
    strS = strS & vbTab & tb!ActiveToDate & vbTab & _
           tb!Code & vbTab & tb!DP & vbTab & _
           tb!AgeFromDays & vbTab & _
           tb!AgeToDays & vbTab & _
           tb!PrintPriority
    
    gBio.AddItem strS
    gBio.Row = gBio.Rows - 1
  End If
  tb.MoveNext
Loop

If ShowAllAgeRanges <> "" Then
  gBio.Col = 0
  For intN = 1 To gBio.Rows - 1
    If gBio.TextMatrix(intN, 0) = ShowAllAgeRanges Then
      gBio.Row = intN
      Set gBio.CellPicture = imgSquareMinus.Picture
      gBio.CellPictureAlignment = flexAlignRightCenter
    End If
  Next
End If

If RowTop <> 0 Then
  gBio.TopRow = RowTop
End If

If gBio.Rows > 2 Then gBio.RemoveItem 1
gBio.Visible = True

AdjustAgeView

cmdListAll.Visible = False

End Sub

Private Sub optDiscipline_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

FillG

End Sub


Private Sub optLongOrShort_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

If optLongOrShort(0) Then
  gBio.ColWidth(0) = 1500
  gBio.ColWidth(1) = 0
Else
  gBio.ColWidth(0) = 0
  gBio.ColWidth(1) = 1500
End If

End Sub

Private Sub optRange_Click(Index As Integer)

Dim intN As Integer

'strS = "<Long Name             |<Short Name   |^Age From |^Age To   "                      '0 to 3
'strS = strS & "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
'strS = strS & "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
'strS = strS & "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
'strS = strS & "|^Previous |<Active From      |<Active To       |<Code    |^Dec.Pl  "       '16 to 20
'strS = strS & "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
'strS = strS & "|<RowNumber|<ModifiedFlag"

Select Case Index
  
  Case 0:
    For intN = 4 To 7
      gBio.ColWidth(intN) = 720
    Next
    For intN = 8 To 15
      gBio.ColWidth(intN) = 0
    Next
  
  Case 1:
    For intN = 4 To 7
      gBio.ColWidth(intN) = 0
    Next
    For intN = 8 To 11
      gBio.ColWidth(intN) = 720
    Next
    For intN = 12 To 15
      gBio.ColWidth(intN) = 0
    Next

End Select

End Sub


