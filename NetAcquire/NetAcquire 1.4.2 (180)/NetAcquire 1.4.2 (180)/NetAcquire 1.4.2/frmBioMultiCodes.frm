VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioMultiCodes 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NetAcquire - Biochemistry Codes, Units & Precision"
   ClientHeight    =   5865
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   10875
   HelpContextID   =   10017
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   885
      Left            =   9030
      TabIndex        =   13
      Top             =   30
      Width           =   1575
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   420
         Width           =   1275
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   630
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Names"
      Height          =   555
      Left            =   5460
      TabIndex        =   10
      Top             =   180
      Width           =   2685
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Short Names"
         Height          =   255
         Index           =   1
         Left            =   1350
         TabIndex        =   12
         Top             =   210
         Width           =   1245
      End
      Begin VB.OptionButton optLongOrShort 
         Alignment       =   1  'Right Justify
         Caption         =   "Long Names"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   3900
      TabIndex        =   9
      Text            =   "cmbHospital"
      Top             =   390
      Width           =   1455
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   2100
      TabIndex        =   7
      Text            =   "cmbCategory"
      Top             =   390
      Width           =   1575
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      HelpContextID   =   10010
      Left            =   330
      TabIndex        =   4
      Text            =   "cmbSampleType"
      Top             =   390
      Width           =   1575
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   645
      HelpContextID   =   10130
      Left            =   7620
      Picture         =   "frmBioMultiCodes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "bprint"
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      HelpContextID   =   10026
      Left            =   9210
      Picture         =   "frmBioMultiCodes.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   645
      HelpContextID   =   10120
      Left            =   6060
      Picture         =   "frmBioMultiCodes.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   4065
      HelpContextID   =   10090
      Left            =   330
      TabIndex        =   0
      Top             =   960
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   7170
      _Version        =   393216
      Cols            =   12
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
      FormatString    =   $"frmBioMultiCodes.frx":133E
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   90
      Picture         =   "frmBioMultiCodes.frx":13D7
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   90
      Picture         =   "frmBioMultiCodes.frx":16AD
      Top             =   420
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   3930
      TabIndex        =   8
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   180
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Type"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   180
      Width           =   930
   End
End
Attribute VB_Name = "frmBioMultiCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strUnits() As String

Private Sub FillStrUnits()

Dim sql As String
Dim tb As Recordset
Dim intN As Integer

sql = "Select Text from Lists where " & _
      "ListType = 'UN' " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql

ReDim strUnits(0 To 0)
strUnits(0) = ""
intN = 1

Do While Not tb.EOF

  ReDim Preserve strUnits(0 To intN) As String
  strUnits(intN) = tb!Text & ""
  
  intN = intN + 1
  
  tb.MoveNext
  
Loop

End Sub

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

LoadDefinitions

End Sub

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbHospital_Click()

LoadDefinitions

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbSampleType_Click()

LoadDefinitions

End Sub

Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub Form_Activate()

If optLongOrShort(0) Then
  gBio.ColWidth(0) = 1500
  gBio.ColWidth(1) = 0
Else
  gBio.ColWidth(0) = 0
  gBio.ColWidth(1) = 1500
End If

End Sub

Private Sub Form_Load()

Dim intN As Integer
Dim Discipline As String

FillStrUnits
FillSampleTypes
FillCategories
FillHospitals

For intN = 0 To 2
  If optDiscipline(intN) Then
    Discipline = Left$(optDiscipline(intN).Caption, 3)
  End If
Next
CheckDisciplineActive Discipline

LoadDefinitions

End Sub

Private Sub gbio_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

Dim s As String
Dim F As Form

'<Long Name  |<Short Name                         0 to 1
'^Code |^Bar Code |^Immuno Code|^Units  |^Dec.Pl  2 to 6
'^Printable|^Known to Analyser|^Analyser Code|    7 to 9
'^In Use|^End of Day                              10 to 11
Select Case gBio.MouseCol
    
  Case 2
    iMsg "Code is not Editable!", vbExclamation
  
  Case 3
    gBio.TextMatrix(gBio.Row, 3) = iBOX("Scan Barcode for " & gBio.TextMatrix(gBio.Row, 0), , gBio.TextMatrix(gBio.Row, 3))
    cmdSave.Enabled = True

  Case 4, 9
    s = "Enter " & gBio.TextMatrix(0, gBio.Col) & " for " & gBio.TextMatrix(gBio.Row, 0)
    gBio = iBOX(s, , gBio.TextMatrix(gBio.Row, gBio.Col))
    cmdSave.Enabled = True
  
  Case 5
    Set F = New fcdrDBox
    With F
      .Options = strUnits
      .Prompt = "Enter Units for " & gBio.TextMatrix(gBio.Row, 0)
      .Show 1
      gBio = .ReturnValue
    End With
    Unload F
    Set F = Nothing
    cmdSave.Enabled = True
  
  Case 6
    Select Case gBio.TextMatrix(gBio.Row, 6)
      Case "0": gBio.TextMatrix(gBio.Row, 6) = "1"
      Case "1": gBio.TextMatrix(gBio.Row, 6) = "2"
      Case "2": gBio.TextMatrix(gBio.Row, 6) = "3"
      Case Else: gBio.TextMatrix(gBio.Row, 6) = "0"
    End Select
    cmdSave.Enabled = True
  
  
  Case 7, 8, 10, 11
    If gBio.CellPicture = imgSquareCross.Picture Then
      Set gBio.CellPicture = imgSquareTick.Picture
    Else
      Set gBio.CellPicture = imgSquareCross.Picture
    End If
    cmdSave.Enabled = True
  
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


Private Sub cmdCancel_Click()

Unload Me

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
Dim sql As String
Dim strSampleType As String
Dim blnP As Boolean
Dim blnK As Boolean
Dim blnI As Boolean
Dim blnE As Boolean
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

For intN = 1 To gBio.Rows - 1
  gBio.Row = intN
  
  gBio.Col = 7
  If gBio.CellPicture = imgSquareTick.Picture Then blnP = True Else blnP = False
  
  gBio.Col = 8
  If gBio.CellPicture = imgSquareTick.Picture Then blnK = True Else blnK = False
  
  gBio.Col = 10
  If gBio.CellPicture = imgSquareTick.Picture Then blnI = True Else blnI = False
  
  gBio.Col = 11
  If gBio.CellPicture = imgSquareTick.Picture Then blnE = True Else blnE = False

'<Long Name  |<Short Name                         0 to 1
'^Code |^Bar Code |^Immuno Code|^Units  |^Dec.Pl  2 to 6
'^Printable|^Known to Analyser|^Analyser Code|    7 to 9
'^In Use|^End of Day                              10 to 11
    
  sql = "Update " & Discipline & "TestDefinitions " & _
        "Set Code = '" & gBio.TextMatrix(intN, 2) & "', " & _
        "BarCode = '" & gBio.TextMatrix(intN, 3) & "', " & _
        "ImmunoCode = '" & gBio.TextMatrix(intN, 4) & "', " & _
        "Units = '" & gBio.TextMatrix(intN, 5) & "', " & _
        "DP = '" & Val(gBio.TextMatrix(intN, 6)) & "', " & _
        "Printable = " & IIf(blnP, 1, 0) & ", " & _
        "KnownToAnalyser = " & IIf(blnK, 1, 0) & ", " & _
        "Analyser = '" & gBio.TextMatrix(intN, 9) & "', " & _
        "InUse = " & IIf(blnI, 1, 0) & ", " & _
        "EOD = " & IIf(blnE, 1, 0) & " " & _
        "where LongName = '" & gBio.TextMatrix(intN, 0) & "' " & _
        "and SampleType = '" & strSampleType & "' " & _
        "and Category = '" & cmbCategory & "' " & _
        "and Hospital = '" & cmbHospital & "'"
  Cnxn(0).Execute sql
Next

cmdSave.Enabled = False

CheckDisciplineActive Discipline
LoadDefinitions

End Sub
Private Sub LoadDefinitions()

Dim tb As Recordset
Dim sql As String
Dim strS As String
Dim strSampleType As String
Dim blnFound As Boolean
Dim intN As Integer
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

gBio.Visible = False
gBio.Rows = 2
gBio.AddItem ""
gBio.RemoveItem 1
'<Long Name  |<Short Name                         0 to 1
'^Code |^Bar Code |^Immuno Code|^Units  |^Dec.Pl  2 to 6
'^Printable|^Known to Analyser|^Analyser Code|    7 to 9
'^In Use|^End of Day                              10 to 11

sql = "Select * from " & Discipline & "TestDefinitions where " & _
      "SampleType = '" & strSampleType & "' " & _
      "and Category = '" & cmbCategory & "' " & _
      "and Hospital = '" & cmbHospital & "' " & _
      "order by PrintPriority"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  blnFound = False
  For intN = 1 To gBio.Rows - 1
    If gBio.TextMatrix(intN, 0) = tb!LongName & "" Then
      blnFound = True
      Exit For
    End If
  Next
  If Not blnFound Then
    strS = tb!LongName & vbTab & _
           tb!ShortName & vbTab & _
           tb!Code & vbTab & _
           tb!BarCode & vbTab & _
           tb!ImmunoCode & vbTab & _
           tb!Units & vbTab & _
           tb!DP & vbTab & _
           vbTab & vbTab & _
           Trim$(tb!Analyser & "")
    
    gBio.AddItem strS
    
    gBio.Row = gBio.Rows - 1
    gBio.Col = 7 'Printable
    If Not IsNull(tb!Printable) Then
      Set gBio.CellPicture = IIf(tb!Printable, imgSquareTick.Picture, imgSquareCross.Picture)
    Else
      Set gBio.CellPicture = imgSquareCross.Picture
    End If
    gBio.CellPictureAlignment = flexAlignCenterCenter
     
    gBio.Col = 8 'Known to Analyser
    If Not IsNull(tb!KnownToAnalyser) Then
      Set gBio.CellPicture = IIf(tb!KnownToAnalyser, imgSquareTick.Picture, imgSquareCross.Picture)
    Else
      Set gBio.CellPicture = imgSquareCross.Picture
    End If
    gBio.CellPictureAlignment = flexAlignCenterCenter
     
    gBio.Col = 10 'In Use
    If Not IsNull(tb!InUse) Then
      Set gBio.CellPicture = IIf(tb!InUse, imgSquareTick.Picture, imgSquareCross.Picture)
    Else
      Set gBio.CellPicture = imgSquareCross.Picture
    End If
    gBio.CellPictureAlignment = flexAlignCenterCenter
     
    gBio.Col = 11 'End of Day
    If Not IsNull(tb!EOD) Then
      Set gBio.CellPicture = IIf(tb!EOD, imgSquareTick.Picture, imgSquareCross.Picture)
    Else
      Set gBio.CellPicture = imgSquareCross.Picture
    End If
    gBio.CellPictureAlignment = flexAlignCenterCenter
     
  End If
  
  tb.MoveNext
Loop

If gBio.Rows > 2 Then gBio.RemoveItem 1
gBio.Visible = True

End Sub


Private Sub optDiscipline_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

LoadDefinitions

End Sub


Private Sub optLongOrShort_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

If optLongOrShort(0) Then
  gBio.ColWidth(0) = 1500
  gBio.ColWidth(1) = 0
Else
  gBio.ColWidth(0) = 0
  gBio.ColWidth(1) = 1500
End If

End Sub


