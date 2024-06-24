VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioMultiMasks 
   Caption         =   "NetAcquire - Biochemistry Masks"
   ClientHeight    =   6300
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   7020
   HelpContextID   =   10017
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7020
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   885
      Left            =   5220
      TabIndex        =   10
      Top             =   30
      Width           =   1575
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   420
         Width           =   1275
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   11
         Top             =   630
         Width           =   1305
      End
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   3570
      TabIndex        =   9
      Text            =   "cmbHospital"
      Top             =   390
      Width           =   1455
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   1950
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
      Left            =   3780
      Picture         =   "frmBioMultiMasks.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "bprint"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      HelpContextID   =   10026
      Left            =   5340
      Picture         =   "frmBioMultiMasks.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   645
      HelpContextID   =   10120
      Left            =   2220
      Picture         =   "frmBioMultiMasks.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   4005
      HelpContextID   =   10090
      Left            =   240
      TabIndex        =   0
      Top             =   1380
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   7064
      _Version        =   393216
      Cols            =   6
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Long Name                 |<Short Name      |^Old    |^  Lipaemic  |^  Icteric  |^  Haemolysed  "
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Do not Print Result if >="
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   3180
      TabIndex        =   14
      Top             =   1080
      Width           =   3270
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   90
      Picture         =   "frmBioMultiMasks.frx":133E
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   90
      Picture         =   "frmBioMultiMasks.frx":1614
      Top             =   510
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   3600
      TabIndex        =   8
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   2010
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
Attribute VB_Name = "frmBioMultiMasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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



Private Sub Form_Load()

Dim intN As Integer
Dim strDiscipline As String

FillSampleTypes
FillCategories
FillHospitals

For intN = 0 To 2
  If optDiscipline(intN) Then
    strDiscipline = Left$(optDiscipline(intN).Caption, 3)
  End If
Next
CheckDisciplineActive strDiscipline

LoadDefinitions

End Sub

Private Sub gbio_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

Select Case gBio.MouseCol
    
  Case 2 'Old
    If gBio.CellPicture = imgSquareCross.Picture Then
      Set gBio.CellPicture = imgSquareTick.Picture
    Else
      Set gBio.CellPicture = imgSquareCross.Picture
    End If
    cmdSave.Enabled = True

  Case 3, 4, 5 'LIH
    Select Case gBio.TextMatrix(gBio.Row, gBio.Col)
      Case "": gBio.TextMatrix(gBio.Row, gBio.Col) = "1+"
      Case "1+": gBio.TextMatrix(gBio.Row, gBio.Col) = "2+"
      Case "2+": gBio.TextMatrix(gBio.Row, gBio.Col) = "3+"
      Case "3+": gBio.TextMatrix(gBio.Row, gBio.Col) = "4+"
      Case "4+": gBio.TextMatrix(gBio.Row, gBio.Col) = "5+"
      Case Else: gBio.TextMatrix(gBio.Row, gBio.Col) = ""
    End Select
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
Dim intLIH As Integer
Dim intO As Integer
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

gBio.Col = 2 'Old
For intN = 1 To gBio.Rows - 1
  gBio.Row = intN
  intLIH = (Val(gBio.TextMatrix(intN, 3)) * 100) + _
           (Val(gBio.TextMatrix(intN, 4)) * 10) + _
           (Val(gBio.TextMatrix(intN, 5)))
  intO = IIf(gBio.CellPicture = imgSquareTick, 1, 0)

  sql = "Update " & Discipline & "TestDefinitions SET " & _
        "LIH = '" & intLIH & "', " & _
        "O = '" & intO & "' " & _
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
Dim intLIH As Integer
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
CheckDisciplineActive Discipline

gBio.Visible = False
gBio.Rows = 2
gBio.AddItem ""
gBio.RemoveItem 1
gBio.Col = 2 'Old

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
           tb!ShortName & vbTab & vbTab
    'LIH
    'L 100 to 600
    'I  10 to  60
    'H   1 to   6
    intLIH = IIf(Not IsNull(tb!LIH), tb!LIH, 0)
    If intLIH > 99 Then
      strS = strS & Format$(intLIH \ 100) & "+"
      intLIH = intLIH Mod 100
    End If
    strS = strS & vbTab
    
    If intLIH > 9 Then
      strS = strS & Format$(intLIH \ 10) & "+"
      intLIH = intLIH Mod 10
    End If
    strS = strS & vbTab
    If intLIH > 0 Then
      strS = strS & Format$(intLIH) & "+"
    End If
    
    gBio.AddItem strS
    
    gBio.Row = gBio.Rows - 1
    If Not IsNull(tb!o) Then
      Set gBio.CellPicture = IIf(tb!o, imgSquareTick.Picture, imgSquareCross.Picture)
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


