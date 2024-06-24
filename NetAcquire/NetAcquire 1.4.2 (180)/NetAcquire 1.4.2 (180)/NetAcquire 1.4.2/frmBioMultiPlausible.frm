VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioMultiPlausible 
   Caption         =   "NetAcquire - Biochemistry Plausible, AutoVal & Delta Ranges"
   ClientHeight    =   5880
   ClientLeft      =   330
   ClientTop       =   495
   ClientWidth     =   9120
   HelpContextID   =   10017
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9120
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   885
      Left            =   7260
      TabIndex        =   13
      Top             =   60
      Width           =   1575
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   630
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
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Names"
      Height          =   885
      Left            =   5550
      TabIndex        =   10
      Top             =   60
      Width           =   1425
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Long Names"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Short Names"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   510
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
      Left            =   5790
      Picture         =   "frmBioMultiPlausible.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "bprint"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      HelpContextID   =   10026
      Left            =   7380
      Picture         =   "frmBioMultiPlausible.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   645
      HelpContextID   =   10120
      Left            =   4200
      Picture         =   "frmBioMultiPlausible.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   4005
      HelpContextID   =   10090
      Left            =   300
      TabIndex        =   0
      Top             =   1020
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7064
      _Version        =   393216
      Cols            =   8
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
      FormatString    =   $"frmBioMultiPlausible.frx":133E
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   60
      Picture         =   "frmBioMultiPlausible.frx":13D4
      Top             =   150
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   60
      Picture         =   "frmBioMultiPlausible.frx":16AA
      Top             =   450
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
Attribute VB_Name = "frmBioMultiPlausible"
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

Select Case gBio.MouseCol
    
  Case 2 To 5, 7
    s = "Enter " & gBio.TextMatrix(0, gBio.Col) & " Range for " & gBio.TextMatrix(gBio.Row, 0)
    gBio = iBOX(s, , gBio.TextMatrix(gBio.Row, gBio.Col))
    cmdSave.Enabled = True
  
  Case 6 'DoDelta
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

Printer.Print "Biochemistry Plausible, AutoVal and Delta Ranges"
Printer.Print
Printer.Print
Printer.Print "Parameter        Plausible            Auto-Val             Delta"
Printer.Print "                Low    High          Low    High         (Absolute)"
Printer.Print
gBio.Col = 6
For n = 1 To gBio.Rows - 1
  gBio.Row = n
  s = Left$(gBio.TextMatrix(n, 0) & Space$(15), 15) & " "
  s = s & Left$(gBio.TextMatrix(n, 2) & Space$(6), 6) & " "
  s = s & Left$(gBio.TextMatrix(n, 3) & Space$(6), 6) & "        "
  s = s & Left$(gBio.TextMatrix(n, 4) & Space$(6), 6) & " "
  s = s & Left$(gBio.TextMatrix(n, 5) & Space$(6), 6) & "        "
  If gBio.CellPicture = imgSquareTick.Picture Then
    s = s & gBio.TextMatrix(n, 7)
  Else
    s = s & "  -"
  End If
  
  Printer.Print s
  
Next

Printer.EndDoc

End Sub



Private Sub SaveDetails()

Dim intN As Integer
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

gBio.Col = 6
For intN = 1 To gBio.Rows - 1
  gBio.Row = intN
  sql = "Update " & Discipline & "TestDefinitions " & _
        "Set PlausibleLow = '" & Val(gBio.TextMatrix(intN, 2)) & "', " & _
        "PlausibleHigh = '" & Val(gBio.TextMatrix(intN, 3)) & "', " & _
        "AutoValLow = '" & Val(gBio.TextMatrix(intN, 4)) & "', " & _
        "AutoValHigh = '" & Val(gBio.TextMatrix(intN, 5)) & "', " & _
        "DeltaLimit = '" & Val(gBio.TextMatrix(intN, 7)) & "', " & _
        "DoDelta = '" & IIf(gBio.CellPicture = imgSquareTick.Picture, 1, 0) & "' " & _
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
Dim TableName As String

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
    TableName = Left$(optDiscipline(intN).Caption, 3)
  End If
Next

gBio.Visible = False
gBio.Rows = 2
gBio.AddItem ""
gBio.RemoveItem 1

sql = "Select * from " & TableName & "TestDefinitions where " & _
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
           tb!ShortName & vbTab
    If Not IsNull(tb!PlausibleLow) Then strS = strS & tb!PlausibleLow
    strS = strS & vbTab
    If Not IsNull(tb!PlausibleHigh) Then strS = strS & tb!PlausibleHigh
    strS = strS & vbTab
    If Not IsNull(tb!AutoValLow) Then strS = strS & tb!AutoValLow
    strS = strS & vbTab
    If Not IsNull(tb!AutoValHigh) Then strS = strS & tb!AutoValHigh
    strS = strS & vbTab & vbTab
    If Not IsNull(tb!DeltaLimit) Then strS = strS & tb!DeltaLimit
    
    gBio.AddItem strS
    
    gBio.Row = gBio.Rows - 1
    gBio.Col = 6 'DoDelta
    If Not IsNull(tb!DoDelta) Then
      Set gBio.CellPicture = IIf(tb!DoDelta, imgSquareTick.Picture, imgSquareCross.Picture)
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


Private Sub optDiscipline_Click(Index As Integer)

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


