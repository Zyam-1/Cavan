VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrackMatchNames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDistrict 
      Height          =   315
      Left            =   3870
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   510
      Width           =   3375
   End
   Begin VB.ComboBox cmbProvince 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   510
      Width           =   3375
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   525
      Left            =   12540
      TabIndex        =   4
      Top             =   8010
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   12450
      TabIndex        =   3
      Top             =   6000
      Width           =   1425
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   12420
      TabIndex        =   2
      Top             =   1350
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid gNames 
      Height          =   7455
      Left            =   90
      TabIndex        =   0
      Top             =   1080
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"frmTrackMatchNames.frx":0000
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
   Begin VB.Label Label3 
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   12450
      TabIndex        =   9
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "District Laboratory"
      Height          =   195
      Left            =   3900
      TabIndex        =   8
      Top             =   270
      Width           =   3315
   End
   Begin VB.Label Label1 
      Caption         =   "Provincial Laboratory"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   270
      Width           =   3315
   End
   Begin VB.Label lblReady 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready for Scanning"
      Height          =   525
      Left            =   12450
      TabIndex        =   1
      Top             =   5310
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmTrackMatchNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

Dim tb As Recordset
Dim sql As String
Dim d As String
Dim s As String

With gNames
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

sql = "SELECT GUID, NID, SurName, ForeName, Sex, DoB, " & _
      "DistrictBatchNumber, ProvincialBatchNumber FROM TrackMessage WHERE " & _
      "( Province = '" & cmbProvince.Text & "' " & _
      "  AND District = '" & cmbDistrict.Text & "' AND (SampleID IS NULL OR SampleID = 0) ) " & _
      "ORDER BY SurName"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  s = tb!NID & vbTab & _
      tb!SurName & vbTab & _
      tb!ForeName & vbTab & _
      tb!Sex & vbTab
  d = Format(tb!DoB & "", "Short Date")
  If IsDate(d) Then
    s = s & d
  End If
  s = s & vbTab & tb!DistrictBatchNumber & vbTab & _
      tb!ProvincialBatchNumber & vbTab & _
      "" & vbTab & _
      tb!GUID & ""
  gNames.AddItem s
  tb.MoveNext
Loop

lblReady.Visible = False
With gNames
  If .Rows > 2 Then
    .RemoveItem 1
    lblReady.Visible = True
  End If
End With

End Sub

Private Function FindFirstBlank() As Boolean

'returns true if a blank line is found
Dim y As Integer

FindFirstBlank = False

gNames.Col = 7
For y = 1 To gNames.Rows - 1
  gNames.Row = y
  gNames.CellBackColor = 0
Next

If gNames.Rows = 2 And gNames.TextMatrix(1, 1) = "" Then
  Exit Function
End If

For y = 1 To gNames.Rows - 1
  If gNames.TextMatrix(y, 7) = "" Then
    gNames.Row = y
    gNames.CellBackColor = vbYellow
    FindFirstBlank = True
    Exit For
  End If
Next

End Function

Private Sub InitialFill()

Dim f As Integer
Dim ImportPath As String
Dim FileName As String
Dim strIP As String
Dim sql As String
Dim tb As Recordset
Dim FileLen As Long
Dim n As Integer

gNames.Rows = 2
gNames.AddItem ""
gNames.RemoveItem 1

lblReady.Visible = False

sql = "Select * from Options where " & _
      "Description = 'TransportImport'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  iMsg "Import Path not set"
  Exit Sub
Else
  ImportPath = tb!Contents & ""
End If

FileName = Dir(ImportPath)
Do While FileName <> ""
  f = FreeFile
  Open ImportPath & FileName For Input As #f
  strIP = ""
  Do While Not EOF(f)
    strIP = strIP & Input(1, #f)
  Loop
  Close f
  If Not Parse(strIP) Then
    FileCopy ImportPath & FileName, ImportPath & "Hold\" & FileName
  Else
    FileCopy ImportPath & FileName, ImportPath & "Success\" & FileName
  End If
  Kill ImportPath & FileName
    
  FileName = Dir()
Loop

FillProvince

End Sub

Private Sub cmbDistrict_Click()

FillG

If FindFirstBlank Then
  If txtInput.Visible Then
    txtInput.SetFocus
  End If
End If

End Sub

Private Sub cmbProvince_Click()

FillDistrict

End Sub


Private Sub cmdExit_Click()

Unload Me

End Sub

Private Function Parse(ByVal strIP As String) As Boolean
'Returns True if success
Dim Lines() As String
Dim Items() As String
Dim n As Integer

On Error GoTo Parse_Error

Parse = False

If strIP = "" Then Exit Function

strIP = Replace(strIP, vbLf, "")
Lines = Split(strIP, vbCr)
For n = 0 To UBound(Lines)
  Items() = Split(Lines(n), "|")
  If UBound(Items()) < 42 Then Exit Function
  Select Case UCase$(Items(0))
    Case "RESULT"
      If Not ParseResult(Items()) Then
        Exit Function
      End If
    Case "REQUEST"
      If Not ParseRequest(Items()) Then
        Exit Function
      End If
    Case Else
      Exit Function
  End Select
Next

Parse = True

Exit Function

Parse_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogErrorA "frmTrackMatchNames", "Parse", intEL, strES

End Function
Private Sub FillProvince()

Dim tb As Recordset
Dim sql As String

cmbProvince.Clear

sql = "SELECT DISTINCT Province FROM TrackMessage WHERE " & _
      "SampleID IS NULL OR SampleID = 0 " & _
      "ORDER BY Province"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbProvince.AddItem tb!Province & ""
  tb.MoveNext
Loop
If cmbProvince.ListCount > 0 Then
  cmbProvince.ListIndex = 0
End If

End Sub

Private Sub FillDistrict()

Dim tb As Recordset
Dim sql As String

cmbDistrict.Clear

sql = "SELECT DISTINCT District FROM TrackMessage WHERE " & _
      "( Province = '" & cmbProvince.Text & "' " & _
      "  AND (SampleID IS NULL OR SampleID = 0) ) " & _
      "ORDER BY District"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbDistrict.AddItem tb!District & ""
  tb.MoveNext
Loop

If cmbDistrict.ListCount > 0 Then
  cmbDistrict.ListIndex = 0
End If

End Sub

Private Function ParseResult(ByRef Items() As String) As Boolean

End Function



Private Function ParseRequest(ByRef Items() As String) As Boolean

Dim tb As Recordset
Dim sql As String
Dim GUID As String
Dim s As String
Dim d As String

On Error GoTo ParseRequest_Error

ParseRequest = True

If UBound(Items()) < 42 Then Exit Function

GUID = Items(1)

sql = "SELECT * FROM TrackMessage WHERE " & _
      "GUID = '" & GUID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If

tb!MessageType = Items(0)
tb!GUID = Items(1)
tb!NID = Items(2)
tb!SurName = Items(3)
tb!ForeName = Items(4)
tb!NameOfMother = Items(5)
tb!NameOfFather = Items(6)
tb!Race = Items(7)
tb!Profession = Items(8)
tb!WhereBorn = Items(9)
tb!Address0 = Items(10)
tb!Address1 = Items(11)
tb!Phone = Items(12)
tb!Sex = Items(13)

d = Convert82Date(Items(14))
If IsDate(d) Then
  tb!DoB = d
End If

tb!Age = Items(15)
tb!BreastFed = IIf(Items(16) <> "", 1, 0)
tb!HealthFacility = Items(17)
tb!Province = Items(18)
tb!District = Items(19)
tb!ClinicType = Items(20)
tb!Clinician = Items(21)

d = Convert82Date(Items(22))
If IsDate(d) Then
  tb!CollectionDate = d
End If

d = Convert82Date(Items(23))
If IsDate(d) Then
  tb!BreastFedEnd = Format(d, "yyyy mm dd")
End If

d = Convert82Date(Items(24))
If IsDate(d) Then
  tb!DateSent = d
End If

d = Convert82Date(Items(25))
If IsDate(d) Then
  tb!DateResultSent = d
End If

d = Convert82Date(Items(26))
If IsDate(d) Then
  tb!DateResultReceived = d
End If

tb!TestRequired = Items(27)
tb!TestRequiredCode = Items(28)
tb!Discipline = Items(29)
tb!SampleType = Items(30)
tb!Result = Items(31)
tb!Units = Items(32)
tb!Flags = Items(33)
tb!NormalRange = Items(34)
tb!PreviousPCR = Items(35)
tb!PreviousPCRResult = Items(36)
tb!DemographicsComment = Items(37)
tb!ResultComment = Items(38)
tb!DistrictBatchNumber = Items(39)
tb!ProvincialBatchNumber = Items(40)
tb!SampleID = Val(Items(41))
tb!AnalysedAt = Items(42)
tb.Update

Exit Function

ParseRequest_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogErrorA "frmTrackMatchNames", "ParseRequest", intEL, strES, sql
ParseRequest = False

End Function

Private Function Convert82Date(ByVal s As String) _
                               As String

Dim d As String

If Len(s) <> 8 Then
  Convert82Date = s
  Exit Function
End If

d = Right$(s, 2) & "/" & Mid$(s, 5, 2) & "/" & Left$(s, 4)
If IsDate(d) Then
  Convert82Date = Format$(d, "Short Date")
Else
  Convert82Date = s
End If


End Function

Private Sub cmdSave_Click()

Dim tb As Recordset
Dim tbDem As Recordset
Dim tbBat As Recordset
Dim sql As String
Dim y As Integer

For y = 1 To gNames.Rows - 1
  If Val(gNames.TextMatrix(y, 7)) <> 0 Then 'SID
    sql = "SELECT * FROM TrackMessage WHERE " & _
          "GUID = '" & gNames.TextMatrix(y, 8) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
      sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & gNames.TextMatrix(y, 7) & "'"
      Set tbDem = New Recordset
      RecOpenServer 0, tbDem, sql
      If tbDem.EOF Then
        tbDem.AddNew
      End If
      tbDem!SampleID = Val(gNames.TextMatrix(y, 7))
      tbDem!Chart = tb!NID & ""
      tbDem!PatName = tb!SurName & " " & tb!ForeName & ""
      tbDem!Age = tb!Age & ""
      tbDem!Sex = tb!Sex & ""
      'tbDem!TimeTaken = vbNull
      'Source
      'RunDate
      If IsDate(tb!DoB) Then
        tbDem!DoB = tb!DoB
      Else
        tbDem!DoB = vbNull
      End If
      tbDem!Addr0 = tb!Address0 & ""
      tbDem!Addr1 = tb!Address1 & ""
      tbDem!Ward = tb!District & ""
      tbDem!Clinician = tb!Clinician & ""
      'GP
      If IsDate(tb!CollectionDate) Then
        tbDem!SampleDate = tb!CollectionDate
      Else
        tbDem!SampleDate = vbNull
      End If
      'ClDetails
      tbDem!Hospital = tb!Province & ""
      'RooH
      tbDem!FAXed = False
      tbDem!Fasting = False
      tbDem!OnWarfarin = False
      tbDem!DateTimeDemographics = Format(Now, "Short Date") & " " & Format(Now, "Long Time")
      '0 DateTimeHaemPrinted datetime  8 1
      '0 DateTimeBioPrinted  datetime  8 1
      '0 DateTimeCoagPrinted datetime  8 1
      '0 Pregnant  bit 1 1
      '0 AandE nvarchar  50  1
      '0 NOPAS nvarchar  50  1
      tbDem!RecDate = Format(Now, "Short Date") & " " & Format(Now, "Long Time")
      tbDem!RecordDateTime = Format(Now, "Short Date") & " " & Format(Now, "Long Time")
      '0 Category  nvarchar  50  1
      '0 HistoValid  bit 1 1
      '0 CytoValid bit 1 1
      '0 Mrn nvarchar  50  1
      tbDem!UserName = UserName
      '0 Urgent  int 4 1
      tbDem!Valid = False
      '0 HYear nvarchar  50  1
      '0 SentToEMedRenal int 4 1
      tbDem!SurName = tb!SurName & ""
      tbDem!ForeName = tb!ForeName & ""
      tbDem!NameOfMother = tb!NameOfMother & ""
      tbDem!NameOfFather = tb!NameOfFather & ""
      tbDem!ClinicType = tb!ClinicType & ""
      tbDem!BreastFed = tb!BreastFed
      If IsDate(tb!BreastFedEnd) Then
        tbDem!BreastFedEnd = tb!BreastFedEnd
      Else
        tbDem!BreastFedEnd = vbNull
      End If
      tbDem!Race = tb!Race & ""
      tbDem!Profession = tb!Profession & ""
      tbDem!WhereBorn = tb!WhereBorn & ""
      tbDem!Phone = tb!Phone & ""
      tbDem!HealthFacility = tb!HealthFacility & ""
      tbDem!ClinicType = tb!ClinicType & ""
      tbDem.Update
      
      tb!SampleID = Val(gNames.TextMatrix(y, 7))
      tb.Update
      
    End If
    
    sql = "Select * from TrackBatchNumbers WHERE " & _
          "SampleID = '" & Val(gNames.TextMatrix(y, 7)) & "'"
    Set tbBat = New Recordset
    RecOpenServer 0, tbBat, sql
    If tbBat.EOF Then
      tbBat.AddNew
    End If
    tbBat!SampleID = Val(gNames.TextMatrix(y, 7))
    tbBat!DistrictBatchNumber = tb!DistrictBatchNumber & ""
    tbBat!ProvincialBatchNumber = tb!ProvincialBatchNumber & ""
    tbBat!GUID = tb!GUID & ""
    tbBat.Update
  End If
Next

gNames.Rows = 2
gNames.AddItem ""
gNames.RemoveItem 1

FillProvince

'0 DateSent  datetime  8 1
'0 DateResultSent  datetime  8 1
'0 DateResultReceived  datetime  8 1
End Sub

Private Sub Form_Activate()

If FindFirstBlank Then
  txtInput.SetFocus
End If

End Sub

Private Sub Form_Load()

gNames.ColWidth(8) = 0

InitialFill

End Sub

Private Sub gNames_Click()

Dim y As Integer
Dim ySave As Integer

If gNames.MouseRow = 0 Then
  If gNames.MouseCol = 4 Then
    gNames.Sort = 9
  ElseIf SortOrder Then
    gNames.Sort = flexSortGenericAscending
  Else
    gNames.Sort = flexSortGenericDescending
  End If
  SortOrder = Not SortOrder
  
  Exit Sub
End If

ySave = gNames.Row

gNames.Col = 7
For y = 1 To gNames.Rows - 1
  gNames.Row = y
  gNames.CellBackColor = 0
Next

gNames.Row = ySave
gNames.CellBackColor = vbYellow

txtInput.SetFocus
lblReady.Visible = True

End Sub


Private Sub gNames_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

Dim d1 As String
Dim d2 As String

If Not IsDate(gNames.TextMatrix(Row1, gNames.Col)) Then
  Cmp = 0
  Exit Sub
End If

If Not IsDate(gNames.TextMatrix(Row2, gNames.Col)) Then
  Cmp = 0
  Exit Sub
End If

d1 = Format(gNames.TextMatrix(Row1, gNames.Col), "dd/mmm/yyyy hh:mm:ss")
d2 = Format(gNames.TextMatrix(Row2, gNames.Col), "dd/mmm/yyyy hh:mm:ss")

If SortOrder Then
  Cmp = Sgn(DateDiff("s", d1, d2))
Else
  Cmp = Sgn(DateDiff("s", d2, d1))
End If

End Sub


Private Sub txtinput_GotFocus()

txtInput = ""

End Sub

Private Sub txtinput_LostFocus()

Dim y As Integer

If txtInput = "" Then Exit Sub

gNames.Col = 7
For y = 1 To gNames.Rows - 1
  gNames.Row = y
  If gNames.CellBackColor = vbYellow Then
    gNames.TextMatrix(y, 7) = txtInput
    Exit For
  End If
Next

If FindFirstBlank Then
  lblReady.Visible = True
  txtInput.SetFocus
Else
  lblReady.Visible = False
End If

End Sub


