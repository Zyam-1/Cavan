VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form fpathistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Search"
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15015
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
   ScaleHeight     =   10410
   ScaleWidth      =   15015
   Begin VB.TextBox txtDoB 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2970
      TabIndex        =   23
      Text            =   "Date of Birth"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10950
      TabIndex        =   19
      Top             =   30
      Width           =   2175
      Begin VB.OptionButton optExact 
         Caption         =   "Exact Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   1275
      End
      Begin VB.OptionButton optLeading 
         Caption         =   "Leading Characters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   600
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton optTrailing 
         Caption         =   "Trailing Characters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   900
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   2340
      TabIndex        =   15
      Top             =   60
      Width           =   5985
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   90
         MaxLength       =   20
         TabIndex        =   17
         Top             =   150
         Width           =   3375
      End
      Begin VB.CheckBox cRemote 
         Caption         =   "Also Search Monaghan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3510
         TabIndex        =   16
         Top             =   210
         Value           =   1  'Checked
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   1995
      Begin VB.OptionButton oHD 
         Caption         =   "Historic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1110
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton oHD 
         Alignment       =   1  'Right Justify
         Caption         =   "Download"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   9045
      Left            =   60
      TabIndex        =   6
      Top             =   1290
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   15954
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"fPatHistory.frx":0000
   End
   Begin VB.CommandButton bcopy 
      Appearance      =   0  'Flat
      Caption         =   "Copy to &Edit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4620
      Picture         =   "fPatHistory.frx":00B6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
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
      Height          =   735
      Left            =   3480
      Picture         =   "fPatHistory.frx":0720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2340
      Picture         =   "fPatHistory.frx":0D8A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   570
      Width           =   1095
   End
   Begin VB.PictureBox SSPanel3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13290
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.TextBox tRecords 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   330
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "fPatHistory.frx":11CC
         Top             =   240
         Width           =   765
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   225
         Left            =   330
         TabIndex        =   13
         Top             =   540
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   397
         _Version        =   327681
         Value           =   25
         BuddyControl    =   "tRecords"
         BuddyDispid     =   196624
         OrigLeft        =   150
         OrigTop         =   450
         OrigRight       =   915
         OrigBottom      =   690
         Increment       =   20
         Max             =   9999
         Min             =   5
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   8340
      ScaleHeight     =   1035
      ScaleWidth      =   2505
      TabIndex        =   5
      Top             =   120
      Width           =   2565
      Begin VB.OptionButton oFor 
         Caption         =   "Name+DoB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1050
         TabIndex        =   24
         Top             =   720
         Width           =   1125
      End
      Begin VB.CheckBox chkSoundex 
         Caption         =   "Use Soundex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1050
         TabIndex        =   18
         Top             =   180
         Width           =   1305
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
         Caption         =   "D.o.B."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Width           =   825
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
         Caption         =   "Chart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   450
         Width           =   735
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   180
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.Label lNoPrevious 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Previous Details"
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   6990
      TabIndex        =   14
      Top             =   750
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "fpathistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private NoPrevious As Boolean
Private mFromEdit As Boolean
Private mEditScreen As Form
Private Activated As Boolean

Private mFromLookup As Boolean

Private pWithin As Integer 'Used for fuzzy DoB search

Private SortOrder As Boolean

Private Sub FillG()

lNoPrevious.Visible = False

With g
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

If Trim$(txtName) = "" Then Exit Sub

If HospName(0) = "Mallow" Then cRemote = 0
If HospName(0) = "Bantry" Then cRemote = 0
If HospName(0) = "Hogwarts" Then cRemote = 0

LocalFillG

'If sysOptRemote(0) And cRemote = 1 Then
'  If Ping(Remote) Then
'    RemoteFillG
'  Else
'    cRemote.Enabled = False
'    cRemote.Caption = IIf(HospName(0) = "Cavan", "Monaghan", "Cavan") & " Network Down."
'    cRemote.Value = 2
'  End If
'End If

With g
  If .Rows > 2 Then
    .RemoveItem 1
    .Row = 1
    .Col = 6
    .ColSel = .Cols - 1
    .RowSel = 1
    .HighLight = flexHighlightAlways
  End If
End With

bcopy.Enabled = mFromEdit

g.Visible = True

Screen.MousePointer = 0
  
End Sub

Public Property Let EditScreen(ByVal F As Form)

Set mEditScreen = F

End Property

Private Function GetRemoteChart(ByVal LocalChart As String) As String

Dim sql As String
Dim tb As Recordset
Dim RegionalNumber As String

On Error GoTo ehgrc

sql = "select * from PatientIFs where " & _
      "Chart = '" & LocalChart & "' " & _
      "and Entity = '" & Entity & "'"

Set tb = New Recordset
RecOpenClient 0, tb, sql

If Not tb.EOF Then
  RegionalNumber = Trim$(tb!RegionalNumber & "")
  If RegionalNumber <> "" Then
    sql = "select * from PatientIFs where " & _
          "RegionalNumber = '" & RegionalNumber & "' " & _
          "and Entity = '" & RemoteEntity & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    If Not tb.EOF Then
      GetRemoteChart = Trim$(tb!Chart & "")
    End If
  End If
End If

Exit Function

ehgrc:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description
    
Screen.MousePointer = 0
LogError "fPatHistory/GetRemoteChart:" & Str(er) & ":" & ers

Exit Function


End Function

Private Sub RemoteFillG()

Dim sql As String
Dim S As String
Dim tb As Recordset
Dim Criteria As String


On Error GoTo eh2

sql = "select top " & Format$(Val(tRecords)) & " * from " & _
      IIf(oHD(0), "PatientIFs", "Demographics") & " where "

Criteria = txtName

If oFor(1) Then
  Criteria = GetRemoteChart(txtName)
  If Criteria = "" Then
    Exit Sub
  End If
End If

If oFor(0) Then
  If chkSoundex = 1 Then
    sql = sql & "Soundex(PatName) "
  Else
    sql = sql & "patname "
  End If
  Criteria = AddTicks(txtName)
ElseIf oFor(1) Then
  sql = sql & "chart "
Else
  sql = sql & "dob "
End If

If oFor(2) Then
  txtName = Convert62Date(txtName, BACKWARD)
  If Not IsDate(txtName) Then
    Screen.MousePointer = 0
    iMsg "Invalid Date", vbExclamation, "Date of Birth Search"
    Exit Sub
  End If
  sql = sql & "= '" & Format$(txtName, "dd/mmm/yyyy") & "'"
ElseIf chkSoundex = 1 Then
  sql = sql & "= soundex('" & Criteria & "') "
ElseIf optExact Then
  sql = sql & "= '" & Criteria & "'"
ElseIf optLeading Then
  sql = sql & "like '" & Criteria & "%'"
Else
  sql = sql & "like '%" & Criteria & "'"
End If
sql = sql & " order by " & IIf(oHD(0), "PatName asc", "RunDate desc")

'NoPrevious = False

Set tb = New Recordset
RecOpenClientRemote tb, sql
With tb
  If .EOF Then
    Screen.MousePointer = 0
    'NoPrevious = True
    'lNoPrevious.Visible = True
  End If

  g.Visible = False
  If oHD(1) Then
    Do While Not .EOF
      S = vbTab & vbTab & vbTab & vbTab
      'If Not IsNull(tb!cFilm) Then
      '  s = s & IIf(tb!cFilm, "F", "")
      'End If
      S = S & vbTab & vbTab & vbTab & _
          Format$(!RunDate, "dd/mm/yy") & vbTab & _
          Trim$(!SampleID & "") & vbTab & _
          !Chart & vbTab & _
          !PatName & vbTab & _
          Format$(!DoB, "dd/mm/yyyy") & vbTab & _
          !Age & vbTab & _
          !Sex & vbTab & _
          !Addr0 & vbTab & _
          !Addr1 & vbTab & _
          !Ward & vbTab & _
          !Clinician & vbTab & _
          !GP & ""
      g.AddItem S
      g.Row = g.Rows - 1
      
      
      If !ForCoag Then
        g.Col = 3
        g.CellBackColor = vbRed
      End If
      If !ForHaem Then
        g.Col = 4
        g.CellBackColor = vbRed
      End If
      If !ForBio Then
        g.Col = 5
        g.CellBackColor = vbRed
      End If
      .MoveNext
    Loop
  Else
    Do While Not .EOF
      S = !Chart & vbTab & _
          !PatName & vbTab & _
          Format$(!DoB, "dd/mm/yyyy") & vbTab & _
          !Sex & vbTab & _
          !Address0 & vbTab & _
          !Address1 & vbTab & _
          !Ward & vbTab & _
          !Clinician & ""
      If !Entity & "" <> "" Then
        S = S & vbTab & IIf(!Entity = "01", "Cavan", "Monaghan")
      End If
      g.AddItem S
      .MoveNext
    Loop
  End If
End With

Exit Sub

eh2:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description
    
Screen.MousePointer = 0
LogError "fPatHistory/RemoteFillG:" & Str(er) & ":" & ers
Exit Sub

End Sub

Private Sub LocalFillG()

Dim sql As String
Dim S As String
Dim tb As Recordset
Dim tbBGA As Recordset
Dim lngSID As Long

On Error GoTo eh2

If oHD(0) Then
  sql = "select top " & Format$(Val(tRecords)) & " * from " & _
        "PatientIFs where "
Else
  sql = "select top " & Format$(Val(tRecords)) & " D.*, H.cFilm " & _
        "from Demographics as D left outer join haemresults as H " & _
        "on D.sampleid = H.sampleid " & _
        "where "
End If

If oFor(0) Then
  If chkSoundex = 1 Then
    sql = sql & "Soundex(PatName) = soundex('" & AddTicks(txtName) & "') "
  Else
    If optExact Then
      sql = sql & "PatName = '" & AddTicks(txtName) & "' "
    ElseIf optLeading Then
      sql = sql & "PatName like '" & AddTicks(txtName) & "%' "
    Else
      sql = sql & "PatName like '%" & AddTicks(txtName) & "' "
    End If
  End If
ElseIf oFor(1) Then
  sql = sql & "chart = '" & AddTicks(txtName) & "' "
ElseIf oFor(2) Then
  txtName = Convert62Date(txtName, BACKWARD)
  If Not IsDate(txtName) Then
    Screen.MousePointer = 0
    iMsg "Invalid Date", vbExclamation, "Date of Birth Search"
    Exit Sub
  End If
  sql = sql & "DoB = '" & Format$(txtName, "dd/mmm/yyyy") & "'"
Else 'Name+DoB
  If chkSoundex = 1 Then
    sql = sql & "Soundex(PatName) = soundex('" & AddTicks(txtName) & "') "
  Else
    If optExact Then
      sql = sql & "PatName = '" & AddTicks(txtName) & "' "
    ElseIf optLeading Then
      sql = sql & "PatName like '" & AddTicks(txtName) & "%' "
    Else
      sql = sql & "PatName like '%" & AddTicks(txtName) & "' "
    End If
  End If
  If pWithin = 0 Then
    sql = sql & "and DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "' "
  Else
    sql = sql & "and DoB between '" & Format$(DateAdd("yyyy", -pWithin, txtDoB), "dd/mmm/yyyy") & "' " & _
                "and '" & Format$(DateAdd("yyyy", pWithin, txtDoB), "dd/mmm/yyyy") & "' "
  End If
End If

sql = sql & " order by " & IIf(oHD(0), "DateTimeAmended desc", "D.RunDate desc")

NoPrevious = False

Set tb = New Recordset
RecOpenClient 0, tb, sql
With tb
  If .EOF Then
    Screen.MousePointer = 0
    NoPrevious = True
'    If mFromLookup Then
'      Unload Me
'      Exit Sub
'    End If
    lNoPrevious.Visible = True
  End If

  g.Visible = False
  If oHD(1) Then
    Do While Not .EOF
      S = vbTab & vbTab & vbTab & vbTab
      If Not IsNull(tb!cFilm) Then
        S = S & IIf(tb!cFilm, "F", "")
      End If
      S = S & vbTab & vbTab & vbTab & _
          Format$(!RunDate, "dd/mm/yy") & vbTab & _
          Trim$(!SampleID & "") & vbTab & _
          !Chart & vbTab & _
          !PatName & vbTab & _
          Format$(!DoB, "dd/mm/yyyy") & vbTab & _
          !Age & vbTab & _
          !Sex & vbTab & _
          !Addr0 & vbTab & _
          !Addr1 & vbTab & _
          !Ward & vbTab & _
          !Clinician & vbTab & _
          !GP & vbTab
      If Trim$(!Chart & "") <> "" And Trim$(!PatName & "") <> "" Then
        sql = "Select * from PatientIFs where " & _
              "Chart = '" & !Chart & "' " & _
              "and PatName = '" & AddTicks(!PatName) & "'"
        Set tbBGA = New Recordset
        RecOpenServer 0, tbBGA, sql
        If Not tbBGA.EOF Then
          If tbBGA!Entity & "" <> "" Then
            If tbBGA!Entity = "01" Then
              S = S & "Cavan"
              sql = "Update Demographics " & _
                    "set Hospital = 'Cavan' where " & _
                    "Chart = '" & !Chart & "' " & _
                    "and PatName = '" & AddTicks(!PatName) & "'"
              Cnxn(0).Execute sql
            ElseIf tbBGA!Entity = "31" Then
              S = S & "Monaghan"
              sql = "Update Demographics " & _
                    "set Hospital = 'Monaghan' where " & _
                    "Chart = '" & !Chart & "' " & _
                    "and PatName = '" & AddTicks(!PatName) & "'"
              Cnxn(0).Execute sql
            End If
          End If
        End If
      End If
      'If !Entity & "" <> "" Then
     '   s = s & vbTab & IIf(!Entity = "01", "Cavan", "Monaghan")
     ' End If
      g.AddItem S
      g.Row = g.Rows - 1
      
      If sysOptDeptExt(0) Then
        sql = "Select * from ExtResults where SampleID = '" & !SampleID & "'"
        Set tbBGA = New Recordset
        RecOpenServer 0, tbBGA, sql
        If Not tbBGA.EOF Then
          g.Col = 0
          g.CellBackColor = vbRed
        End If
      End If
      
      If sysOptDeptMedibridge(0) Then
        sql = "Select * from MedibridgeResults where SampleID = '" & !SampleID & "'"
        Set tbBGA = New Recordset
        RecOpenServer 0, tbBGA, sql
        If Not tbBGA.EOF Then
          g.Col = 1
          g.CellBackColor = vbRed
        End If
      End If
      
      If sysOptDeptMicro(0) Then
      
        lngSID = Val(g.TextMatrix(g.Row, 8))
        
        If lngSID > sysOptMicroOffset(0) Then
          g.TextMatrix(g.Row, 8) = Format$(Val(g.TextMatrix(g.Row, 8)) - sysOptMicroOffset(0))
          g.Col = 2
          g.CellBackColor = vbRed
        ElseIf lngSID > sysOptSemenOffset(0) Then
          'Semen Result
        End If
      End If
      
      If !ForCoag Then
        g.Col = 3
        g.CellBackColor = vbRed
      End If
      If !ForHaem Then
        g.Col = 4
        g.CellBackColor = vbRed
      End If
      If !ForBio Then
        g.Col = 5
        g.CellBackColor = vbRed
      End If
      
      If sysOptDeptBga(0) Then
        sql = "Select * from BGAResults where SampleID = '" & !SampleID & "'"
        Set tbBGA = New Recordset
        RecOpenServer 0, tbBGA, sql
        If Not tbBGA.EOF Then
          g.Col = 6
          g.CellBackColor = vbRed
        End If
      End If
      .MoveNext
    Loop
  Else
    Do While Not .EOF
      S = !Chart & vbTab & _
          !PatName & vbTab & _
          Format$(!DoB, "dd/mm/yyyy") & vbTab & _
          !Sex & vbTab & _
          !Address0 & vbTab & _
          !Address1 & vbTab & _
          !Ward & vbTab & _
          !Clinician & ""
      If !Entity & "" <> "" Then
        S = S & vbTab & IIf(!Entity = "01", "Cavan", "Monaghan")
      End If
      g.AddItem S
      .MoveNext
    Loop
  End If
End With

Exit Sub

eh2:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description
    
Screen.MousePointer = 0
LogError "fPatHistory/LocalFillG:" & Str(er) & ":" & ers
Exit Sub

End Sub

Public Property Get NoPreviousDetails() As Variant

NoPreviousDetails = NoPrevious
  
End Property

Private Sub bcancel_Click()

Unload Me

End Sub

Private Sub bCopy_Click()

Dim gRow As Integer
Dim strWard As String
Dim strGP As String
Dim strSex As String
Dim strName As String

'On Error Resume Next

gRow = g.Row

With mEditScreen
  If oHD(1) Then
    If .txtChart = "" Then
      .txtChart = g.TextMatrix(gRow, 9)
    End If
    strName = Initial2Upper(g.TextMatrix(gRow, 10))
    .txtName = strName
    .txtDoB = g.TextMatrix(gRow, 11)
    .txtAge = CalcAge(.txtDoB)
    strSex = g.TextMatrix(gRow, 13)
    If strSex = "" Then
      NameLostFocus strName, strSex
    End If
    .txtSex = strSex
    .txtAddress(0) = Initial2Upper(g.TextMatrix(gRow, 14))
    .txtAddress(1) = Initial2Upper(g.TextMatrix(gRow, 15))
    strWard = Initial2Upper(g.TextMatrix(gRow, 16))
    strGP = Initial2Upper(g.TextMatrix(gRow, 18))
    If strWard = "" And strGP <> "" Then
      strWard = "GP"
    End If
    .cmbWard = strWard
    .cmbGP = strGP
    .cmbClinician = Initial2Upper(g.TextMatrix(gRow, 17))
  Else
    .txtChart = g.TextMatrix(gRow, 0)
    .txtName = Initial2Upper(g.TextMatrix(gRow, 1))
    .txtDoB = g.TextMatrix(gRow, 2)
    .txtAge = CalcAge(.txtDoB)
    .txtSex = g.TextMatrix(gRow, 3)
    .txtAddress(0) = Initial2Upper(g.TextMatrix(gRow, 4))
    .txtAddress(1) = Initial2Upper(g.TextMatrix(gRow, 5))
    .cmbWard = Initial2Upper(g.TextMatrix(gRow, 6))
    .cmbClinician = Initial2Upper(g.TextMatrix(gRow, 7))
  End If
End With

Unload Me

End Sub

Private Sub bsearch_Click()

FillG

End Sub

Private Sub chkSoundex_Click()

If chkSoundex = 1 Then
  fraSearch.Visible = False
Else
  fraSearch.Visible = True
End If

FillG

End Sub

Private Sub Form_Activate()

If Activated Then Exit Sub

If HospName(0) = "Monaghan" Then
  If mFromEdit Then
    bcopy.Enabled = True
  Else
    bcopy.Enabled = False
  End If
End If

If HospName(0) = "Mallow" Then cRemote = 0

Activated = True

txtName.SetFocus

End Sub

Private Sub Form_Load()
    
Activated = False

LoadHeading IIf(oHD(0), 0, 1)

If (HospName(0) = "Cavan" Or HospName(0) = "Monaghan") And sysOptRemote(0) Then
  cRemote.Visible = True
  If Ping(Remote) Then
    cRemote.Enabled = True
    cRemote.Caption = "Also Search " & IIf(HospName(0) = "Cavan", "Monaghan", "Cavan")
    cRemote.Value = Val(GetSetting("NetAcquire", "PatSearch", "cRemote", "1"))
  Else
    cRemote.Enabled = False
    cRemote.Caption = IIf(HospName(0) = "Cavan", "Monaghan", "Cavan") & " Network Down."
    cRemote.Value = 2
  End If
Else
  cRemote.Visible = False
End If

End Sub


Private Sub LoadHeading(ByVal Index As Integer)

Dim n As Integer

For n = 0 To 6
  g.ColWidth(n) = 250
Next

If Index = 0 Then
  
  g.Cols = 9
  g.FormatString = "<Chart     |<Name                         |<Date of Birth" & _
                   "|<Sex|<Address                     |<                    " & _
                   "|<Ward                 |<Clinician              |<Hospital         "
  g.Col = 2
  g.Row = 0
  txtDoB.Left = g.Left + g.CellLeft
  txtDoB.Width = g.CellWidth

Else
  g.Cols = 20
  g.FormatString = "E|B|M|C|H|B|G|<Run Date |<Run #           |<Chart      |" & _
                   "<Name                             |<Date of Birth|<Age|^Sex|" & _
                   "<Address        |<  |<Ward           |<Clinician           |<GP                |<Hospital   "
  
  If Not sysOptDeptExt(0) Then
    g.ColWidth(0) = 0
  End If
  If Not sysOptDeptMedibridge(0) Then
    g.ColWidth(1) = 0
  End If
  If Not sysOptDeptMicro(0) Then
    g.ColWidth(2) = 0
  End If
  If Not sysOptDeptBga(0) Then
    g.ColWidth(6) = 0
  End If
  
  g.Col = 11
  g.Row = 0
  txtDoB.Left = g.Left + g.CellLeft
  txtDoB.Width = g.CellWidth

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Activated = False
  
If cRemote.Value <> 2 Then
  SaveSetting "NetAcquire", "PatSearch", "cRemote", Format$(cRemote.Value)
End If

End Sub


Private Sub g_Click()

Dim tb As Recordset
Dim sql As String
Dim NewChart As String
Dim PatName As String
Dim DoB As String

On Error GoTo ehgc

If g.MouseRow = 0 Then
  If InStr(UCase$(g.TextMatrix(0, g.Col)), "DATE") <> 0 Then
    g.Sort = 9
  Else
    If SortOrder Then
      g.Sort = flexSortGenericAscending
    Else
      g.Sort = flexSortGenericDescending
    End If
  End If
  SortOrder = Not SortOrder
  Exit Sub
End If

If oHD(0) Then
  If g.Col > 1 And mFromEdit Then
    g.Col = 0
    g.ColSel = g.Cols - 1
    g.RowSel = g.Row
    g.HighLight = flexHighlightAlways
    bcopy.Enabled = True
  ElseIf g.Col = 0 Then
    If Trim$(g.TextMatrix(g.Row, 0)) = "" Then Exit Sub
    PatName = g.TextMatrix(g.Row, 1)
    If Trim$(PatName) = "" Then Exit Sub
    DoB = g.TextMatrix(g.Row, 2)
    If IsDate(DoB) Then
      DoB = Format$(DoB, "dd/mmm/yyyy")
    Else
      Exit Sub
    End If
    If iMsg("Do you want to change the Chart Number?", vbQuestion + vbYesNo) = vbYes Then
      NewChart = iBOX("New Chart Number", , g.TextMatrix(g.Row, 0))
      sql = "Update Demographics " & _
            "set Chart = '" & NewChart & "' where " & _
            "PatName = '" & AddTicks(PatName) & "' " & _
            "and dob = '" & DoB & "'"
      Set tb = New Recordset
      RecOpenClient 0, tb, sql
      FillG
    End If
  End If
Else
  If g.Col = 9 Then 'chart
    If Trim$(g.TextMatrix(g.Row, 9)) = "" Then Exit Sub
    PatName = g.TextMatrix(g.Row, 10)
    If Trim$(PatName) = "" Then Exit Sub
    DoB = g.TextMatrix(g.Row, 11)
    If IsDate(DoB) Then
      DoB = Format$(DoB, "dd/mmm/yyyy")
    Else
      Exit Sub
    End If
    If iMsg("Do you want to change the Chart Number?", vbQuestion + vbYesNo) = vbYes Then
      NewChart = iBOX("New Chart Number", , g.TextMatrix(g.Row, 9))
      sql = "Update Demographics " & _
            "set Chart = '" & NewChart & "' where " & _
            "PatName = '" & AddTicks(PatName) & "' " & _
            "and dob = '" & DoB & "'"
      Set tb = New Recordset
      RecOpenClient 0, tb, sql
      FillG
    End If
  ElseIf g.Col > 6 Then
    g.Col = 7
    g.ColSel = g.Cols - 1
    g.RowSel = g.Row
    g.HighLight = flexHighlightAlways
    If mFromEdit Then
      bcopy.Enabled = True
    End If
  Else
    bcopy.Enabled = False

    If g.CellBackColor <> vbRed Then Exit Sub
  
    If g.Col = 1 Then
      With frmViewMedibridge
        .SampleID = g.TextMatrix(g.Row, 8)
        .Show 1
      End With
    ElseIf g.Col = 2 Then 'Micro
      With frmMicroReport
        .lblChart = g.TextMatrix(g.Row, 9)
        .lblName = g.TextMatrix(g.Row, 10)
        .lblDoB = g.TextMatrix(g.Row, 11)
        .Show 1
      End With
    Else
      With frmViewResults
        .lblSampleID = g.TextMatrix(g.Row, 8)
        .lblChart = g.TextMatrix(g.Row, 9)
        .lblName = g.TextMatrix(g.Row, 10)
        .lblDoB = g.TextMatrix(g.Row, 11)
        .Show 1
      End With
    End If
  End If
End If

Exit Sub

ehgc:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "fPatHistory/G_Click:" & Str(er) & ":" & ers
Exit Sub

End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim d1 As String
Dim d2 As String

If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
  Cmp = 0
  Exit Sub
End If

If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
  Cmp = 0
  Exit Sub
End If

d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

If SortOrder Then
  Cmp = Sgn(DateDiff("s", d1, d2))
Else
  Cmp = Sgn(DateDiff("s", d2, d1))
End If
End Sub


Private Sub oFor_Click(Index As Integer)

Dim F As Form

Select Case Index
  
  Case 0:    optLeading = True
             chkSoundex.Enabled = True
             txtDoB.Visible = False
  
  Case 1, 2: optExact = True
             chkSoundex.Enabled = False
             chkSoundex = 0
             txtDoB.Visible = False
  
  Case 3:    optLeading = True
             chkSoundex.Enabled = True
             
             Set F = New frmGetDoB
             F.Show 1
             txtDoB = F.txtDoB
             If F.lblWithin.Enabled Then
               pWithin = F.lblWithin
             Else
               pWithin = 0
             End If
             Unload F
             Set F = Nothing
             
             txtDoB.Visible = True
           
             If txtName.Visible Then
               txtName.SetFocus
             End If
             
End Select


g.Rows = 2
g.AddItem ""
g.RemoveItem 1
txtName = ""

End Sub

Private Sub oHD_Click(Index As Integer)

g.Rows = 2
g.AddItem ""
g.RemoveItem 1

LoadHeading Index

If Not Activated Then Exit Sub

FillG

End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)

If oFor(0) Or oFor(3) Then
  If Len(Trim$(txtName)) > 3 Then
    FillG
  End If
End If

End Sub


Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

FillG

End Sub



Public Property Let FromEdit(ByVal x As Boolean)

mFromEdit = x

End Property

Public Property Let FromLookUp(ByVal bNewValue As Boolean)

mFromLookup = bNewValue

End Property
