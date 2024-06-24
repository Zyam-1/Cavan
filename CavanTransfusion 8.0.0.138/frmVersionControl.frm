VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVersionControl 
   Caption         =   "NetAcquire - Version Control"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "frmVersionControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLat2 
      BackColor       =   &H8000000A&
      Height          =   1980
      Left            =   360
      TabIndex        =   8
      Top             =   1980
      Width           =   7620
      Begin VB.Frame fraLatest 
         BackColor       =   &H0080FFFF&
         Caption         =   "Latest Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   165
         TabIndex        =   9
         Top             =   270
         Width           =   7290
         Begin VB.CommandButton cmdActivateLatest 
            Caption         =   "Activate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   5760
            TabIndex        =   10
            Top             =   435
            Width           =   1350
         End
         Begin VB.Label lblLatestDate 
            BackColor       =   &H0080FFFF&
            Height          =   195
            Left            =   3900
            TabIndex        =   14
            Top             =   900
            Width           =   1545
         End
         Begin VB.Label lblLatestVerNumber 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   990
            TabIndex        =   13
            Top             =   930
            Width           =   2025
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "Version"
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   930
            Width           =   525
         End
         Begin VB.Label lblLatestVersion 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   11
            Top             =   480
            Width           =   5145
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   7020
      Picture         =   "frmVersionControl.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "cancel"
      ToolTipText     =   "Exit"
      Top             =   690
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdVersions 
      Height          =   3795
      Left            =   330
      TabIndex        =   0
      Top             =   4500
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "<Filename                                                 |<Version      |<Deployed |<DateTime Created        "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraActLat 
      Height          =   1725
      Left            =   360
      TabIndex        =   1
      Top             =   90
      Width           =   6030
      Begin VB.Frame fraActive 
         BackColor       =   &H0080C0FF&
         Caption         =   "Active Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   135
         TabIndex        =   2
         Top             =   240
         Width           =   5730
         Begin VB.Label lblActiveVersion 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   285
            TabIndex        =   5
            Top             =   465
            Width           =   5160
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Version:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   300
            TabIndex        =   4
            Top             =   915
            Width           =   930
         End
         Begin VB.Label lblActiveVersionNo 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1155
            TabIndex        =   3
            Top             =   900
            Width           =   2595
         End
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   330
      TabIndex        =   15
      Top             =   8460
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Version Deployment History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2918
      TabIndex        =   6
      Top             =   4200
      Width           =   2505
   End
End
Attribute VB_Name = "frmVersionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActivateLatest_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdActivateLatest_Click_Error

20    If lblLatestVersion = "" Then Exit Sub

30    If UCase$(Trim$(lblLatestVersion)) = UCase$(Trim$(lblActiveVersion)) Then
40        iMsg "This version of NetAcquire is currently active!"
50        If TimedOut Then Unload Me: Exit Sub
60        Exit Sub
70    End If

80    If Not AllowedToActivateVersion(lblLatestVersion) Then
90        iMsg "This version of NetAcquire cannot be activated!"
100       If TimedOut Then Unload Me: Exit Sub
110       Exit Sub
120   End If

130   blnEndApp = True

140   sql = "Update VersionControl set Active = 0"
150   CnxnBB(0).Execute sql

160   sql = "Select * from VersionControl where FileName = '" & lblLatestVersion & "'"

170   Set tb = New Recordset
180   RecOpenServerBB 0, tb, sql

190   If tb.EOF Then
200     tb.AddNew
210   End If
220   tb!FileName = Trim$(lblLatestVersion)
230   tb!File_Version = lblLatestVerNumber
240   tb!File_DateCreated = lblLatestDate
250   tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
260   tb!Deployed = 1
270   tb!Active = 1
280   tb.Update

290   strLatestVersion = Trim$(lblLatestVersion)
      'Change Desktop Shortcut
300   CreateShortcut (strLatestVersion)

310   Unload Me

320   Exit Sub

cmdActivateLatest_Click_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmVersionControl", "cmdActivateLatest_Click", intEL, strES, sql


End Sub



Private Sub cmdCancel_Click()
10    Unload Me
End Sub


Private Sub Form_Load()

      Dim LatestFileName As String
      Dim LatestFileDate As Date

      Dim ActiveFileName As String
      Dim ActiveFileDate As Date

      Dim FileBeingTested As String

      Dim Found As Boolean
      Dim Path As String
      Dim fso As New FileSystemObject
      Dim fil As File
      Dim sql As String
      Dim tb As Recordset
      Dim LatestVerNumber As String
      Dim s As String

10    On Error GoTo Form_Load_Error

20    Found = False
    
30    Set fso = CreateObject("Scripting.filesystemobject")

      'Application Path
40    Path = App.Path & "\"

      '1st EXE LatestFileName to check
50    ActiveFileName = UCase$(App.EXEName & ".EXE") 'UCase$(Dir(Path & "*.exe", vbNormal))

      'Display Active Application name and version
60    lblActiveVersion = ActiveFileName
70    lblActiveVersionNo = fso.GetFileVersion(Path & lblActiveVersion)
80    Set fil = fso.GetFile(Path & ActiveFileName)
90    ActiveFileDate = fil.DateLastModified

100   LatestFileDate = ActiveFileDate
110   FileBeingTested = UCase$(Dir(Path))
120   Do While FileBeingTested <> ""
130     If Right$(FileBeingTested, 4) = ".EXE" Then
140       Set fil = fso.GetFile(Path & FileBeingTested)
150       If fil.DateLastModified > LatestFileDate Then
160         LatestFileDate = fil.DateLastModified
170         LatestVerNumber = fso.GetFileVersion(Path & FileBeingTested)
180         LatestFileName = FileBeingTested
190       End If
200     End If
210     FileBeingTested = UCase$(Dir())
220   Loop
230   If LatestFileDate = ActiveFileDate Then
240       fraLat2.Visible = False
250       fraActive.Caption = fraActive.Caption & " and " & fraLatest.Caption
260   Else
270     lblLatestVersion = UCase$(Trim$(LatestFileName))
280     lblLatestVerNumber = UCase$(Trim$(LatestVerNumber))
290     lblLatestDate = UCase$(Trim$(LatestFileDate))
300   End If

      'CheckVersionControlInDb CnxnBB(0)

310   sql = "Select * from VersionControl where FileName = '" & lblActiveVersion & "'"

320   Set tb = New Recordset
330   RecOpenServerBB 0, tb, sql

340   If tb.EOF Then
350     tb.AddNew
360     tb!FileName = Trim$(lblActiveVersion)
370     tb!File_Version = lblActiveVersionNo
380     tb!File_DateCreated = ActiveFileDate
390     tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
400     tb!Deployed = 1
410     tb!Active = 1
420     tb.Update
430   End If
      '    ResultsCreationTime = fil.DateLastModified
      '
      '    SQL = "Select * from VersionControl where LatestFileName = '" & LatestFileName & "'"
      '
      '    Set tb = New Recordset
      '
      '    RecOpenServerBB 0, tb, SQL
      '
      '    If tb.EOF Then
      '        tb.AddNew
      '    End If
      '    tb!LatestFileName = Trim$(LatestFileName)
      '    tb!File_Version = LatestVerNumber
      '    tb!File_DateCreated = fil.DateLastModified
      '    tb!DateTime = Format(Now, "dd/mm/yyyy")
      '    tb.Update
      '
      '  End If
      '  LatestFileName = UCase$(Dir)
      'Loop


      'Find the LATEST version from the list of EXEs
      'The most recent file DateCreated time should be the latest file
440   sql = "Select * from VersionControl order by File_DateCreated desc"

450   Set tb = New Recordset
460   RecOpenServerBB 0, tb, sql

      'Display most recent EXE
      'If Not tb.EOF Then
      '    lblLatestVersion = ucase$(Trim$(tb!LatestFileName))
      'End If
    
      'Display old EXEs
      'LatestFileName - File Version - Deployed (Y/N) - Date Created
470   Do While Not tb.EOF
480     s = tb!FileName & vbTab & _
            tb!File_Version & vbTab & _
            IIf(tb!Deployed = True, "Yes", "No") & vbTab & _
            Format(tb!File_DateCreated, "dd/mm/yyyy hh:mm:ss")
490     grdVersions.AddItem s
500     tb.MoveNext
510   Loop

      'If Latest and Active version the same THEN change screen display
      'If Trim$(lblLatestVersion.Caption) = Trim$(lblActiveVersion.Caption) Then
      '    fraLat2.Visible = False
      '    fraActive.Caption = fraActive.Caption & " and " & fraLatest.Caption
      'End If

520   If grdVersions.TextMatrix(1, 0) = "" And grdVersions.Rows > 2 Then
530       grdVersions.RemoveItem 1
540   End If
    
      'Highlight clicked row
      'grdVersions.Row = 1
      'For c = 0 To grdVersions.Cols - 1
      '    grdVersions.Col = c
      '    grdVersions.CellBackColor = &H80FFFF
      'Next
      '
      'AlignGridCols

550   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "frmVersionControl", "Form_Load", intEL, strES, sql


End Sub

Private Sub grdVersions_Click()

      Dim intNewRowSelected As Integer
      Dim R As Integer
      Dim intRowSelected As Integer
      Dim C As Integer

      'Remember row clicked
10    intNewRowSelected = grdVersions.Row
20    grdVersions.Col = 0

30    If grdVersions = "" Then Exit Sub

      'What row is currently hightlighted
40    grdVersions.Col = 1
50    For R = 0 To grdVersions.Rows - 1
60        grdVersions.Row = R
70        If grdVersions.CellBackColor = &H80FFFF Then
80            intRowSelected = R
90            Exit For
100       Else
110           intRowSelected = 1
120       End If
130   Next

      'Clear Row selected already

140   grdVersions.Row = intRowSelected
150   grdVersions.Col = 0

160   For C = 0 To grdVersions.Cols - 1
170       grdVersions.Col = C
180       grdVersions.CellBackColor = 0
190   Next

      'Highlight clicked row
200   grdVersions.Row = intNewRowSelected
210   For C = 0 To grdVersions.Cols - 1
220       grdVersions.Col = C
230       grdVersions.CellBackColor = &H80FFFF
240   Next

250   lblLatestVersion = grdVersions.TextMatrix(intNewRowSelected, 0)
260   lblLatestVerNumber = grdVersions.TextMatrix(intNewRowSelected, 1)
270   lblLatestDate = grdVersions.TextMatrix(intNewRowSelected, 3)

280   If intNewRowSelected = 1 Then
290       fraLatest.Caption = "Latest Version"
300   Else
310       fraLatest.Caption = "Old Version"
320   End If

330   fraLat2.Visible = True
    
End Sub
