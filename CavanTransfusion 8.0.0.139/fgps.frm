VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form fgps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - G. P. Entry"
   ClientHeight    =   7710
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13995
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7710
   ScaleWidth      =   13995
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   39
      Top             =   6450
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox cmbListItems 
      Height          =   315
      ItemData        =   "fgps.frx":0000
      Left            =   9690
      List            =   "fgps.frx":0002
      TabIndex        =   37
      Top             =   7260
      Width           =   1875
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3720
      Picture         =   "fgps.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   150
      Picture         =   "fgps.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6960
      Width           =   1245
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8940
      Top             =   7110
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7560
      Top             =   7110
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   5370
      Picture         =   "fgps.frx":0BD8
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6960
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8490
      Picture         =   "fgps.frx":101A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7050
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8010
      Picture         =   "fgps.frx":299C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7050
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   705
      Left            =   12870
      Picture         =   "fgps.frx":431E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2640
      Picture         =   "fgps.frx":4988
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add GP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   13665
      Begin VB.TextBox txtMCNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11670
         TabIndex        =   35
         Top             =   810
         Width           =   1395
      End
      Begin VB.CommandButton cmdAddToPractice 
         Caption         =   "..."
         Height          =   315
         Left            =   5250
         TabIndex        =   29
         ToolTipText     =   "Add/Edit Practices"
         Top             =   270
         Width           =   405
      End
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   8580
         TabIndex        =   28
         Text            =   "cmbHospital"
         Top             =   300
         Width           =   1965
      End
      Begin VB.ComboBox cmbPractice 
         Height          =   315
         Left            =   2970
         TabIndex        =   19
         Text            =   "cmbPractice"
         Top             =   270
         Width           =   2295
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   8580
         TabIndex        =   17
         Top             =   1140
         Width           =   1965
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   8580
         TabIndex        =   16
         Top             =   810
         Width           =   1965
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   4350
         TabIndex        =   15
         Top             =   810
         Width           =   3525
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   810
         Width           =   2295
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   810
         TabIndex        =   13
         Top             =   810
         Width           =   1185
      End
      Begin VB.TextBox txtAddr1 
         Height          =   285
         Left            =   4350
         TabIndex        =   12
         Top             =   1140
         Width           =   3525
      End
      Begin VB.TextBox txtAddr0 
         Height          =   285
         Left            =   810
         TabIndex        =   11
         Top             =   1140
         Width           =   3525
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12510
         TabIndex        =   4
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   2
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "M.C.Number"
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
         Left            =   10740
         TabIndex        =   34
         Top             =   810
         Width           =   885
      End
      Begin VB.Label lblHealthlink 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11670
         TabIndex        =   33
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Healthlink"
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
         Left            =   10920
         TabIndex        =   32
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label12 
         Caption         =   "Compiled Report"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblCompiled 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   20
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Practice"
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
         Left            =   2340
         TabIndex        =   18
         Top             =   330
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
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
         Left            =   8250
         TabIndex        =   10
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
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
         Left            =   8100
         TabIndex        =   9
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Surname"
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
         Left            =   4380
         TabIndex        =   8
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Forename"
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
         Left            =   2070
         TabIndex        =   7
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Title"
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
         Left            =   810
         TabIndex        =   6
         Top             =   630
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   210
         TabIndex        =   5
         Top             =   1170
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         Left            =   390
         TabIndex        =   3
         Top             =   300
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4725
      Left            =   150
      TabIndex        =   0
      Top             =   1680
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"fgps.frx":4FF2
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
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   120
      TabIndex        =   27
      Top             =   6690
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Items to be displayed in list"
      Height          =   195
      Left            =   9510
      TabIndex        =   38
      Top             =   7020
      Width           =   2325
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   1410
      TabIndex        =   31
      Top             =   7110
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "fgps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Dim strFullName As String

Private FireCounter As Integer

Private strPractices() As String


















Private Sub cmbListItems_Click()
10    cmdSave.Visible = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
10    KeyAscii = 0
End Sub

Private Sub cmdDelete_Click()
      'check if record doesn't exist in demographics
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdDelete_Click_Error

20    If g.row = 0 Or g.Rows <= 2 Then Exit Sub

30    sql = "SELECT Count(*) as RC FROM Demographics WHERE GP = '" & g.TextMatrix(g.row, 2) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If tb!rc > 0 Then
70        iMsg "Reference to " & g.TextMatrix(g.row, 2) & " is in use so cannot be deleted"
80        Exit Sub
90    Else
100       If iMsg("Are you sure you want to delete " & g.TextMatrix(g.row, 2) & "?", vbYesNo) = vbYes Then
110           Cnxn(0).Execute "DELETE FROM gps WHERE " & _
                      "Code = '" & g.TextMatrix(g.row, 0) & "' AND Text = '" & g.TextMatrix(g.row, 2) & "'"
120           FillG
130       End If
    
140   End If

150   Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "fgps", "cmdDelete_Click", intEL, strES, sql

    
End Sub
Private Function BuildName(ByVal Title As String, _
                           ByVal ForeName As String, _
                           ByVal SurName As String) _
                           As String

      Dim s As String

10    s = ""
20    Title = Trim$(Title)
30    ForeName = Trim$(ForeName)
40    SurName = Trim$(SurName)

50    If Title <> "" Then
60      s = Title & " "
70    End If

80    If ForeName <> "" Then
90      s = s & ForeName & " "
100   End If

110   s = Trim$(s & SurName)

120   BuildName = s

End Function

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

10    If g.row = g.Rows - 1 Then Exit Sub
20    n = g.row

30    VisibleRows = g.Height \ g.RowHeight(1) - 1

40    FireCounter = FireCounter + 1
50    If FireCounter > 5 Then
60      tmrDown.Interval = 100
70    End If

80    g.Visible = False

90    s = ""
100   For X = 0 To g.Cols - 1
110     s = s & g.TextMatrix(n, X) & vbTab
120   Next
130   s = Left$(s, Len(s) - 1)

140   g.RemoveItem n
150   If n < g.Rows Then
160     g.AddItem s, n + 1
170     g.row = n + 1
180   Else
190     g.AddItem s
200     g.row = g.Rows - 1
210   End If

220   For X = 0 To g.Cols - 1
230     g.col = X
240     g.CellBackColor = vbYellow
250   Next

260   If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
270     If g.row - VisibleRows + 1 > 0 Then
280       g.TopRow = g.row - VisibleRows + 1
290     End If
300   End If

310   g.Visible = True

320   cmdSave.Visible = True

End Sub

Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

10    If g.row = 1 Then Exit Sub

20    FireCounter = FireCounter + 1
30    If FireCounter > 5 Then
40      tmrUp.Interval = 100
50    End If

60    n = g.row

70    g.Visible = False

80    s = ""
90    For X = 0 To g.Cols - 1
100     s = s & g.TextMatrix(n, X) & vbTab
110   Next
120   s = Left$(s, Len(s) - 1)

130   g.RemoveItem n
140   g.AddItem s, n - 1

150   g.row = n - 1
160   For X = 0 To g.Cols - 1
170     g.col = X
180     g.CellBackColor = vbYellow
190   Next

200   If Not g.RowIsVisible(g.row) Then
210     g.TopRow = g.row
220   End If

230   g.Visible = True

240   cmdSave.Visible = True

End Sub



Private Sub cmdAddToPractice_Click()

10    frmPractice.Show 1

20    FillPractices

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FireDown

20    tmrDown.Interval = 250
30    FireCounter = 0

40    tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FireUp

20    tmrUp.Interval = 250
30    FireCounter = 0

40    tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

           Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    On Error GoTo cmdPrint_Click_Error

20    OriginalPrinter = Printer.DeviceName

30    If Not SetFormPrinter() Then Exit Sub

40    Printer.FontName = "Courier New"
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.FontSize = 10
70    Printer.Font.Bold = True
80    Printer.Print
90    Printer.Print FormatString("List Of G. P.'s.", 99, , AlignCenter)

      '****Report body heading

100   Printer.Font.Size = 9
110   For i = 1 To 108
120       Printer.Print "-";
130   Next i
140   Printer.Print

150   Printer.Print FormatString("", 0, "|");
160   Printer.Print FormatString("In Use", 6, "|", AlignCenter);
170   Printer.Print FormatString("Code", 8, "|", AlignCenter);
180   Printer.Print FormatString("Name", 30, "|", AlignCenter);
190   Printer.Print FormatString("Address", 30, "|", AlignCenter);
200   Printer.Print FormatString("Phone", 15, "|", AlignCenter);
210   Printer.Print FormatString("MC Number", 12, "|", AlignCenter)
      '****Report body

220   Printer.Font.Bold = False

230   For i = 1 To 108
240       Printer.Print "-";
250   Next i
260   Printer.Print
270   For Y = 1 To g.Rows - 1
280       Printer.Print FormatString("", 0, "|");
290       Printer.Print FormatString(g.TextMatrix(Y, 1), 6, "|", AlignCenter);
300       Printer.Print FormatString(g.TextMatrix(Y, 0), 8, "|", Alignleft);
310       Printer.Print FormatString(g.TextMatrix(Y, 2), 30, "|", Alignleft);
320       Printer.Print FormatString(g.TextMatrix(Y, 3), 30, "|", Alignleft);
330       Printer.Print FormatString(g.TextMatrix(Y, 8), 15, "|", Alignleft);
340       Printer.Print FormatString(g.TextMatrix(Y, 13), 12, "|", Alignleft)

 
350   Next

360   Printer.EndDoc

370   For Each Px In Printers
380     If Px.DeviceName = OriginalPrinter Then
390       Set Printer = Px
400       Exit For
410     End If
420   Next

430   Exit Sub

cmdPrint_Click_Error:

Dim strES As String
Dim intEL As Integer

440   intEL = Erl
450   strES = Err.Description
460   LogError "fgps", "cmdPrint_Click", intEL, strES

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim HospitalCode As String

10    g.Rows = 2
20    g.AddItem ""
30    g.RemoveItem 1

40    HospitalCode = ListCodeFor("HO", cmbHospital)

50    sql = "SELECT Code, InUse, Text, Addr0, Addr1, Title, ForeName, SurName, Phone, FAX, Practice, " & _
            "COALESCE(Compiled, 0) AS Compiled, " & _
            "COALESCE(Healthlink, 0) AS Healthlink, MCNumber FROM GPs WHERE " & _
            "HospitalCode = '" & HospitalCode & "' " & _
            "ORDER BY ListOrder"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql

80    Do While Not tb.EOF
90      With tb
100       s = !code & vbTab & _
              IIf(!InUse, "Yes", "No") & vbTab & _
              !Text & vbTab & _
              !Addr0 & vbTab & _
              !Addr1 & vbTab & _
              !Title & vbTab & _
              !ForeName & vbTab & _
              !SurName & vbTab & _
              !Phone & vbTab & _
              !FAX & vbTab & _
              !Practice & vbTab & _
              IIf(!Compiled, "Compiled", "Full") & vbTab & _
              IIf(!Healthlink, "Yes", "No") & vbTab & _
              !MCNumber & ""
110       g.AddItem s
120     End With
130     tb.MoveNext
140   Loop

150   If g.Rows > 2 Then g.RemoveItem 1

End Sub

Private Sub cmdSave_Click()

      Dim HospitalCode As String
      Dim Y As Integer
      Dim sql As String
      Dim tb As Recordset

10    HospitalCode = ListCodeFor("HO", cmbHospital)

20    pb.max = g.Rows - 1
30    pb.Visible = True
40    cmdSave.Caption = "Saving..."

50    For Y = 1 To g.Rows - 1
60      pb = Y
70      sql = "Select * from GPs where " & _
              "Code = '" & g.TextMatrix(Y, 0) & "' " & _
              "and HospitalCode = '" & HospitalCode & "'"
80      Set tb = New Recordset
90      RecOpenClient 0, tb, sql
100     If tb.EOF Then
110       tb.AddNew
120     End If
130     With tb
140       !code = g.TextMatrix(Y, 0)
150       !InUse = IIf(g.TextMatrix(Y, 1) = "Yes", True, False)
160       !Text = g.TextMatrix(Y, 2)
170       !Addr0 = g.TextMatrix(Y, 3)
180       !Addr1 = g.TextMatrix(Y, 4)
190       !Title = Trim$(g.TextMatrix(Y, 5))
200       !ForeName = g.TextMatrix(Y, 6)
210       !SurName = g.TextMatrix(Y, 7)
220       !Phone = g.TextMatrix(Y, 8)
230       !FAX = g.TextMatrix(Y, 9)
240       !Practice = g.TextMatrix(Y, 10)
250       !Compiled = g.TextMatrix(Y, 11) = "Compiled"
260       !Healthlink = g.TextMatrix(Y, 12) = "Yes"
270       !MCNumber = g.TextMatrix(Y, 13)
280       !HospitalCode = HospitalCode
290       !ListOrder = Y
300       .Update
310     End With
320   Next

330   Call SaveOptionSetting("GPListLength", cmbListItems)

340   pb.Visible = False
350   cmdSave.Visible = False
360   cmdSave.Caption = "Save"

370   FillG

End Sub

Private Sub cmbHospital_Click()

10    FillPractices

20    FillG

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub cmdadd_Click()

      Dim strCode As String
      Dim strSurName As String
      Dim s As String

10    strCode = Trim$(UCase$(txtCode))
20    If strCode = "" Then
30      iMsg "Enter Code", vbCritical
40      Exit Sub
50    End If
60    strSurName = Trim$(txtSurname)
70    If strSurName = "" Then
80      iMsg "Enter Surname", vbCritical
90      Exit Sub
100   End If

110   s = strCode & vbTab & _
          "Yes" & vbTab & _
          strFullName & vbTab & _
          txtAddr0 & vbTab & _
          txtAddr1 & vbTab & _
          txtTitle & vbTab & _
          txtForeName & vbTab & _
          txtSurname & vbTab & _
          txtPhone & vbTab & _
          txtFAX & vbTab & _
          cmbPractice & vbTab & _
          IIf(lblCompiled = "Yes", "Compiled", "Full") & vbTab & _
          IIf(lblHealthlink = "Yes", "Yes", "No") & vbTab & _
          txtMCNumber

120   g.AddItem s

130   txtCode = ""
140   txtAddr0 = ""
150   txtAddr1 = ""
160   txtTitle = ""
170   txtForeName = ""
180   txtSurname = ""
190   txtPhone = ""
200   txtFAX = ""
210   cmbPractice = ""
220   lblCompiled = "No"
230   lblHealthlink = "No"
240   txtMCNumber = ""

250   cmdSave.Visible = True

End Sub

Private Sub cmbPractice_Click()

      Dim tb As New Recordset
      Dim sql As String

10    On Error Resume Next

20    sql = "Select * from Practices where " & _
            "Text = '" & cmbPractice.Text & "'"
30    RecOpenServer 0, tb, sql
40    If tb.EOF Then
50      txtFAX = ""
60    Else
70      txtFAX = tb!FAX & ""
80    End If

End Sub


Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

10    If Activated Then
20      Exit Sub
30    End If

40    Activated = True

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    cmbHospital.Clear
20    sql = "Select * from Lists where " & _
            "ListType = 'HO' and InUse = 1 " & _
            "order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    Do While Not tb.EOF
60      cmbHospital.AddItem tb!Text & ""
70      tb.MoveNext
80    Loop
90    For n = 0 To cmbHospital.ListCount - 1
100     If cmbHospital.List(n) = HospName(0) Then
110       cmbHospital = HospName(0)
120     End If
130   Next

140   FillPractices

150   EnsureColumnExists "GPs", "Healthlink", "bit"


160   FillG

      Dim i  As Integer
170   cmbListItems.Clear
180   For i = 8 To 32 Step 8
190       cmbListItems.AddItem i
200   Next i
210   cmbListItems.Text = GetOptionSetting("GPListLength", 8)

End Sub

Private Sub FillPractices()

      Dim sql As String
      Dim tb As Recordset
      Dim intN As Integer

10    ReDim strPractices(0 To 0)
20    strPractices(0) = ""
30    intN = 1

40    cmbPractice.Clear
50    sql = "Select * from Practices where " & _
            "Hospital = '" & cmbHospital & "'"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql

80    Do While Not tb.EOF
90      cmbPractice.AddItem tb!Text & ""

100     ReDim Preserve strPractices(0 To intN) As String
110     strPractices(intN) = tb!Text & ""
  
120     intN = intN + 1
  
130     tb.MoveNext
  
140   Loop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If cmdSave.Visible Then
20      If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
30        Cancel = True
40      End If
50    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim f As Form



10    ySave = g.row

20    If g.MouseRow = 0 Then
30      If SortOrder Then
40        g.Sort = flexSortGenericAscending
50      Else
60        g.Sort = flexSortGenericDescending
70      End If
80      SortOrder = Not SortOrder
90      cmdMoveUp.Enabled = False
100     cmdMoveDown.Enabled = False
110     cmdSave.Visible = True
120     Exit Sub
130   End If



140   g.Enabled = False
150   Select Case g.col
        Case 0
160       g.Visible = False
170       g.col = 0
180       For Y = 1 To g.Rows - 1
190           g.row = Y
200           If g.CellBackColor = vbYellow Then
210             For X = 0 To g.Cols - 1
220               g.col = X
230               g.CellBackColor = 0
240             Next
250             Exit For
260           End If
270       Next
280       g.row = ySave
290       For X = 0 To g.Cols - 1
300   g.col = X
310   g.CellBackColor = vbYellow
320       Next
330       g.Visible = True
340     Case 1, 11, 12
350       g = IIf(g = "No", "Yes", "No")
360     Case 2 'Name
370       iMsg "Name Cannot be changed." & vbCrLf & "Alter Title, ForeName or SurName columns."
380     Case 3 'Addr 1
390       g = Trim$(iBOX("Enter First line of Address", , g))
400     Case 4 'Addr 2
410       g = Trim$(iBOX("Enter Second line of Address", , g))
420     Case 5 'Title
430       g = Trim$(iBOX("Enter Title", , g))
440       g.TextMatrix(g.row, 2) = BuildName(g.TextMatrix(g.row, 5), g.TextMatrix(g.row, 6), g.TextMatrix(g.row, 7))
450     Case 6 'ForeName
460       g = Trim$(iBOX("Enter ForeName", , g))
470       g.TextMatrix(g.row, 2) = BuildName(g.TextMatrix(g.row, 5), g.TextMatrix(g.row, 6), g.TextMatrix(g.row, 7))
480     Case 7 'Surname
490       g = Trim$(iBOX("Enter SurName", , g))
500       g.TextMatrix(g.row, 2) = BuildName(g.TextMatrix(g.row, 5), g.TextMatrix(g.row, 6), g.TextMatrix(g.row, 7))
510     Case 8 'phone
520       g = Trim$(iBOX("Enter Phone", , g))
530     Case 9 'FAX
540       g = Trim$(iBOX("Enter FAX", , g))
550     Case 10 'Practice
560       Set f = New fcdrDBox
570       With f
580   .Options = strPractices
590   .Prompt = "Enter Practice for " & g.TextMatrix(g.row, 2)
600   .Show 1
610   g = .ReturnValue
620       End With
630       Unload f
640       Set f = Nothing
650     Case 13 'Medical Council Number
660       g = Trim$(iBOX("Enter Medical Council Number", , g))
670   End Select

680   If g.col <> 0 And g.col <> 2 Then
690     cmdSave.Visible = True
700   End If

710   g.Enabled = True

720   cmdMoveUp.Enabled = True
730   cmdMoveDown.Enabled = True

End Sub

Private Sub lblCompiled_Click()

10    lblCompiled = IIf(lblCompiled = "Yes", "No", "Yes")

End Sub

Private Sub lblHealthlink_Click()

10    lblHealthlink = IIf(lblHealthlink = "Yes", "No", "Yes")

End Sub


Private Sub tmrDown_Timer()

10    FireDown

End Sub

Private Sub tmrUp_Timer()

10    FireUp

End Sub


Private Sub txtCode_LostFocus()

10    txtCode = UCase$(Trim$(txtCode))

End Sub

Private Sub txtForeName_Change()

10    txtForeName = Trim$(txtForeName)

20    strFullName = txtTitle & " " & txtForeName & " " & txtSurname

End Sub

Private Sub txtSurName_Change()

10    txtSurname = Trim$(txtSurname)

20    strFullName = txtTitle & " " & txtForeName & " " & txtSurname

End Sub


Private Sub txtTitle_Change()

10    txtTitle = Trim$(txtTitle)

20    strFullName = txtTitle & " " & txtForeName & " " & txtSurname

End Sub


