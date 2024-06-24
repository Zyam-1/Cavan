VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGPs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - G. P. Entry"
   ClientHeight    =   8655
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16980
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8655
   ScaleWidth      =   16980
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbListItems 
      Height          =   315
      ItemData        =   "frmGPs.frx":0000
      Left            =   15015
      List            =   "frmGPs.frx":0002
      TabIndex        =   38
      Top             =   360
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
      Left            =   15825
      Picture         =   "frmGPs.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2970
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
      Left            =   15825
      Picture         =   "frmGPs.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   810
      Width           =   975
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   16095
      Top             =   4950
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   16095
      Top             =   4410
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   15825
      Picture         =   "frmGPs.frx":0BD8
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6435
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   555
      Left            =   15615
      Picture         =   "frmGPs.frx":101A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4920
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   15615
      Picture         =   "frmGPs.frx":299C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4350
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
      Left            =   15825
      Picture         =   "frmGPs.frx":431E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7500
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
      Left            =   15825
      Picture         =   "frmGPs.frx":4988
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2160
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
      TabIndex        =   19
      Top             =   90
      Width           =   13965
      Begin VB.TextBox txtPracticeNumber 
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
         Left            =   11640
         TabIndex        =   14
         Top             =   1200
         Width           =   1395
      End
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
         Left            =   11640
         TabIndex        =   13
         Top             =   810
         Width           =   1395
      End
      Begin VB.CommandButton cmdAddToPractice 
         Caption         =   "..."
         Height          =   315
         Left            =   5250
         TabIndex        =   2
         ToolTipText     =   "Add/Edit Practices"
         Top             =   270
         Width           =   405
      End
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   8580
         TabIndex        =   9
         Text            =   "cmbHospital"
         Top             =   300
         Width           =   1965
      End
      Begin VB.ComboBox cmbPractice 
         Height          =   315
         Left            =   2970
         TabIndex        =   1
         Text            =   "cmbPractice"
         Top             =   270
         Width           =   2295
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   8580
         TabIndex        =   11
         Top             =   1140
         Width           =   1965
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   8580
         TabIndex        =   10
         Top             =   810
         Width           =   1965
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   4350
         TabIndex        =   6
         Top             =   810
         Width           =   3525
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   810
         Width           =   2295
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   810
         Width           =   1185
      End
      Begin VB.TextBox txtAddr1 
         Height          =   285
         Left            =   4350
         TabIndex        =   8
         Top             =   1140
         Width           =   3525
      End
      Begin VB.TextBox txtAddr0 
         Height          =   285
         Left            =   810
         TabIndex        =   7
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
         Height          =   525
         Left            =   13260
         TabIndex        =   15
         Top             =   930
         Width           =   555
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   0
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Practice No."
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
         Left            =   10710
         TabIndex        =   40
         Top             =   1245
         Width           =   885
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
         Left            =   10710
         TabIndex        =   36
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
         Left            =   11640
         TabIndex        =   12
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
         Left            =   10890
         TabIndex        =   35
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Print Report"
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
         Left            =   6540
         TabIndex        =   28
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblPrintReport 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yes"
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
         TabIndex        =   3
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6525
      Left            =   90
      TabIndex        =   18
      Top             =   1695
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   11509
      _Version        =   393216
      Cols            =   18
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
      FormatString    =   $"frmGPs.frx":4FF2
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
      Left            =   150
      TabIndex        =   32
      Top             =   8235
      Visible         =   0   'False
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Items to be displayed in list"
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
      Left            =   14955
      TabIndex        =   39
      Top             =   180
      Width           =   1875
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   15720
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmGPs"
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
57470     cmdSave.Visible = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
57480     KeyAscii = 0
End Sub

Private Sub cmdDelete_Click()
          'check if record doesn't exist in demographics
          Dim tb As Recordset
          Dim sql As String

57490     On Error GoTo cmdDelete_Click_Error

57500     If g.row = 0 Or g.Rows <= 2 Then Exit Sub

57510     sql = "SELECT Count(*) as RC FROM Demographics WHERE GP = '" & g.TextMatrix(g.row, 2) & "'"
57520     Set tb = New Recordset
57530     RecOpenClient 0, tb, sql
57540     If tb!rc > 0 Then
57550         iMsg "Reference to " & g.TextMatrix(g.row, 2) & " is in use so cannot be deleted"
57560         Exit Sub
57570     Else
57580         If iMsg("Are you sure you want to delete " & g.TextMatrix(g.row, 2) & "?", vbYesNo) = vbYes Then
57590             Cnxn(0).Execute "DELETE FROM gps WHERE " & _
                      "Code = '" & g.TextMatrix(g.row, 0) & "' AND Text = '" & g.TextMatrix(g.row, 2) & "'"
57600             FillG
57610         End If
          
57620     End If

57630     Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

57640     intEL = Erl
57650     strES = Err.Description
57660     LogError "fgps", "cmdDelete_Click", intEL, strES, sql
          
End Sub
Private Function BuildName(ByVal Title As String, _
          ByVal ForeName As String, _
          ByVal SurName As String) _
          As String

          Dim s As String

57670     s = ""
57680     Title = Trim$(Title)
57690     ForeName = Trim$(ForeName)
57700     SurName = Trim$(SurName)

57710     If Title <> "" Then
57720         s = Title & " "
57730     End If

57740     If ForeName <> "" Then
57750         s = s & ForeName & " "
57760     End If

57770     s = Trim$(s & SurName)

57780     BuildName = s

End Function

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

57790     If g.row = g.Rows - 1 Then Exit Sub
57800     n = g.row

57810     VisibleRows = g.height \ g.RowHeight(1) - 1

57820     FireCounter = FireCounter + 1
57830     If FireCounter > 5 Then
57840         tmrDown.Interval = 100
57850     End If

57860     g.Visible = False

57870     s = ""
57880     For X = 0 To g.Cols - 1
57890         s = s & g.TextMatrix(n, X) & vbTab
57900     Next
57910     s = Left$(s, Len(s) - 1)

57920     g.RemoveItem n
57930     If n < g.Rows Then
57940         g.AddItem s, n + 1
57950         g.row = n + 1
57960     Else
57970         g.AddItem s
57980         g.row = g.Rows - 1
57990     End If

58000     For X = 0 To g.Cols - 1
58010         g.Col = X
58020         g.CellBackColor = vbYellow
58030     Next

58040     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
58050         If g.row - VisibleRows + 1 > 0 Then
58060             g.TopRow = g.row - VisibleRows + 1
58070         End If
58080     End If

58090     g.Visible = True

58100     cmdSave.Visible = True

End Sub

Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

58110     If g.row = 1 Then Exit Sub

58120     FireCounter = FireCounter + 1
58130     If FireCounter > 5 Then
58140         tmrUp.Interval = 100
58150     End If

58160     n = g.row

58170     g.Visible = False

58180     s = ""
58190     For X = 0 To g.Cols - 1
58200         s = s & g.TextMatrix(n, X) & vbTab
58210     Next
58220     s = Left$(s, Len(s) - 1)

58230     g.RemoveItem n
58240     g.AddItem s, n - 1

58250     g.row = n - 1
58260     For X = 0 To g.Cols - 1
58270         g.Col = X
58280         g.CellBackColor = vbYellow
58290     Next

58300     If Not g.RowIsVisible(g.row) Then
58310         g.TopRow = g.row
58320     End If

58330     g.Visible = True

58340     cmdSave.Visible = True

End Sub



Private Sub cmdAddToPractice_Click()

58350     frmPractice.Show 1

58360     FillPractices

End Sub

Private Sub cmdCancel_Click()

58370     Unload Me

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

58380     FireDown

58390     tmrDown.Interval = 250
58400     FireCounter = 0

58410     tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

58420     tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

58430     FireUp

58440     tmrUp.Interval = 250
58450     FireCounter = 0

58460     tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

58470     tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

          Dim Y As Integer
          Dim Px As Printer
          Dim OriginalPrinter As String
          Dim i As Integer

58480     On Error GoTo cmdPrint_Click_Error

58490     Screen.MousePointer = vbHourglass

58500     OriginalPrinter = Printer.DeviceName

58510     If Not SetFormPrinter() Then Exit Sub

58520     Printer.FontName = "Courier New"
58530     Printer.Orientation = vbPRORPortrait

          '****Report heading
58540     Printer.FontSize = 10
58550     Printer.Font.Bold = True
58560     Printer.Print
58570     Printer.Print FormatString("List Of G. P.'s.", 99, , AlignCenter)

          '****Report body heading

58580     Printer.Font.size = 9
58590     For i = 1 To 108
58600         Printer.Print "-";
58610     Next i
58620     Printer.Print


58630     Printer.Print FormatString("", 0, "|");
58640     Printer.Print FormatString("In Use", 6, "|", AlignCenter);
58650     Printer.Print FormatString("Code", 8, "|", AlignCenter);
58660     Printer.Print FormatString("Name", 30, "|", AlignCenter);
58670     Printer.Print FormatString("Address", 30, "|", AlignCenter);
58680     Printer.Print FormatString("Phone", 15, "|", AlignCenter);
58690     Printer.Print FormatString("MC Number", 12, "|", AlignCenter)
          '****Report body

58700     Printer.Font.Bold = False

58710     For i = 1 To 108
58720         Printer.Print "-";
58730     Next i
58740     Printer.Print
58750     For Y = 1 To g.Rows - 1
58760         Printer.Print FormatString("", 0, "|");
58770         Printer.Print FormatString(g.TextMatrix(Y, 1), 6, "|", AlignCenter);
58780         Printer.Print FormatString(g.TextMatrix(Y, 0), 8, "|", AlignLeft);
58790         Printer.Print FormatString(g.TextMatrix(Y, 2), 30, "|", AlignLeft);
58800         Printer.Print FormatString(g.TextMatrix(Y, 3), 30, "|", AlignLeft);
58810         Printer.Print FormatString(g.TextMatrix(Y, 8), 15, "|", AlignLeft);
58820         Printer.Print FormatString(g.TextMatrix(Y, 13), 12, "|", AlignLeft)
58830     Next

58840     Printer.EndDoc

58850     Screen.MousePointer = vbDefault

58860     For Each Px In Printers
58870         If Px.DeviceName = OriginalPrinter Then
58880             Set Printer = Px
58890             Exit For
58900         End If
58910     Next

58920     Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

58930     intEL = Erl
58940     strES = Err.Description
58950     LogError "frmGPs", "cmdPrint_Click", intEL, strES

End Sub

Private Sub FillG()

          Dim s As String
          Dim Lx As New List
          Dim GXs As New GPs
          Dim Gx As GP

58960     On Error GoTo FillG_Error

58970     g.Rows = 2
58980     g.AddItem ""
58990     g.RemoveItem 1

59000     Lx.GetCode "HO", cmbHospital

59010     GXs.Load Lx.Code, False
59020     For Each Gx In GXs
59030         s = Gx.Code & vbTab & _
                  IIf(Gx.InUse, "Yes", "No") & vbTab & _
                  Gx.Text & vbTab & _
                  Gx.Addr0 & vbTab & _
                  Gx.Addr1 & vbTab & _
                  Gx.Title & vbTab & _
                  Gx.ForeName & vbTab & _
                  Gx.SurName & vbTab & _
                  Gx.Phone & vbTab & _
                  Gx.FAX & vbTab & _
                  Gx.Practice & vbTab & _
                  IIf(Gx.PrintReport, "Yes", "No") & vbTab & _
                  IIf(Gx.HealthLink, "Yes", "No") & vbTab & _
                  Gx.McNumber & vbTab & _
                  Gx.PracticeNumber & vbTab & _
                  IIf(Gx.EGFR, "Yes", "No") & vbTab & _
                  IIf(Gx.AutoCC, "Yes", "No") & vbTab & _
                  IIf(Gx.Interim, "Yes", "No")
59040         g.AddItem s
59050     Next

59060     If g.Rows > 2 Then g.RemoveItem 1

59070     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

59080     intEL = Erl
59090     strES = Err.Description
59100     LogError "frmGPs", "FillG", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim HospitalCode As String
          Dim Y As Integer
          Dim sql As String
          Dim tb As Recordset

59110     On Error GoTo cmdSave_Click_Error

59120     HospitalCode = ListCodeFor("HO", cmbHospital)

59130     pb.max = g.Rows - 1
59140     pb.Visible = True
59150     cmdSave.Caption = "Saving..."

59160     For Y = 1 To g.Rows - 1
59170         pb = Y
59180         sql = "Select * from GPs where " & _
                  "Code = '" & g.TextMatrix(Y, 0) & "' " & _
                  "and HospitalCode = '" & HospitalCode & "'"
59190         Set tb = New Recordset
59200         RecOpenClient 0, tb, sql
59210         If tb.EOF Then
59220             tb.AddNew
59230         End If
59240         With tb
59250             !Code = g.TextMatrix(Y, 0)
59260             !InUse = IIf(g.TextMatrix(Y, 1) = "Yes", True, False)
59270             !Text = g.TextMatrix(Y, 2)
59280             !Addr0 = g.TextMatrix(Y, 3)
59290             !Addr1 = g.TextMatrix(Y, 4)
59300             !Title = Trim$(g.TextMatrix(Y, 5))
59310             !ForeName = g.TextMatrix(Y, 6)
59320             !SurName = g.TextMatrix(Y, 7)
59330             !Phone = g.TextMatrix(Y, 8)
59340             !FAX = g.TextMatrix(Y, 9)
59350             !Practice = g.TextMatrix(Y, 10)
59360             !PrintReport = g.TextMatrix(Y, 11) = "Yes"
59370             !HealthLink = g.TextMatrix(Y, 12) = "Yes"
59380             !McNumber = g.TextMatrix(Y, 13)
59390             !PracticeNumber = g.TextMatrix(Y, 14)
59400             !HospitalCode = HospitalCode
59410             !ListOrder = Y
59420             !AutoCC = g.TextMatrix(Y, 16) = "Yes"
59430             !Interim = g.TextMatrix(Y, 17) = "Yes"
59440             .Update

59450             sql = "IF EXISTS (SELECT * FROM IncludeEGFR " & _
                      "           WHERE SourceType = 'GP' " & _
                      "           AND Hospital = '" & cmbHospital & "' " & _
                      "           AND SourceName = '" & AddTicks(g.TextMatrix(Y, 2)) & "' ) " & _
                      "    UPDATE IncludeEGFR " & _
                      "    SET Include = '" & IIf(g.TextMatrix(Y, 14) = "Yes", 1, 0) & "' " & _
                      "    WHERE SourceType = 'GP' " & _
                      "    AND SourceName = '" & AddTicks(g.TextMatrix(Y, 2)) & "' " & _
                      "ELSE " & _
                      "    INSERT INTO IncludeEGFR (SourceType, Hospital, SourceName, Include) " & _
                      "    VALUES ('GP', " & _
                      "            '" & cmbHospital & "', " & _
                      "            '" & AddTicks(g.TextMatrix(Y, 2)) & "', " & _
                      "            '" & IIf(g.TextMatrix(Y, 14) = "Yes", 1, 0) & "')"
59460             Cnxn(0).Execute sql

59470         End With
59480     Next

59490     Call SaveOptionSetting("GPListLength", cmbListItems)

59500     pb.Visible = False
59510     cmdSave.Visible = False
59520     cmdSave.Caption = "Save"

59530     FillG

59540     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

59550     intEL = Erl
59560     strES = Err.Description
59570     LogError "fgps", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmbHospital_Click()

59580     FillPractices

59590     FillG

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

59600     KeyAscii = 0

End Sub


Private Sub cmdAdd_Click()

          Dim strCode As String
          Dim strSurName As String
          Dim s As String

59610     On Error GoTo cmdAdd_Click_Error

59620     strCode = Trim$(UCase$(txtCode))
59630     If strCode = "" Then
59640         iMsg "Enter Code", vbCritical
59650         Exit Sub
59660     End If
59670     strSurName = Trim$(txtSurName)
59680     If strSurName = "" Then
59690         iMsg "Enter Surname", vbCritical
59700         Exit Sub
59710     End If

59720     s = strCode & vbTab & _
              "Yes" & vbTab & _
              strFullName & vbTab & _
              txtAddr0 & vbTab & _
              txtAddr1 & vbTab & _
              txtTitle & vbTab & _
              txtForeName & vbTab & _
              txtSurName & vbTab & _
              txtPhone & vbTab & _
              txtFAX & vbTab & _
              cmbPractice & vbTab & _
              lblPrintReport & vbTab & _
              IIf(lblHealthlink = "Yes", "Yes", "No") & vbTab & _
              txtMCNumber & vbTab & _
              txtPracticeNumber

59730     g.AddItem s

59740     txtCode = ""
59750     txtAddr0 = ""
59760     txtAddr1 = ""
59770     txtTitle = ""
59780     txtForeName = ""
59790     txtSurName = ""
59800     txtPhone = ""
59810     txtFAX = ""
59820     cmbPractice = ""
59830     lblPrintReport = "Yes"
59840     lblHealthlink = "No"
59850     txtMCNumber = ""
59860     txtPracticeNumber = ""

59870     cmdSave.Visible = True

59880     Exit Sub

cmdAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

59890     intEL = Erl
59900     strES = Err.Description
59910     LogError "fgps", "cmdadd_Click", intEL, strES

End Sub

Private Sub cmbPractice_Click()

          Dim tb As New Recordset
          Dim sql As String

59920     On Error GoTo cmbPractice_Click_Error

59930     sql = "SELECT FAX FROM Practices WHERE " & _
              "Text = '" & cmbPractice.Text & "'"
59940     RecOpenServer 0, tb, sql
59950     If tb.EOF Then
59960         txtFAX = ""
59970     Else
59980         txtFAX = tb!FAX & ""
59990     End If

60000     Exit Sub

cmbPractice_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60010     intEL = Erl
60020     strES = Err.Description
60030     LogError "fgps", "cmbPractice_Click", intEL, strES, sql

End Sub


Private Sub cmdXL_Click()

60040     ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

60050     If Activated Then
60060         Exit Sub
60070     End If

60080     Activated = True

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

60090     On Error GoTo Form_Load_Error

60100     cmbHospital.Clear
60110     sql = "SELECT [Text] FROM Lists WHERE " & _
              "ListType = 'HO' " & _
              "AND InUse = 1 " & _
              "ORDER BY ListOrder"
60120     Set tb = New Recordset
60130     RecOpenServer 0, tb, sql
60140     Do While Not tb.EOF
60150         cmbHospital.AddItem tb!Text & ""
60160         tb.MoveNext
60170     Loop
60180     For n = 0 To cmbHospital.ListCount - 1
60190         If UCase$(cmbHospital.List(n)) = UCase$(HospName(0)) Then
60200             cmbHospital = HospName(0)
60210         End If
60220     Next

60230     FillPractices

60240     EnsureColumnExists "GPs", "Healthlink", "bit"

60250     FillG

          Dim i  As Integer
60260     cmbListItems.Clear
60270     For i = 8 To 32 Step 8
60280         cmbListItems.AddItem i
60290     Next i
60300     cmbListItems.Text = GetOptionSetting("GPListLength", 8)

60310     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60320     intEL = Erl
60330     strES = Err.Description
60340     LogError "fgps", "Form_Load", intEL, strES, sql

End Sub

Private Sub FillPractices()

          Dim sql As String
          Dim tb As Recordset
          Dim intN As Integer

60350     On Error GoTo FillPractices_Error

60360     ReDim strPractices(0 To 0)
60370     strPractices(0) = ""
60380     intN = 1

60390     cmbPractice.Clear
60400     sql = "Select * from Practices where " & _
              "Hospital = '" & cmbHospital & "'"
60410     Set tb = New Recordset
60420     RecOpenServer 0, tb, sql

60430     Do While Not tb.EOF
60440         cmbPractice.AddItem tb!Text & ""

60450         ReDim Preserve strPractices(0 To intN) As String
60460         strPractices(intN) = tb!Text & ""
        
60470         intN = intN + 1
        
60480         tb.MoveNext
        
60490     Loop

60500     Exit Sub

FillPractices_Error:

          Dim strES As String
          Dim intEL As Integer

60510     intEL = Erl
60520     strES = Err.Description
60530     LogError "fgps", "FillPractices", intEL, strES, sql

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

60540     If cmdSave.Visible Then
60550         If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
60560             Cancel = True
60570         End If
60580     End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

60590     Activated = False

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X          As Integer
          Dim Y          As Integer
          Dim ySave      As Integer
          Dim f          As Form

60600     On Error GoTo g_Click_Error

60610     ySave = g.row

60620     If g.MouseRow = 0 Then
60630         If SortOrder Then
60640             g.Sort = flexSortGenericAscending
60650         Else
60660             g.Sort = flexSortGenericDescending
60670         End If
60680         SortOrder = Not SortOrder
60690         cmdMoveUp.Enabled = False
60700         cmdMoveDown.Enabled = False
60710         cmdSave.Visible = True
60720         Exit Sub
60730     End If

60740     g.Enabled = False
60750     Select Case g.Col
              Case 0
60760             g.Visible = False
60770             g.Col = 0
60780             For Y = 1 To g.Rows - 1
60790                 g.row = Y
60800                 If g.CellBackColor = vbYellow Then
60810                     For X = 0 To g.Cols - 1
60820                         g.Col = X
60830                         g.CellBackColor = 0
60840                     Next
60850                     Exit For
60860                 End If
60870             Next
60880             g.row = ySave
60890             For X = 0 To g.Cols - 1
60900                 g.Col = X
60910                 g.CellBackColor = vbYellow
60920             Next
60930             g.Visible = True
60940         Case 1, 11, 12, 15, 16, 17
60950             g = IIf(g = "No", "Yes", "No")
60960         Case 2    'Name
60970             iMsg "Name Cannot be changed." & vbCrLf & "Alter Title, ForeName or SurName columns."
60980         Case 3    'Addr 1
60990             g = Trim$(iBOX("Enter First line of Address", , g))
61000         Case 4    'Addr 2
61010             g = Trim$(iBOX("Enter Second line of Address", , g))
61020         Case 5    'Title
61030             g = Trim$(iBOX("Enter Title", , g))
61040             g.TextMatrix(g.row, 2) = BuildName(g.TextMatrix(g.row, 5), g.TextMatrix(g.row, 6), g.TextMatrix(g.row, 7))
61050         Case 6    'ForeName
61060             g = Trim$(iBOX("Enter ForeName", , g))
61070             g.TextMatrix(g.row, 2) = BuildName(g.TextMatrix(g.row, 5), g.TextMatrix(g.row, 6), g.TextMatrix(g.row, 7))
61080         Case 7    'Surname
61090             g = Trim$(iBOX("Enter SurName", , g))
61100             g.TextMatrix(g.row, 2) = BuildName(g.TextMatrix(g.row, 5), g.TextMatrix(g.row, 6), g.TextMatrix(g.row, 7))
61110         Case 8    'phone
61120             g = Trim$(iBOX("Enter Phone", , g))
61130         Case 9    'FAX
61140             g = Trim$(iBOX("Enter FAX", , g))
61150         Case 10    'Practice
61160             Set f = New fcdrDBox
61170             With f
61180                 .Options = strPractices
61190                 .Prompt = "Enter Practice for " & g.TextMatrix(g.row, 2)
61200                 .Show 1
61210                 g = .ReturnValue
61220             End With
61230             Unload f
61240             Set f = Nothing
61250         Case 13    'Medical Council Number
61260             g = Trim$(iBOX("Enter Medical Council Number", , g))
61270         Case 14    'Practice Number
61280             g = Trim$(iBOX("Enter Practice Number", , g))
61290     End Select

61300     If g.Col <> 0 And g.Col <> 2 Then
61310         cmdSave.Visible = True
61320     End If

61330     g.Enabled = True

61340     cmdMoveUp.Enabled = True
61350     cmdMoveDown.Enabled = True

61360     Exit Sub

g_Click_Error:

          Dim strES      As String
          Dim intEL      As Integer

61370     intEL = Erl
61380     strES = Err.Description
61390     LogError "fgps", "g_Click", intEL, strES

End Sub

Private Sub lblPrintReport_Click()

61400     lblPrintReport = IIf(lblPrintReport = "Yes", "No", "Yes")

End Sub

Private Sub lblHealthlink_Click()

61410     lblHealthlink = IIf(lblHealthlink = "Yes", "No", "Yes")

End Sub


Private Sub tmrDown_Timer()

61420     FireDown

End Sub

Private Sub tmrUp_Timer()

61430     FireUp

End Sub


Private Sub txtCode_LostFocus()

61440     txtCode = UCase$(Trim$(txtCode))

End Sub

Private Sub txtForeName_Change()

61450     txtForeName = Trim$(txtForeName)

61460     strFullName = txtTitle & " " & txtForeName & " " & txtSurName

End Sub

Private Sub txtSurName_Change()

61470     txtSurName = Trim$(txtSurName)

61480     strFullName = txtTitle & " " & txtForeName & " " & txtSurName

End Sub


Private Sub txtTitle_Change()

61490     txtTitle = Trim$(txtTitle)

61500     strFullName = txtTitle & " " & txtForeName & " " & txtSurName

End Sub


