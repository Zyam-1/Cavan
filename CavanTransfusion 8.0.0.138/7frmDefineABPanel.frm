VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDefineABPanel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Define Antibody Panel"
   ClientHeight    =   7050
   ClientLeft      =   210
   ClientTop       =   690
   ClientWidth     =   12360
   ControlBox      =   0   'False
   Icon            =   "7frmDefineABPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid tempFlexGrid 
      Height          =   375
      Left            =   540
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   765
      Left            =   8520
      Picture         =   "7frmDefineABPanel.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   90
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   765
      Left            =   7500
      Picture         =   "7frmDefineABPanel.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   90
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   11010
      Picture         =   "7frmDefineABPanel.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Panel"
      Height          =   765
      Left            =   9660
      Picture         =   "7frmDefineABPanel.frx":18A8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   90
      Width           =   1245
   End
   Begin VB.ComboBox cmbLotNumber 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Text            =   "cLotNumber"
      Top             =   300
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker dtExpiry 
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   300
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtIssued 
      Height          =   315
      Left            =   330
      TabIndex        =   6
      Top             =   300
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin VB.TextBox txtSupplier 
      Height          =   315
      Left            =   4650
      MaxLength       =   15
      TabIndex        =   1
      Top             =   300
      Width           =   2745
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5445
      Left            =   240
      TabIndex        =   0
      Top             =   1230
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   9604
      _Version        =   393216
      Rows            =   21
      Cols            =   51
      FixedCols       =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   12648384
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   1
      FormatString    =   "<Donor # |<ABO |<Rh=Hr |^Cell |     "
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   30
      TabIndex        =   12
      Top             =   6780
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8520
      TabIndex        =   14
      Top             =   900
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Panel No"
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   90
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Issued Date"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Expiry Date"
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   3
      Top             =   90
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Supplier:-"
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   90
      Width           =   660
   End
End
Attribute VB_Name = "frmDefineABPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private mLotNumber As String

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdImport_Click()

10    frmImportDiamed.Show 1

20    LoadPanel

End Sub

Private Sub LoadPanel()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim Y As Integer
      Dim Position As Integer
      Dim Pattern() As String
10    On Error GoTo LoadPanel_Error

20    txtSupplier = ""

30    g.Rows = 2
40    g.AddItem ""
50    g.RemoveItem 1
60    g.Rows = 21

70    s = Trim$(cmbLotNumber)
80    If s = "" Then
90      iMsg "Specify Panel Number.", vbExclamation
100     If TimedOut Then Unload Me: Exit Sub
110     Exit Sub
120   End If

130   sql = "Select * from AntibodyPanels where " & _
            "LotNumber = '" & cmbLotNumber & "'"

140   Set tb = New Recordset
150   RecOpenServerBB 0, tb, sql

160   If Not tb.EOF Then
170     dtIssued = Format(tb!IssuedDate, "dd/mm/yyyy")
180     dtExpiry = Format(tb!ExpiryDate, "dd/mm/yyyy")
190     txtSupplier = tb!Supplier & ""
200   End If

210   sql = "Select * from AntibodyPatterns where " & _
            "LotNumber = '" & cmbLotNumber & "' " & _
            "order by Position"
220   Set tb = New Recordset
230   RecOpenServerBB 0, tb, sql

240   Do While Not tb.EOF
250     Position = tb!Position
260     Pattern = Split(tb!Pattern, vbTab)
270     For Y = 0 To UBound(Pattern)
280       If Y < g.Rows Then
290         g.TextMatrix(Y, Position) = Pattern(Y)
300       End If
310     Next
320     If Position > 3 Then
330       g.ColWidth(Position) = TextWidth(Pattern(0) & "W")
340     End If
350     tb.MoveNext
360   Loop

370   Exit Sub

LoadPanel_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmDefineABPanel", "LoadPanel", intEL, strES, sql

End Sub



Private Sub cmdSave_Click()

      Dim s As String
      Dim Y As Integer
      Dim X As Integer
      Dim Pattern As String
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmdSave_Click_Error

20    s = Trim$(cmbLotNumber)
30    If s = "" Then
40      iMsg "Specify Panel Number.", vbExclamation
50      If TimedOut Then Unload Me: Exit Sub
60      Exit Sub
70    End If

80    If Trim$(txtSupplier) = "" Then
90      iMsg "Specify Supplier.", vbCritical
100     If TimedOut Then Unload Me: Exit Sub
110     Exit Sub
120   End If

130   For X = 0 To g.Cols - 1
140     Pattern = ""
150     For Y = 0 To g.Rows - 1
160       If g.TextMatrix(Y, X) <> "" Then
170         Pattern = Pattern & g.TextMatrix(Y, X) & vbTab
180       End If
190     Next
  
200     If Right$(Pattern, 1) = vbTab Then
210       Pattern = Left$(Pattern, Len(Pattern) - 1)
220     End If
  
230     If Trim$(Pattern) <> "" Then
240       sql = "Select * from AntibodyPanels where " & _
                "LotNumber = '" & cmbLotNumber & "'"
250       Set tb = New Recordset
260       RecOpenServerBB 0, tb, sql
270       If tb.EOF Then
280         tb.AddNew
290       End If
300       tb!LotNumber = cmbLotNumber
310       tb!IssuedDate = Format(dtIssued, "dd/mmm/yyyy")
320       tb!ExpiryDate = Format(dtExpiry, "dd/mmm/yyyy")
330       tb!Supplier = txtSupplier
340       tb!DateEntered = Format(Now, "dd/mmm/yyyy")
350       tb!EnteredBy = UserName
360       tb.Update
    
370       If Pattern <> "" Then
380         sql = "Select * from AntibodyPatterns where " & _
                  "LotNumber = '" & cmbLotNumber & "' " & _
                  "and Position = '" & X & "'"
390         Set tb = New Recordset
400         RecOpenServerBB 0, tb, sql
410         If tb.EOF Then
420           tb.AddNew
430         End If
440         tb!LotNumber = cmbLotNumber
450         tb!Position = X
460         tb!Pattern = Pattern
470         tb.Update
480       End If
490     End If
500   Next

510   cmdSave.Visible = False

520   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

530   intEL = Erl
540   strES = Err.Description
550   LogError "frmDefineABPanel", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub cmbLotNumber_Click()

10    LoadPanel

End Sub


Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Antibody Panel" & vbCr
20    strHeading = strHeading & "Issued Date : " & dtIssued & vbCr
30    strHeading = strHeading & "Expiry Date : " & dtExpiry & vbCr
40    strHeading = strHeading & "Panel Number : " & cmbLotNumber & vbCr
50    strHeading = strHeading & "Supplier : " & txtSupplier & vbCr
60    strHeading = strHeading & " " & vbCr


70    With tempFlexGrid
          Dim i As Integer, J As Integer
80        .Rows = g.Rows
90        .Cols = g.Cols
100       For i = 0 To g.Cols - 1
110           For J = 0 To g.Rows - 1
120               .TextMatrix(J, i) = g.TextMatrix(J, i)
130           Next J
140       Next i
150       For i = 0 To tempFlexGrid.Cols - 2
160           If i >= tempFlexGrid.Cols Then Exit For
170           If tempFlexGrid.TextMatrix(0, i) = "" Then
180               tempFlexGrid.ColPosition(i) = tempFlexGrid.Cols - 1
190               tempFlexGrid.Cols = tempFlexGrid.Cols - 1
200               i = i - 1
210           End If

220       Next i
230   End With
240   ExportFlexGrid tempFlexGrid, Me, strHeading
End Sub

Private Sub dtExpiry_CloseUp()

10    cmdSave.Visible = True

End Sub


Private Sub dtIssued_CloseUp()

10    cmdSave.Visible = True

End Sub




Private Sub Form_Load()


      Dim sn As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo Form_Load_Error

20    sql = "SELECT LotNumber FROM AntibodyPanels " & _
            "ORDER BY DateEntered DESC"
30    Set sn = New Recordset
40    RecOpenServerBB 0, sn, sql
50    cmbLotNumber.Clear
60    Do While Not sn.EOF
70      cmbLotNumber.AddItem sn!LotNumber & ""
80      sn.MoveNext
90    Loop
  
100   With g
110     For n = 1 To 20
120       .TextMatrix(n, 3) = Format(n)
130     Next
140     For n = 4 To g.Cols - 1
150       g.ColAlignment(n) = flexAlignCenterCenter
160     Next
170   End With

180   dtIssued = Format(Now, "dd/mmm/yyyy")
190   dtExpiry = Format(Now + 31, "dd/mmm/yyyy")


      '*****NOTE
          'This code might be dependent on many components so for any future
          'update in code try to keep this on bottom most line of form load.
200       If mLotNumber <> "" Then
210           cmbLotNumber = mLotNumber
220           LoadPanel
230         End If
      '**************************************

240   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmDefineABPanel", "Form_Load", intEL, strES, sql

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Visible Then
30      Answer = iMsg("Cancel without saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

10    mLotNumber = ""
20    Activated = False

End Sub


Private Sub g_Click()

10    g.Row = g.MouseRow
20    g.Col = g.MouseCol

30    If g.MouseRow = 0 Then
40      If g.MouseCol < 4 Then
50        Exit Sub
60      Else
70        g = iBOX("Title", , g)
80        If TimedOut Then Unload Me: Exit Sub
90      End If
100   Else
110     If g.Col > 3 Then
120       Select Case g
            Case "": g = "O"
130         Case "O": g = "+"
140         Case Else: g = ""
150       End Select
160     ElseIf g.Col <> 3 Then
170       g = iBOX("Enter " & g.TextMatrix(0, g.Col), , g)
180       If TimedOut Then Unload Me: Exit Sub
190     End If
200   End If

210   cmdSave.Visible = True

End Sub


Public Property Let Panel(ByVal LotNumber As String)

10    mLotNumber = LotNumber

End Property

