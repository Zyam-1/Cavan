VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmASOT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Add Result"
   ClientHeight    =   8670
   ClientLeft      =   105
   ClientTop       =   360
   ClientWidth     =   12195
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8670
   ScaleWidth      =   12195
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10230
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10470
      TabIndex        =   9
      Top             =   1410
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   11070
      Picture         =   "frmASOT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7410
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   1605
      Begin VB.OptionButton view 
         Caption         =   "&All"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   210
         Width           =   495
      End
      Begin VB.OptionButton view 
         Caption         =   "&Incomplete"
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   420
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton view 
         Caption         =   "&Tagged"
         Height          =   225
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   630
         Width           =   885
      End
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   218955777
      CurrentDate     =   37082
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   1065
      Left            =   11070
      Picture         =   "frmASOT.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   7365
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   12991
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmASOT.frx":1D94
   End
   Begin VB.CommandButton bSave 
      Caption         =   "&Save"
      Height          =   1065
      Left            =   11070
      Picture         =   "frmASOT.frx":1E23
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   885
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   8100
      TabIndex        =   17
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   6750
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   5430
      TabIndex        =   15
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label lblLotNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   8100
      TabIndex        =   14
      Top             =   570
      Width           =   1275
   End
   Begin VB.Label lblLotNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   6750
      TabIndex        =   13
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label lblLotNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   5430
      TabIndex        =   12
      Top             =   570
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Card Lot Numbers / Expiry"
      Height          =   255
      Left            =   5430
      TabIndex        =   11
      Top             =   300
      Width           =   3945
   End
End
Attribute VB_Name = "frmASOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private DataChanged As Boolean

Private Enum ColumnType
    ctBoolean
    ctText
    ctCombo
    ctNoEdits
End Enum

'ColumnEditType(0 To grid.cols-1)
Private ColumnEditType(0 To 8) As ColumnType

Private Sub CheckForLotNumbers()

          Dim Y As Long
          Dim X As Long
          Dim tb As Recordset
          Dim sql As String
          Dim Analyte As String
          Dim SampleID As String
        
53760     On Error GoTo CheckForLotNumbers_Error

53770     For X = 0 To 2
53780         lblLotNumber(X) = ""
53790         lblExpiry(X) = ""
53800     Next

53810     If grd.Rows = 2 And grd.TextMatrix(1, 0) = "" Then Exit Sub 'Grid is empty

53820     For X = 5 To 7
53830         For Y = 1 To grd.Rows - 1
53840             If grd.TextMatrix(Y, X) <> "" Then
53850                 Analyte = Choose(X - 4, "Monospot", "Malaria", "Sickledex")
53860                 SampleID = Val(grd.TextMatrix(Y, 0))
53870                 sql = "Select LotNumber, Expiry from ReagentLotNumbers where " & _
                          "Analyte = '" & Analyte & "' " & _
                          "and SampleID = " & SampleID
53880                 Set tb = New Recordset
53890                 RecOpenServer 0, tb, sql
53900                 If Not tb.EOF Then
53910                     lblLotNumber(X - 5) = tb!LotNumber & ""
53920                     lblExpiry(X - 5) = Format(tb!Expiry, "dd/mm/yyyy")
53930                     Exit For
53940                 End If
53950             End If
53960         Next
53970     Next

53980     Exit Sub

CheckForLotNumbers_Error:

          Dim strES As String
          Dim intEL As Integer

53990     intEL = Erl
54000     strES = Err.Description
54010     LogError "fasot", "CheckForLotNumbers", intEL, strES, sql


End Sub

Private Sub GrdEdit(ByVal KeyAscii As Integer)

54020     Select Case KeyAscii

              Case 0 To 32
54030             txtInput = grd
54040             txtInput.SelStart = 1000
          
54050         Case Else
54060             txtInput = Chr$(KeyAscii)
54070             txtInput.SelStart = 1
54080     End Select

54090     With grd
        
54100         txtInput.Move .Left + .CellLeft, _
                  .Top + .CellTop, _
                  .CellWidth, _
                  .CellHeight
54110     End With

54120     txtInput.Visible = True
54130     txtInput.SetFocus

End Sub

Private Sub SaveDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim X As Long
          Dim Analyte As String
          Dim LotNumber As String

54140     On Error GoTo SaveDetails_Error

54150     Screen.MousePointer = 11

54160     For n = 1 To grd.Rows - 1
54170         sql = "Select * from HaemResults where " & _
                  "SampleID = '" & Val(grd.TextMatrix(n, 0)) & "'"
54180         Set tb = New Recordset
54190         RecOpenServer 0, tb, sql
54200         If Not tb.EOF Then
54210             tb!ESR = Left$(Trim$(grd.TextMatrix(n, 3)), 5)
54220             tb!RetA = Left$(Trim$(grd.TextMatrix(n, 4)), 5)
54230             tb!MonoSpot = Left$(grd.TextMatrix(n, 5), 1)
54240             tb!Malaria = Left$(grd.TextMatrix(n, 6), 8)
54250             tb!Sickledex = Left$(grd.TextMatrix(n, 7), 8)
54260             tb!RA = UCase$(Left$(grd.TextMatrix(n, 8), 1))
54270             tb.Update
54280         End If

54290         For X = 5 To 7
54300             If grd.TextMatrix(n, X) <> "" Then
54310                 LotNumber = lblLotNumber(X - 5)
54320                 If Trim$(LotNumber) <> "" Then
54330                     Analyte = Choose(X - 4, "Monospot", "Malaria", "Sickledex")
54340                     sql = "Select * from ReagentLotNumbers where " & _
                              "Analyte = '" & Analyte & "' " & _
                              "and SampleID = " & Val(grd.TextMatrix(n, 0))
54350                     Set tb = New Recordset
54360                     RecOpenServer 0, tb, sql
54370                     If tb.EOF Then tb.AddNew
54380                     tb!LotNumber = LotNumber
54390                     tb!Expiry = Format$(lblExpiry(X - 5), "dd/mmm/yyyy")
54400                     tb!Analyte = Analyte
54410                     tb!EntryDateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
54420                     tb!SampleID = Val(grd.TextMatrix(n, 0))
54430                     tb.Update
54440                 End If
54450             End If
54460         Next
        
54470     Next

54480     Screen.MousePointer = 0

54490     Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

54500     intEL = Erl
54510     strES = Err.Description
54520     LogError "fasot", "SaveDetails", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

54530     Unload Me

End Sub


Private Sub bSave_Click()

54540     If txtInput.Visible Then
54550         grd = txtInput
54560         txtInput = ""
54570         txtInput.Visible = False
54580     End If

54590     SaveDetails
54600     DataChanged = False

End Sub
Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

54610     On Error GoTo FillGrid_Error

54620     Screen.MousePointer = vbHourglass

54630     grd.Rows = 2
54640     grd.AddItem ""
54650     grd.RemoveItem 1
54660     grd.Visible = False

54670     sql = "SELECT D.Chart, D.PatName, H.SampleID, " & _
              "resESR = CASE cESR " & _
              "         WHEN 1 THEN " & _
              "                CASE COALESCE(ESR, '') " & _
              "                WHEN '' THEN '?' " & _
              "                ELSE ESR " & _
              "                END " & _
              "         ELSE '' END, " & _
              "resRetA = CASE cRetics " & _
              "         WHEN 1 THEN " & _
              "                CASE COALESCE(RetA, '') " & _
              "                WHEN '' THEN '?' " & _
              "                ELSE RetA " & _
              "                END " & _
              "         ELSE '' END, "
54680     sql = sql & "resMonoSpot = CASE cMonoSpot WHEN 1 THEN " & _
              "         CASE COALESCE(MonoSpot, '') " & _
              "         WHEN '' THEN '?' " & _
              "         WHEN 'N' THEN 'Negative' " & _
              "         WHEN 'P' THEN 'Positive' " & _
              "         ELSE MonoSpot END " & _
              "         ELSE '' END, "
54690     sql = sql & "resMalaria = CASE cMalaria WHEN 1 THEN " & _
              "         CASE COALESCE(Malaria, '') " & _
              "         WHEN '' THEN '?' " & _
              "         ELSE Malaria END " & _
              "         ELSE '' END, "
54700     sql = sql & "resSickledex = CASE cSickledex WHEN 1 THEN " & _
              "         CASE COALESCE(Sickledex, '') " & _
              "         WHEN '' THEN '?' " & _
              "         ELSE Sickledex END " & _
              "         ELSE '' END, "
54710     sql = sql & "resRA = CASE cRA WHEN 1 THEN " & _
              "         CASE COALESCE(RA, '') " & _
              "         WHEN '' THEN '?' " & _
              "         WHEN 'N' THEN 'Negative' " & _
              "         WHEN 'P' THEN 'Positive' " & _
              "         ELSE RA END " & _
              "         ELSE '' END " & _
              "FROM Demographics D JOIN HaemResults H " & _
              "ON D.SampleID = H.SampleID " & _
              "WHERE D.RunDate = '" & Format$(dt, "dd/mmm/yyyy") & "' "
54720     If view(1) Then
54730         sql = sql & "AND ( " & _
                  "(cESR = 1  AND ((COALESCE(ESR, '') = '') OR ESR = '?')) " & _
                  "or (cretics = 1 and ((RetA is null) or retA = '?')) " & _
                  "or (cmonospot = 1 and ((MonoSpot is null) or monospot = '?'))" & _
                  "or (cMalaria = 1 and ((Malaria is null) or Malaria = '?')) " & _
                  "or (cRA = 1 and ((RA is null) or RA = '?')) " & _
                  "or (cSickledex = 1 and ((Sickledex is null) or Sickledex = '?'))" & _
                  ")"
54740     ElseIf view(2) Then
54750         sql = sql & "and ( cesr = 1 " & _
                  "or cretics = 1 " & _
                  "or cmonospot = 1 " & _
                  "or cMalaria = 1 " & _
                  "or cSickledex = 1" & _
                  "or cRA =1 ) "
        
54760     End If
54770     sql = sql & "order by D.SampleID"

54780     Set tb = New Recordset
54790     RecOpenClient 0, tb, sql
54800     Do While Not tb.EOF
54810         s = Trim$(tb!SampleID) & vbTab & _
                  Trim$(tb!Chart & "") & vbTab & _
                  tb!PatName & vbTab & _
                  tb!resESR & vbTab & _
                  tb!resRetA & vbTab & _
                  tb!resMonoSpot & vbTab & _
                  tb!resMalaria & vbTab & _
                  tb!resSickledex & vbTab & _
                  tb!resRA
        
54820         grd.AddItem s
54830         tb.MoveNext
54840     Loop

54850     If grd.Rows > 2 Then
54860         grd.RemoveItem 1
54870     End If

54880     Screen.MousePointer = vbNormal
54890     grd.Visible = True

54900     CheckForLotNumbers

54910     Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

54920     intEL = Erl
54930     strES = Err.Description
54940     LogError "fasot", "FillGrid", intEL, strES, sql
54950     Screen.MousePointer = vbNormal

End Sub

Private Sub bPrint_Click()

          Dim Y As Integer

54960     For Y = 0 To grd.Rows - 1
54970         Printer.Print grd.TextMatrix(Y, 0);
54980         Printer.Print Tab(10); grd.TextMatrix(Y, 1); 'chart
54990         Printer.Print Tab(18); Left$(grd.TextMatrix(Y, 2), 25); 'name
55000         Printer.Print Tab(44); grd.TextMatrix(Y, 3); 'esr
55010         Printer.Print Tab(50); grd.TextMatrix(Y, 4); 'retic
55020         Printer.Print Tab(68); grd.TextMatrix(Y, 5) 'im
55030     Next

55040     Printer.EndDoc

End Sub

Private Sub cmbCombo_Click()

55050     DataChanged = True

55060     grd.SetFocus

End Sub


Private Sub cmbCombo_KeyDown(KeyCode As Integer, Shift As Integer)

55070     Select Case KeyCode
              Case 27:
55080             cmbCombo.Visible = False
55090             grd.SetFocus
55100         Case 13:
55110             grd.SetFocus
55120         Case 38: 'Up
55130             grd.SetFocus
55140             DoEvents
55150             If grd.row > grd.FixedRows Then
55160                 grd.row = grd.row - 1
55170             End If
55180         Case 40: 'Down
55190             grd.SetFocus
55200             DoEvents
55210             If grd.row < grd.row - 1 Then
55220                 grd.row = grd.row + 1
55230             End If
55240     End Select

55250     DataChanged = True

End Sub


Private Sub cmbCombo_KeyPress(KeyAscii As Integer)

55260     If KeyAscii = 13 Then KeyAscii = 0

End Sub


Private Sub dt_CloseUp()

55270     FillGrid

End Sub


Private Sub Form_Activate()

55280     grd.Font.Bold = True

55290     If Not Activated Then
55300         Activated = True
55310         FillGrid
55320     End If

End Sub

Private Sub Form_Load()

55330     Activated = False

55340     dt = Format$(Now, "dd/mm/yyyy")

55350     ColumnEditType(0) = ctNoEdits
55360     ColumnEditType(1) = ctNoEdits
55370     ColumnEditType(2) = ctNoEdits
55380     ColumnEditType(3) = ctText
55390     ColumnEditType(4) = ctText
55400     ColumnEditType(5) = ctCombo
55410     ColumnEditType(6) = ctCombo
55420     ColumnEditType(7) = ctCombo
55430     ColumnEditType(8) = ctCombo

55440     cmbCombo.AddItem ""
55450     cmbCombo.AddItem "Negative"
55460     cmbCombo.AddItem "Positive"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

55470     If DataChanged Then
55480         If MsgBox("Save Changes?", vbQuestion + vbYesNo) = vbYes Then
55490             SaveDetails
55500         End If
55510     End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

55520     Activated = False

55530     DataChanged = False

End Sub

Private Sub grd_Click()

55540     If grd.MouseRow = 0 Then Exit Sub

55550     If ColumnEditType(grd.Col) = ctBoolean Then
        
55560         DataChanged = True
        
55570     ElseIf ColumnEditType(grd.Col) = ctCombo Then
        
55580         cmbCombo.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth
55590         cmbCombo.Visible = True
55600         cmbCombo.SetFocus
55610         DataChanged = True

55620     ElseIf ColumnEditType(grd.Col) = ctText Then
        
55630         GrdEdit 32
55640         DataChanged = True

55650     ElseIf ColumnEditType(grd.Col) <> ctNoEdits Then
55660         MsgBox "ColumnEditType not set for " & grd.Col
55670         Exit Sub
55680     End If

End Sub



Private Sub grd_GotFocus()

          Dim Analyte As String
          Dim SampleID As Long
          Dim f As Form

55690     If txtInput.Visible Then
        
55700         grd = txtInput
55710         txtInput.Visible = False

55720     ElseIf cmbCombo.Visible Then

55730         grd = cmbCombo
        
55740         cmbCombo.Visible = False
55750         cmbCombo = ""
        
55760         If grd = "" Then Exit Sub
        
55770         Select Case grd.Col
                  Case 5: Analyte = "Monospot"
55780             Case 6: Analyte = "Malaria"
55790             Case 7: Analyte = "Sickledex"
55800             Case Else: Exit Sub
55810         End Select
        
55820         If Trim$(lblLotNumber(grd.Col - 5)) = "" Then
55830             SampleID = grd.TextMatrix(grd.row, 0)
55840             Set f = New frmCheckReagentLotNumber
55850             With f
55860                 .Analyte = Analyte
55870                 .SampleID = SampleID
55880                 .Show 1
55890                 If Trim$(.LotNumber) <> "" Then
55900                     lblLotNumber(grd.Col - 5) = .LotNumber
55910                     lblExpiry(grd.Col - 5) = .Expiry
55920                 End If
55930             End With
55940             Unload f
55950             Set f = Nothing
55960         End If
          
55970     End If

End Sub


Private Sub grd_KeyPress(KeyAscii As Integer)

          Dim n As Integer

55980     If ColumnEditType(grd.Col) = ctText Then

55990         GrdEdit KeyAscii
56000         DataChanged = True

56010     ElseIf ColumnEditType(grd.Col) = ctCombo Then
        
56020         cmbCombo.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth
56030         cmbCombo = Chr$(KeyAscii)
56040         For n = 0 To cmbCombo.ListCount - 1
56050             If UCase$(Chr$(KeyAscii)) = UCase$(Left$(cmbCombo.List(n), 1)) Then
56060                 cmbCombo = cmbCombo.List(n)
56070                 Exit For
56080             End If
56090         Next
56100         cmbCombo.Visible = True
56110         cmbCombo.SetFocus
56120         cmbCombo.SelStart = 1000

56130         DataChanged = True

56140     End If

End Sub


Private Sub grd_LeaveCell()

56150     If txtInput.Visible = False Then Exit Sub

56160     grd = txtInput
56170     txtInput.Visible = False

End Sub


Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)

56180     Select Case KeyCode
              Case 27:
56190             txtInput.Visible = False
56200             grd.SetFocus
56210         Case 13:
56220             grd.SetFocus
56230         Case 38: 'Up
56240             grd.SetFocus
56250             DoEvents
56260             If grd.row > grd.FixedRows Then
56270                 grd.row = grd.row - 1
56280             End If
56290         Case 40: 'Down
56300             grd.SetFocus
56310             DoEvents
56320             If grd.row < grd.row - 1 Then
56330                 grd.row = grd.row + 1
56340             End If
56350     End Select

56360     DataChanged = True

End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)

56370     If KeyAscii = 13 Then KeyAscii = 0

End Sub


Private Sub view_Click(Index As Integer)

56380     FillGrid

End Sub

