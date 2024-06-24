VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAmend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Status Amendment"
   ClientHeight    =   6270
   ClientLeft      =   285
   ClientTop       =   675
   ClientWidth     =   12615
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
   Icon            =   "frmAmend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Index           =   1
      Left            =   195
      TabIndex        =   13
      Top             =   3120
      Width           =   12255
      Begin MSFlexGridLib.MSFlexGrid gCurrent 
         Height          =   885
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   1561
         _Version        =   393216
         Cols            =   20
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   32768
         ForeColorFixed  =   65535
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         FormatString    =   $"frmAmend.frx":08CA
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   0
      Left            =   195
      TabIndex        =   11
      Top             =   600
      Width           =   12255
      Begin MSFlexGridLib.MSFlexGrid gProduct 
         Height          =   2085
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "Not Editable - Lookup Only"
         Top             =   225
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   3678
         _Version        =   393216
         Cols            =   20
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
         FormatString    =   $"frmAmend.frx":0A93
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
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   450
      TabIndex        =   5
      Top             =   4680
      Width           =   5445
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Transfer to Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3420
         Picture         =   "frmAmend.frx":0C59
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   510
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   147128321
         CurrentDate     =   37883
      End
      Begin MSComCtl2.DTPicker dtTime 
         Height          =   315
         Left            =   1590
         TabIndex        =   8
         Top             =   510
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   147128323
         UpDown          =   -1  'True
         CurrentDate     =   37883
      End
      Begin VB.ComboBox cmbPossibleValues 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "cmbPossibleValues"
         Top             =   510
         Width           =   3195
      End
      Begin VB.Label lblPossibleValues 
         AutoSize        =   -1  'True
         Caption         =   "Possible Values"
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
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1110
      End
   End
   Begin VB.TextBox txtUnitNumber 
      Height          =   285
      Left            =   1185
      TabIndex        =   0
      Top             =   180
      Width           =   2445
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   9600
      Picture         =   "frmAmend.frx":109B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4905
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   765
      Left            =   11235
      Picture         =   "frmAmend.frx":1705
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4905
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   60
      TabIndex        =   18
      Top             =   6060
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblProduct 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6870
      TabIndex        =   17
      Top             =   180
      Width           =   5580
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4305
      TabIndex        =   16
      Top             =   180
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
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
      Left            =   3840
      TabIndex        =   15
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Product"
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
      TabIndex        =   4
      Top             =   210
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Number"
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
      TabIndex        =   2
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "frmAmend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum ipType
    ipNone
    ipCombo
    ipFreeText
    ipDate
    ipDatetime
End Enum


Private Sub EnableInput(ByVal Enable As ipType)

10  Select Case Enable
    Case ipNone
20      lblPossibleValues.Caption = "Possible Values"
30      lblPossibleValues.Enabled = False
40      cmbPossibleValues.Clear
50      cmdTransfer.Enabled = False
60      cmbPossibleValues.Visible = True
70      dtDate.Visible = False
80      dtTime.Visible = False

90  Case ipCombo
100     lblPossibleValues.Caption = "Possible Values"
110     lblPossibleValues.Enabled = True
120     cmdTransfer.Enabled = True
130     cmbPossibleValues.Visible = True
140     dtDate.Visible = False
150     dtTime.Visible = False

160 Case ipFreeText
170     lblPossibleValues.Caption = "Free Text"
180     lblPossibleValues.Enabled = True
190     cmbPossibleValues.Clear
200     cmdTransfer.Enabled = True
210     cmbPossibleValues.Visible = True
220     dtDate.Visible = False
230     dtTime.Visible = False

240 Case ipDate
250     lblPossibleValues.Caption = "Select Date"
260     lblPossibleValues.Enabled = True
270     cmbPossibleValues.Visible = False
280     dtDate.Visible = True
290     dtTime.Visible = False

300 Case ipDatetime
310     lblPossibleValues.Caption = "Select Date and Time"
320     lblPossibleValues.Enabled = True
330     cmbPossibleValues.Visible = False
340     dtDate.Visible = True
350     dtTime.Visible = True
360 End Select

End Sub

Private Sub FillCurrent()

    Dim s As String
    Dim Ps As New Products

10  On Error GoTo FillCurrent_Error

20  With gCurrent
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60  End With

70  Ps.LoadLatestISBT128 txtUnitNumber, ProductBarCodeFor(lblProduct)

80  If Ps.Count = 0 Then
90      iMsg "Unit Number not found", vbInformation
100     If TimedOut Then Unload Me: Exit Sub
110     Exit Sub
120 End If

130 With Ps.Item(1)
140     s = .ISBT128 & vbTab & _
            gEVENTCODES(.PackEvent).Text & vbTab & _
            .Chart & vbTab & _
            .PatName & vbTab & _
            .UserName & vbTab & _
            .RecordDateTime & vbTab & _
            Bar2Group(.GroupRh & "") & vbTab & _
            SupplierNameFor(.Supplier) & vbTab & _
            Format$(.DateExpiry, "dd/MMM/yyyy HH:mm") & vbTab & _
            .Screen & vbTab & _
            .SampleID & vbTab & _
            IIf(.crt, "Yes", "No") & vbTab & _
            IIf(.cco, "Yes", "No") & vbTab & _
            IIf(.cen, "Yes", "No") & vbTab & _
            IIf(.crtr, "Positive", "Negative") & vbTab & _
            IIf(.ccor, "Positive", "Negative") & vbTab & _
            IIf(.cenr, "Positive", "Negative") & vbTab & _
            .BarCode & vbTab & _
            IIf(.Checked, "Yes", "No") & vbTab & _
            .Notes
150     gCurrent.AddItem s
160 End With

170 If gCurrent.Rows > 2 Then
180     gCurrent.RemoveItem 1
190 End If

200 Exit Sub

FillCurrent_Error:

    Dim strES As String
    Dim intEL As Integer

210 intEL = Erl
220 strES = Err.Description
230 LogError "frmAmend", "FillCurrent", intEL, strES

End Sub

Private Sub FillPossibleWithBarCode()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo FillPossibleWithBarCode_Error

20  sql = "Select * from ProductList order by ListOrder"
30  Set tb = New Recordset
40  RecOpenServerBB 0, tb, sql

50  cmbPossibleValues.Clear

60  Do While Not tb.EOF
70      If Len(Trim$(tb!BarCode & "")) = 5 Then
80          cmbPossibleValues.AddItem tb!Wording
90      End If
100     tb.MoveNext
110 Loop

120 Exit Sub

FillPossibleWithBarCode_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmAmend", "FillPossibleWithBarCode", intEL, strES, sql

End Sub


Private Sub FillPossibleWithOperators()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo FillPossibleWithOperators_Error

20  With cmbPossibleValues

30      .Clear

40      sql = "Select * from Users order by ListOrder"
50      Set tb = New Recordset
60      RecOpenClient 0, tb, sql
70      Do While Not tb.EOF
80          .AddItem tb!Name & ""
90          tb.MoveNext
100     Loop

110 End With

120 Exit Sub

FillPossibleWithOperators_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmAmend", "FillPossibleWithOperators", intEL, strES, sql


End Sub



Private Sub FillPossibleWithYesNo()

10  With cmbPossibleValues
20      .Clear
30      .AddItem "Yes"
40      .AddItem "No"
50  End With

End Sub

Private Sub FillPossibleWithPosNeg()

10  With cmbPossibleValues
20      .Clear
30      .AddItem "Positive"
40      .AddItem "Negative"
50  End With

End Sub


Private Sub FillPossibleWithSupplier()

    Dim tb As Recordset
    Dim sql As String
    Dim s As String

10  On Error GoTo FillPossibleWithSupplier_Error

20  sql = "Select  * from Supplier order by ListOrder"
30  Set tb = New Recordset
40  RecOpenServerBB 0, tb, sql
50  cmbPossibleValues.Clear
60  Do While Not tb.EOF
70      s = tb!Supplier & ""
80      cmbPossibleValues.AddItem s
90      tb.MoveNext
100 Loop

110 Exit Sub

FillPossibleWithSupplier_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmAmend", "FillPossibleWithSupplier", intEL, strES, sql


End Sub


Private Sub FillPossibleWithGRH()

10  With cmbPossibleValues
20      .Clear
30      .AddItem "O Pos"
40      .AddItem "A Pos"
50      .AddItem "B Pos"
60      .AddItem "AB Pos"
70      .AddItem "O Neg"
80      .AddItem "A Neg"
90      .AddItem "B Neg"
100     .AddItem "AB Neg"
110 End With

End Sub


Private Sub cmdCancel_Click()

10  Unload Me

End Sub

Private Sub cmdSave_Click()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cmdSave_Click_Error

20  sql = "Select * from Latest where " & _
          "BarCode  = '" & ProductBarCodeFor(lblProduct) & "' " & _
          "and ISBT128 = '" & txtUnitNumber & "' " & _
          "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy hh:mm") & "'"
30  Set tb = New Recordset
40  RecOpenServerBB 0, tb, sql

50  If tb.EOF Then
60      iMsg "Unit Number not found", vbCritical
70      If TimedOut Then Unload Me: Exit Sub
80      Exit Sub
90  End If

100 tb!Event = gEVENTCODES.CodeFor(gCurrent.TextMatrix(1, 1))
110 tb!Patid = gCurrent.TextMatrix(1, 2)
120 tb!PatName = gCurrent.TextMatrix(1, 3)
130 tb!Operator = TechnicianCodeFor(gCurrent.TextMatrix(1, 4))
140 tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
150 tb!GroupRh = Group2Bar(gCurrent.TextMatrix(1, 6))
160 tb!Supplier = SupplierCodeFor(gCurrent.TextMatrix(1, 7))
170 tb!DateExpiry = Format(gCurrent.TextMatrix(1, 8), "dd/mmm/yyyy hh:mm")
180 tb!Screen = gCurrent.TextMatrix(1, 9)
190 tb!LabNumber = gCurrent.TextMatrix(1, 10)
200 tb!crt = gCurrent.TextMatrix(1, 11) = "Yes"
210 tb!cco = gCurrent.TextMatrix(1, 12) = "Yes"
220 tb!cen = gCurrent.TextMatrix(1, 13) = "Yes"
230 tb!crtr = gCurrent.TextMatrix(1, 14) = "Positive"
240 tb!ccor = gCurrent.TextMatrix(1, 15) = "Positive"
250 tb!cenr = gCurrent.TextMatrix(1, 16) = "Positive"
260 tb!Checked = gCurrent.TextMatrix(1, 18) = "Yes"
270 tb!Notes = gCurrent.TextMatrix(1, 19)

280 tb.Update

    '***************************************
    'BLR: if expiry date is changed then change all instances in product table
    '***************************************

290 If DateDiff("n", Format(lblExpiry, "dd/mmm/yyyy hh:mm"), Format(gCurrent.TextMatrix(1, 8), "dd/mmm/yyyy hh:mm")) <> 0 Then
300     sql = "Update Product " & _
              "Set DateExpiry = '" & Format(gCurrent.TextMatrix(1, 8), "dd/mmm/yyyy hh:mm") & "' where " & _
              "ISBT128 = '" & txtUnitNumber & "' and " & _
              "BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
              "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy hh:mm") & "'"
310     CnxnBB(0).Execute sql
320 End If
    'delete latest event entry from product
330 sql = "Delete From Product Where " & _
          "ISBT128 = '" & txtUnitNumber & "' and " & _
          "BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
          "and DateExpiry = '" & Format(gCurrent.TextMatrix(1, 8), "yyyy-mm-dd hh:mm") & "' " & _
          "and DateTime = '" & Format(gCurrent.TextMatrix(1, 5), "yyyy-mm-dd hh:mm:ss") & "'"
340 CnxnBB(0).Execute sql
    'now insert latest event record with new values
350 sql = "Insert into Product " & _
          "Select * from Latest where " & _
          "ISBT128 = '" & txtUnitNumber & "' and " & _
          "BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
          "and DateExpiry = '" & Format(gCurrent.TextMatrix(1, 8), "dd/mmm/yyyy hh:mm") & "'"
360 CnxnBB(0).Execute sql

370 cmdSave.Visible = False
380 cmdTransfer.Enabled = False
390 txtUnitNumber_LostFocus

400 Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

410 intEL = Erl
420 strES = Err.Description
430 LogError "frmAmend", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub cmbPossibleValues_KeyPress(KeyAscii As Integer)

10  If lblPossibleValues.Caption = "Possible Values" Then
20      KeyAscii = 0
30  End If

End Sub


Private Sub cmdTransfer_Click()

    Dim s As String

10  cmdSave.Visible = True

20  Select Case gCurrent.TextMatrix(0, gCurrent.col)

    Case "Event":
30      If cmbPossibleValues <> "" Then
40          gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
50      End If

60  Case "Patient ID":
70      gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues

80  Case "Patient Name":
90      gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues

100 Case "Operator":
110     If cmbPossibleValues <> "" Then
120         gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
130     End If

140 Case "Date/Time":
150     s = Format(dtDate, "dd/mm/yyyy") & " " & _
            Format(dtTime, "hh:nn:ss")
160     gCurrent.TextMatrix(1, gCurrent.col) = s

170 Case "Group/Rh":
180     If cmbPossibleValues <> "" Then
190         gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
200     End If

210 Case "Supplier":
220     If cmbPossibleValues <> "" Then
230         gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
240     End If

250 Case "Expiry Date":
260     s = Format(dtDate, "dd/mm/yyyy") & " " & Format(dtTime, "hh:mm")
270     gCurrent.TextMatrix(1, gCurrent.col) = s

280 Case "Screen":
290     gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues

300 Case "Lab Number":
310     gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues

320 Case "Room Temp", "Coombs", "Enzyme", _
         "Room Temp Result", "Coombs Result", "Enzyme Result":
330     If cmbPossibleValues <> "" Then
340         gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
350     End If

360 Case "Product":
370     If cmbPossibleValues <> "" Then
380         gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
390     End If

400 Case "Group Checked":
410     If cmbPossibleValues <> "" Then
420         gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
430     End If

440 Case "Notes":
450     gCurrent.TextMatrix(1, gCurrent.col) = cmbPossibleValues
460 End Select

End Sub

Private Sub dtDate_CloseUp()

10  cmdTransfer.Enabled = True

End Sub

Private Sub Form_Load()

10  EnableInput ipNone

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10  If TimedOut Then Exit Sub
20  If cmdSave.Visible Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60          Cancel = True
70      End If
80  End If

End Sub


Private Sub gCurrent_Click()

    Dim xSave As Integer
    Dim n As Integer

10  xSave = gCurrent.col

20  gCurrent.row = 1
30  For n = 0 To gCurrent.Cols - 1
40      gCurrent.col = n
50      gCurrent.CellBackColor = 0
60  Next

70  gCurrent.col = xSave
80  gCurrent.CellBackColor = vbYellow

90  Select Case gCurrent.TextMatrix(0, gCurrent.col)
    Case "Pack Number":
100     EnableInput ipNone

110 Case "Event":
        '    FillPossibleWithEvents
        '    EnableInput ipCombo
120     EnableInput ipNone
130 Case "Patient ID":
140     EnableInput ipFreeText

150 Case "Patient Name":
160     EnableInput ipFreeText

170 Case "Operator":
180     FillPossibleWithOperators
190     EnableInput ipCombo

200 Case "Date/Time":
210     EnableInput ipNone

220 Case "Group/Rh":
230     FillPossibleWithGRH
240     EnableInput ipCombo

250 Case "Supplier":
260     FillPossibleWithSupplier
270     EnableInput ipCombo

280 Case "Expiry Date":
290     EnableInput ipDatetime

300 Case "Screen":
310     EnableInput ipFreeText

320 Case "Lab Number":
330     EnableInput ipFreeText

340 Case "Room Temp":
350     FillPossibleWithYesNo
360     EnableInput ipCombo

370 Case "Coombs":
380     FillPossibleWithYesNo
390     EnableInput ipCombo

400 Case "Enzyme":
410     FillPossibleWithYesNo
420     EnableInput ipCombo

430 Case "Room Temp Result":
440     FillPossibleWithPosNeg
450     EnableInput ipCombo

460 Case "Coombs Result":
470     FillPossibleWithPosNeg
480     EnableInput ipCombo

490 Case "Enzyme Result":
500     FillPossibleWithPosNeg
510     EnableInput ipCombo

520 Case "Product":
530     FillPossibleWithBarCode
540     EnableInput ipNone

550 Case "Group Checked":
560     FillPossibleWithYesNo
570     EnableInput ipCombo

580 Case "Notes":
590     EnableInput ipFreeText
600 End Select

End Sub

Private Sub txtUnitNumber_LostFocus()

    Dim tb As Recordset
    Dim sql As String
    Dim s As String
    Dim Ps As New Products
    Dim f As Form
    Dim p As Product
    Dim strS As String

10  On Error GoTo txtUnitNumber_LostFocus_Error

20  If Trim$(txtUnitNumber) = "" Then Exit Sub

30  txtUnitNumber = UCase(txtUnitNumber)

40  If Left$(txtUnitNumber, 1) = "=" Then
50      strS = ISOmod37_2(Mid$(txtUnitNumber, 2, 13))
60      txtUnitNumber = Mid$(txtUnitNumber, 2, 13) & " " & strS
70  End If


80  With gProduct
90      .Rows = 2
100     .AddItem ""
110     .RemoveItem 1
120 End With

130 Ps.LoadLatestByUnitNumberISBT128 (txtUnitNumber)

140 If Ps.Count = 0 Then
150     iMsg "Unit Number not found."
160     If TimedOut Then Unload Me: Exit Sub
170     txtUnitNumber = ""
180     Exit Sub
190 ElseIf Ps.Count > 1 Then    'multiple products found
200     Set f = New frmSelectFromMultiple
210     f.ProductList = Ps
220     f.Show 1
230     Set p = f.SelectedProduct
240     Unload f
250     Set f = Nothing
260 Else
270     Set p = Ps.Item(1)
280 End If

290 lblExpiry = Format(p.DateExpiry, "dd/mm/yyyy HH:mm")
300 lblProduct = ProductWordingFor(p.BarCode)

310 sql = "Select * from Product where " & _
          "ISBT128 = '" & txtUnitNumber & "' " & _
          "and barcode = '" & ProductBarCodeFor(lblProduct) & "' " & _
          "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy HH:mm") & "' " & _
          "order by Counter"

320 Set tb = New Recordset
330 RecOpenServerBB 0, tb, sql

340 Do While Not tb.EOF
350     s = tb!ISBT128 & vbTab & _
            gEVENTCODES(tb!Event & "").Text & vbTab & _
            tb!Patid & vbTab & _
            tb!PatName & vbTab & _
            tb!Operator & vbTab & _
            tb!DateTime & vbTab & _
            Bar2Group(tb!GroupRh & "") & vbTab & _
            SupplierNameFor(tb!Supplier) & vbTab & _
            Format$(tb!DateExpiry, "dd/MMM/yyyy HH:mm") & vbTab & _
            tb!Screen & vbTab & _
            tb!LabNumber & vbTab & _
            IIf(tb!crt, "Yes", "No") & vbTab & _
            IIf(tb!cco, "Yes", "No") & vbTab & _
            IIf(tb!cen, "Yes", "No") & vbTab & _
            IIf(tb!crtr, "Positive", "Negative") & vbTab & _
            IIf(tb!ccor, "Positive", "Negative") & vbTab & _
            IIf(tb!cenr, "Positive", "Negative") & vbTab & _
            tb!BarCode & vbTab
360     If Not IsNull(tb!Checked) Then
370         s = s & IIf(tb!Checked, "Yes", "No")
380     End If
390     s = s & vbTab & tb!Notes & ""
400     gProduct.AddItem s
410     tb.MoveNext
420 Loop

430 If gProduct.Rows > 2 Then
440     gProduct.RemoveItem 1
450 End If

460 FillCurrent

470 Exit Sub

txtUnitNumber_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

480 intEL = Erl
490 strES = Err.Description
500 LogError "frmAmend", "txtUnitNumber_LostFocus", intEL, strES, sql


End Sub

