VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestDelta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Delta Checking"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1100
      Left            =   8310
      Picture         =   "frmTestDelta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5430
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   8310
      Picture         =   "frmTestDelta.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3690
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   8310
      Picture         =   "frmTestDelta.frx":75A0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1200
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   7890
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "cmbSampleType"
      Top             =   720
      Width           =   1965
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   240
      TabIndex        =   5
      Top             =   60
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdDelta 
      Height          =   7695
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13573
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
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
      FormatString    =   "<Long Name          |<Short Name     |^Enable Delta Check |<Delta Value(Absolute)     |<  Days   "
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
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestDelta.frx":846A
      Top             =   300
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestDelta.frx":8740
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SampleType"
      Height          =   195
      Left            =   8370
      TabIndex        =   3
      Top             =   510
      Width           =   885
   End
End
Attribute VB_Name = "frmTestDelta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDept As String
Private pSampleType1 As String

Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

8470      On Error GoTo FillGrid_Error

8480      With grdDelta

8490          .Visible = False
8500          .Rows = 2
8510          .AddItem ""
8520          .RemoveItem 1
8530          .Col = 2

8540          sql = "SELECT DISTINCT LongName, ShortName, " & _
                    "COALESCE(DoDelta, 0) DoDelta, " & _
                    "COALESCE(DeltaLimit, 9999) DeltaLimit, " & _
                    "PrintPriority,DeltaDaysBackLimit " & _
                    "FROM " & mDept & "TestDefinitions WHERE " & _
                    "SampleType = '" & pSampleType1 & "' " & _
                    "AND InUse = 1 " & _
                    "GROUP BY LongName, ShortName, DoDelta, DeltaLimit, PrintPriority,DeltaDaysBackLimit ORDER BY PrintPriority"
8550          Set tb = New Recordset
8560          RecOpenServer 0, tb, sql
8570          Do While Not tb.EOF
8580              s = tb!LongName & vbTab & _
                      tb!ShortName & vbTab & _
                      vbTab & _
                      tb!DeltaLimit & _
                      vbTab & _
                      tb!DeltaDaysBackLimit
8590              .AddItem s
8600              .row = .Rows - 1
8610              Set .CellPicture = IIf(tb!DoDelta, imgSquareTick.Picture, imgSquareCross.Picture)
8620              .CellPictureAlignment = flexAlignCenterCenter

8630              tb.MoveNext

8640          Loop

8650          If .Rows > 2 Then
8660              .RemoveItem 1
8670          End If

8680          .Visible = True

8690      End With

8700      Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

8710      intEL = Erl
8720      strES = Err.Description
8730      LogError "frmTestDelta", "FillGrid", intEL, strES, sql


End Sub

Private Sub FillSampleTypes()

          Dim sql As String
          Dim tb As Recordset

8740      On Error GoTo FillSampleTypes_Error

8750      sql = "SELECT Text FROM Lists " & _
                "WHERE ListType = 'ST' " & _
                "ORDER BY ListOrder"
8760      Set tb = New Recordset
8770      RecOpenServer 0, tb, sql

8780      cmbSampleType.Clear
8790      Do While Not tb.EOF
8800          cmbSampleType.AddItem tb!Text & ""
8810          tb.MoveNext
8820      Loop

8830      If pSampleType1 <> "" Then
8840          cmbSampleType = ListTextFor("ST", pSampleType1)
8850      Else
8860          pSampleType1 = "S"
8870          cmbSampleType = "Serum"
8880      End If

8890      Exit Sub

FillSampleTypes_Error:

          Dim strES As String
          Dim intEL As Integer

8900      intEL = Erl
8910      strES = Err.Description
8920      LogError "frmTestDelta", "FillSampleTypes", intEL, strES, sql


End Sub
Private Sub cmbSampleType_Click()

8930      pSampleType1 = ListCodeFor("ST", cmbSampleType)

8940      With grdDelta
8950          .Rows = 2
8960          .AddItem ""
8970          .RemoveItem 1
8980      End With

8990      FillGrid

End Sub


Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

9000      KeyAscii = 0

End Sub


Private Sub cmdExit_Click()

9010      If cmdSave.Visible Then
9020          If iMsg("Cancel without Saving?", vbYesNo + vbQuestion) = vbYes Then
9030              Unload Me
9040          End If
9050      Else
9060          Unload Me
9070      End If

End Sub

Private Sub cmdExport_Click()

9080      ExportFlexGrid grdDelta, Me

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Integer

9090      On Error GoTo cmdSave_Click_Error

9100      grdDelta.Col = 2
9110      For Y = 1 To grdDelta.Rows - 1

9120          grdDelta.row = Y
9130          sql = "UPDATE " & mDept & "TestDefinitions " & _
                    "SET DoDelta = " & IIf(grdDelta.CellPicture = imgSquareCross.Picture, 0, 1) & ", " & _
                    "DeltaLimit = " & Val(grdDelta.TextMatrix(Y, 3)) & " , " & _
                    "DeltaDaysBackLimit = " & Val(grdDelta.TextMatrix(Y, 4)) & " " & _
                    "WHERE ShortName = '" & grdDelta.TextMatrix(Y, 1) & "'"
9140          Cnxn(0).Execute sql

9150      Next

9160      cmdSave.Visible = False

9170      Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

9180      intEL = Erl
9190      strES = Err.Description
9200      LogError "frmTestDelta", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

9210      FillSampleTypes

9220      FillGrid

End Sub


Private Sub grdDelta_KeyUp(KeyCode As Integer, Shift As Integer)

      'If grdDelta.Col <> 3 Or grdDelta.Col <> 4 Then Exit Sub
9230      If grdDelta.Col < 3 Then
9240          Exit Sub
9250      End If
9260      If grdDelta.TextMatrix(grdDelta.row, 0) = "" Then Exit Sub

9270      If EditGrid(grdDelta, KeyCode, Shift) Then
9280          cmdSave.Visible = True
9290      End If

End Sub

Private Function EditGrid(ByVal g As MSFlexGrid, _
                          ByVal KeyCode As Integer, _
                          ByVal Shift As Integer) _
                          As Boolean

      'returns true if grid changed

          Dim ShiftDown As Boolean
          Dim RetVal As Boolean

9300      RetVal = False

9310      If g.row < g.FixedRows Then
9320          Exit Function
9330      ElseIf g.Col < g.FixedCols Then
9340          Exit Function
9350      End If
9360      ShiftDown = (Shift And vbShiftMask) > 0

9370      Select Case KeyCode
          Case vbKeyA To vbKeyZ:
9380          If ShiftDown Then
9390              g = g & Chr(KeyCode)
9400              RetVal = True
9410          Else
9420              g = g & Chr(KeyCode + 32)
9430              RetVal = True
9440          End If

9450      Case vbKey0 To vbKey9:
9460          g = g & Chr(KeyCode)
9470          RetVal = True

9480      Case vbKeyBack:
9490          If Len(g) > 0 Then
9500              g = Left$(g, Len(g) - 1)
9510              RetVal = True
9520          End If

9530      Case &HBE, vbKeyDecimal:
9540          g = g & "."
9550          RetVal = True

9560      Case vbKeySpace:
9570          g = g & " "
9580          RetVal = True

9590      Case vbKeyNumpad0 To vbKeyNumpad9:
9600      Case vbKeyDelete:
9610      Case vbKeyLeft:
9620      Case vbKeyRight:
9630      Case vbKeyUp:
9640      Case vbKeyDown:
9650      Case vbKeyTab:
9660      End Select

9670      EditGrid = RetVal

End Function


Public Property Let Discipline(ByVal strNewValue As String)

9680      mDept = UCase$(strNewValue)

End Property

Public Property Let SampleType(ByVal strNewValue As String)

9690      pSampleType1 = UCase$(strNewValue)

End Property


Private Sub grdDelta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9700      With grdDelta

9710          If .MouseRow > 0 And .MouseCol = 2 Then

9720              .row = .MouseRow
9730              .Col = 2

9740              If .CellPicture = imgSquareCross.Picture Then
9750                  Set .CellPicture = imgSquareTick.Picture
9760              Else
9770                  Set .CellPicture = imgSquareCross.Picture
9780              End If
9790              .CellPictureAlignment = flexAlignCenterCenter

9800              cmdSave.Visible = True

9810          End If

9820      End With

End Sub


