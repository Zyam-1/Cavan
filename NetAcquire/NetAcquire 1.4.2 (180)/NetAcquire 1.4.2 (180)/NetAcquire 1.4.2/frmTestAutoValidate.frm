VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestAutoValidate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Auto-validation Ranges"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1100
      Left            =   5970
      Picture         =   "frmTestAutoValidate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5430
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   5970
      Picture         =   "frmTestAutoValidate.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3690
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   5970
      Picture         =   "frmTestAutoValidate.frx":75A0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1200
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   5550
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
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdAutoVal 
      Height          =   7695
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   13573
      _Version        =   393216
      Cols            =   4
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
      FormatString    =   "<Long Name          |<Short Name     |<Auto-Val Low |<Auto-Val High "
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SampleType"
      Height          =   195
      Left            =   6030
      TabIndex        =   3
      Top             =   510
      Width           =   885
   End
End
Attribute VB_Name = "frmTestAutoValidate"
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

5270  On Error GoTo FillGrid_Error

5280  With grdAutoVal
5290    .Visible = False
5300    .Rows = 2
5310    .AddItem ""
5320    .RemoveItem 1

5330    sql = "SELECT DISTINCT LongName, ShortName, " & _
              "COALESCE(AutoValLow, 0) Low, " & _
              "COALESCE(AutoValHigh, 9999) High, " & _
              "PrintPriority " & _
              "FROM " & mDept & "TestDefinitions WHERE " & _
              "SampleType = '" & pSampleType1 & "' " & _
              "AND InUse = 1 " & _
              "ORDER BY PrintPriority"
5340    Set tb = New Recordset
5350    RecOpenServer 0, tb, sql
5360    Do While Not tb.EOF
5370      s = tb!LongName & vbTab & _
              tb!ShortName & vbTab & _
              tb!Low & vbTab & _
              tb!High
5380      .AddItem s
5390      tb.MoveNext
5400    Loop
        
5410    If .Rows > 2 Then
5420      .RemoveItem 1
5430      .AddItem ""
5440    End If
5450    .Visible = True
        
5460  End With

5470  Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

5480  intEL = Erl
5490  strES = Err.Description
5500  LogError "frmTestAutoValidate", "FillGrid", intEL, strES, sql


End Sub

Private Sub FillSampleTypes()

      Dim sql As String
      Dim tb As Recordset

5510  On Error GoTo FillSampleTypes_Error

5520  sql = "SELECT Text FROM Lists " & _
            "WHERE ListType = 'ST' " & _
            "ORDER BY ListOrder"
5530  Set tb = New Recordset
5540  RecOpenServer 0, tb, sql

5550  cmbSampleType.Clear
5560  Do While Not tb.EOF
5570    cmbSampleType.AddItem tb!Text & ""
5580    tb.MoveNext
5590  Loop

5600  If pSampleType1 <> "" Then
5610    cmbSampleType = ListTextFor("ST", pSampleType1)
5620  Else
5630    pSampleType1 = "S"
5640    cmbSampleType = "Serum"
5650  End If

5660  Exit Sub

FillSampleTypes_Error:

      Dim strES As String
      Dim intEL As Integer

5670  intEL = Erl
5680  strES = Err.Description
5690  LogError "frmTestAutoValidate", "FillSampleTypes", intEL, strES, sql


End Sub
Private Sub cmbSampleType_Click()

5700  pSampleType1 = ListCodeFor("ST", cmbSampleType)

5710  FillGrid

End Sub


Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

5720  KeyAscii = 0

End Sub


Private Sub cmdExit_Click()

5730  If cmdSave.Visible Then
5740    If iMsg("Cancel without Saving?", vbYesNo + vbQuestion) = vbYes Then
5750      Unload Me
5760    End If
5770  Else
5780    Unload Me
5790  End If

End Sub

Private Sub cmdExport_Click()

5800  ExportFlexGrid grdAutoVal, Me

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim Y As Integer

5810  On Error GoTo cmdSave_Click_Error

5820  For Y = 1 To grdAutoVal.Rows - 1
        
5830    sql = "UPDATE " & mDept & "TestDefinitions " & _
              "SET AutoValLow = " & Val(grdAutoVal.TextMatrix(Y, 2)) & ", " & _
              "AutoValHigh = " & Val(grdAutoVal.TextMatrix(Y, 3)) & " " & _
              "WHERE ShortName = '" & grdAutoVal.TextMatrix(Y, 1) & "'"
5840    Cnxn(0).Execute sql

5850  Next

5860  cmdSave.Visible = False

5870  Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

5880  intEL = Erl
5890  strES = Err.Description
5900  LogError "frmTestAutoValidate", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Activate()

5910  Select Case mDept
        Case "TM":
5920    Case "CD4"
5930      cmbSampleType.Enabled = False
5940    Case "BIO"
5950    Case "HAEM"
5960      cmbSampleType.Enabled = False
5970  End Select

End Sub

Private Sub Form_Load()

5980  FillSampleTypes

5990  FillGrid

End Sub


Private Sub grdAutoVal_KeyUp(KeyCode As Integer, Shift As Integer)

6000  If grdAutoVal.Col < 2 Then Exit Sub
6010  If grdAutoVal.TextMatrix(grdAutoVal.row, 0) = "" Then Exit Sub

6020  If EditGrid(grdAutoVal, KeyCode, Shift) Then
6030    cmdSave.Visible = True
6040  End If

End Sub

Private Function EditGrid(ByVal g As MSFlexGrid, _
                         ByVal KeyCode As Integer, _
                         ByVal Shift As Integer) _
                         As Boolean

      'returns true if grid changed

      Dim ShiftDown As Boolean
      Dim RetVal As Boolean

6050  RetVal = False

6060  If g.row < g.FixedRows Then
6070    Exit Function
6080  ElseIf g.Col < g.FixedCols Then
6090    Exit Function
6100  End If
6110  ShiftDown = (Shift And vbShiftMask) > 0

6120  Select Case KeyCode
        Case vbKeyA To vbKeyZ:
6130      If ShiftDown Then
6140        g = g & Chr(KeyCode)
6150        RetVal = True
6160      Else
6170        g = g & Chr(KeyCode + 32)
6180        RetVal = True
6190      End If
        
6200    Case vbKey0 To vbKey9:
6210      g = g & Chr(KeyCode)
6220      RetVal = True
        
6230    Case vbKeyBack:
6240      If Len(g) > 0 Then
6250        g = Left$(g, Len(g) - 1)
6260        RetVal = True
6270      End If
        
6280    Case &HBE, vbKeyDecimal:
6290      g = g & "."
6300      RetVal = True
          
6310    Case vbKeySpace:
6320      g = g & " "
6330      RetVal = True
          
6340    Case vbKeyNumpad0 To vbKeyNumpad9:
6350    Case vbKeyDelete:
6360    Case vbKeyLeft:
6370    Case vbKeyRight:
6380    Case vbKeyUp:
6390    Case vbKeyDown:
6400    Case vbKeyTab:
6410  End Select

6420  EditGrid = RetVal

End Function


Public Property Let Discipline(ByVal strNewValue As String)

6430  mDept = UCase$(strNewValue)

End Property

Public Property Let SampleType(ByVal strNewValue As String)

6440  pSampleType1 = UCase$(strNewValue)

End Property


