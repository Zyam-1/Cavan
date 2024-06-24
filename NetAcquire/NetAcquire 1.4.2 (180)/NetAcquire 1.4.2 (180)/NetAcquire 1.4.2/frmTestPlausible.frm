VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestPlausible 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Plausible Ranges"
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
      Picture         =   "frmTestPlausible.frx":0000
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
      Picture         =   "frmTestPlausible.frx":1982
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
      Picture         =   "frmTestPlausible.frx":75A0
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
   Begin MSFlexGridLib.MSFlexGrid grdPlausible 
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
      FormatString    =   "<Long Name          |<Short Name     |<Plausibe Low |<Plausible High "
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
Attribute VB_Name = "frmTestPlausible"
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

17890 On Error GoTo FillGrid_Error

17900 With grdPlausible
17910   .Rows = 2
17920   .AddItem ""
17930   .RemoveItem 1
17940 End With

17950 sql = "SELECT DISTINCT LongName, ShortName, " & _
            "COALESCE(PlausibleLow, 0) Low, " & _
            "COALESCE(PlausibleHigh, 9999) High, PrintPriority " & _
            "FROM " & mDept & "TestDefinitions WHERE " & _
            "SampleType = '" & pSampleType1 & "' " & _
            "AND InUse = 1 " & _
            "ORDER BY PrintPriority"
17960 Set tb = New Recordset
17970 RecOpenServer 0, tb, sql
17980 Do While Not tb.EOF
17990   s = tb!LongName & vbTab & _
            tb!ShortName & vbTab & _
            tb!Low & vbTab & _
            tb!High
18000   grdPlausible.AddItem s
18010   tb.MoveNext
18020 Loop

18030 With grdPlausible
18040   If .Rows > 2 Then
18050     .RemoveItem 1
18060     .AddItem ""
18070   End If
18080 End With

18090 Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

18100 intEL = Erl
18110 strES = Err.Description
18120 LogError "frmTestPlausible", "FillGrid", intEL, strES, sql


End Sub

Private Sub FillSampleTypes()

      Dim sql As String
      Dim tb As Recordset

18130 On Error GoTo FillSampleTypes_Error

18140 sql = "SELECT Text FROM Lists " & _
            "WHERE ListType = 'ST' " & _
            "ORDER BY ListOrder"
18150 Set tb = New Recordset
18160 RecOpenServer 0, tb, sql

18170 cmbSampleType.Clear
18180 Do While Not tb.EOF
18190   cmbSampleType.AddItem tb!Text & ""
18200   tb.MoveNext
18210 Loop

18220 If pSampleType1 <> "" Then
18230   cmbSampleType = ListTextFor("ST", pSampleType1)
18240 Else
18250   pSampleType1 = "S"
18260   cmbSampleType = "Serum"
18270 End If

18280 Exit Sub

FillSampleTypes_Error:

      Dim strES As String
      Dim intEL As Integer

18290 intEL = Erl
18300 strES = Err.Description
18310 LogError "frmTestPlausible", "FillSampleTypes", intEL, strES, sql


End Sub
Private Sub cmbSampleType_Click()

18320 pSampleType1 = ListCodeFor("ST", cmbSampleType)

18330 FillGrid

End Sub


Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

18340 KeyAscii = 0

End Sub


Private Sub cmdExit_Click()

18350 If cmdSave.Visible Then
18360   If iMsg("Cancel without Saving?", vbYesNo + vbQuestion) = vbYes Then
18370     Unload Me
18380   End If
18390 Else
18400   Unload Me
18410 End If

End Sub

Private Sub cmdExport_Click()

18420 ExportFlexGrid grdPlausible, Me

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim Y As Integer

18430 On Error GoTo cmdSave_Click_Error

18440 For Y = 1 To grdPlausible.Rows - 1
        
18450   sql = "UPDATE " & mDept & "TestDefinitions " & _
              "SET PlausibleLow = " & Val(grdPlausible.TextMatrix(Y, 2)) & ", " & _
              "PlausibleHigh = " & Val(grdPlausible.TextMatrix(Y, 3)) & " " & _
              "WHERE ShortName = '" & grdPlausible.TextMatrix(Y, 1) & "'"
18460   Cnxn(0).Execute sql

18470 Next

18480 cmdSave.Visible = False

18490 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

18500 intEL = Erl
18510 strES = Err.Description
18520 LogError "frmTestPlausible", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

18530 FillSampleTypes

18540 FillGrid

End Sub


Private Sub grdPlausible_KeyUp(KeyCode As Integer, Shift As Integer)

18550 If grdPlausible.Col < 2 Then Exit Sub
18560 If grdPlausible.TextMatrix(grdPlausible.row, 0) = "" Then Exit Sub

18570 If EditGrid(grdPlausible, KeyCode, Shift) Then
18580   cmdSave.Visible = True
18590 End If

End Sub

Private Function EditGrid(ByVal g As MSFlexGrid, _
                         ByVal KeyCode As Integer, _
                         ByVal Shift As Integer) _
                         As Boolean

      'returns true if grid changed

      Dim ShiftDown As Boolean
      Dim RetVal As Boolean

18600 RetVal = False

18610 If g.row < g.FixedRows Then
18620   Exit Function
18630 ElseIf g.Col < g.FixedCols Then
18640   Exit Function
18650 End If
18660 ShiftDown = (Shift And vbShiftMask) > 0

18670 Select Case KeyCode
        Case vbKeyA To vbKeyZ:
18680     If ShiftDown Then
18690       g = g & Chr(KeyCode)
18700       RetVal = True
18710     Else
18720       g = g & Chr(KeyCode + 32)
18730       RetVal = True
18740     End If
        
18750   Case vbKey0 To vbKey9:
18760     g = g & Chr(KeyCode)
18770     RetVal = True
        '+++ Junaid 15-02-2024
18780   Case (189)
18790       If g = "" Then
18800           g = "-"
18810       End If
        '--- Junaid
18820   Case vbKeyBack:
18830     If Len(g) > 0 Then
18840       g = Left$(g, Len(g) - 1)
18850       RetVal = True
18860     End If
        
18870   Case &HBE, vbKeyDecimal:
18880     g = g & "."
18890     RetVal = True
          
18900   Case vbKeySpace:
18910     g = g & " "
18920     RetVal = True
          
18930   Case vbKeyNumpad0 To vbKeyNumpad9:
18940   Case vbKeyDelete:
18950   Case vbKeyLeft:
18960   Case vbKeyRight:
18970   Case vbKeyUp:
18980   Case vbKeyDown:
18990   Case vbKeyTab:
19000 End Select

19010 EditGrid = RetVal

End Function


Public Property Let Discipline(ByVal strNewValue As String)

19020 mDept = UCase$(strNewValue)

End Property

Public Property Let SampleType(ByVal strNewValue As String)

19030 pSampleType1 = UCase$(strNewValue)

End Property


