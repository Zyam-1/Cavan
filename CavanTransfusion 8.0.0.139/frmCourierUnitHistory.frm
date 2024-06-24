VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCourierUnitHistory 
   Caption         =   "NetAcquire - Courier Unit History"
   ClientHeight    =   7995
   ClientLeft      =   960
   ClientTop       =   600
   ClientWidth     =   15720
   ControlBox      =   0   'False
   Icon            =   "frmCourierUnitHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   15720
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   900
      Left            =   14700
      Picture         =   "frmCourierUnitHistory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   900
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   900
      Left            =   3840
      Picture         =   "frmCourierUnitHistory.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   900
   End
   Begin VB.TextBox txtUnitNumber 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   450
      Width           =   2325
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   6510
      Left            =   45
      TabIndex        =   0
      Top             =   1170
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   11483
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmCourierUnitHistory.frx":13BF
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   5
      Top             =   7725
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Number"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   525
      Width           =   885
   End
End
Attribute VB_Name = "frmCourierUnitHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdSearch_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo cmdSearch_Click_Error

20    grd.Rows = 2
30    grd.AddItem ""
40    grd.RemoveItem 1

50    sql = "Select * from Courier where " & _
            "UnitNumber = '" & Replace(txtUnitNumber, " ", "") & "' " & _
            "Order by MessageTime desc"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF

90      Select Case tb!Identifier & ""
          Case "RS", "RS3"
100         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Identifier & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                tb!UnitExpiry & vbTab & _
                ReadableGroup(tb!UnitGroup & "") & vbTab & _
                tb!StockComment & vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & vbTab & _
                ReadableGroup(tb!PatientGroup & "") & vbTab & _
                tb!DeReservationDateTime
  
110       Case "SM", "RTS"
120         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Identifier & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                tb!ActionText & vbTab & _
                tb!UserName & ""
    
130       Case "FT", "FT1"
140         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Identifier & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & vbTab & _
                vbTab & _
                vbTab
150         Select Case tb!SampleStatus & ""
              Case "T": s = s & "Transfused"
160           Case "A": s = s & "Aborted"
170           Case "S": s = s & "Spiked"
180           Case "U": s = s & "Unknown"
190           Case "D": s = s & "Destroyed"
200           Case Else: s = s & "???"
210         End Select
220         s = s & vbTab & tb!UserName & ""
    
    
230       Case "SU3"
240         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Identifier & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                tb!UnitExpiry & vbTab & _
                ReadableGroup(tb!UnitGroup & "") & vbTab & _
                tb!StockComment & ""
  
  
250       Case "ST"
260         s = Format$(tb!MessageTime, "dd/mm/yy hh:mm:ss") & vbTab & _
                tb!Identifier & vbTab & _
                ProductWordingFor(tb!ProductCode & "") & vbTab & _
                tb!Location & vbTab & _
                vbTab & _
                vbTab & _
                vbTab & _
                tb!Chart & vbTab & _
                tb!PatName & vbTab & _
                tb!DoB & vbTab & _
                FullSex(tb!Sex & "") & _
                tb!UserName & ""

270     End Select
  
280     grd.AddItem s
290     tb.MoveNext

300   Loop

310   If grd.TextMatrix(1, 0) = "" And grd.Rows > 2 Then
320       grd.RemoveItem 1
330   End If

340   Exit Sub

cmdSearch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "frmCourierUnitHistory", "cmdSearch_Click", intEL, strES, sql


End Sub

Private Sub Form_Resize()

10    If Me.Width < 5955 Then
20      Me.Width = 5955
30    End If
40    If Me.Height < 4920 Then
50      Me.Height = 4920
60    End If

70    grd.Width = Me.Width - 370
80    grd.Height = Me.Height - 1605
End Sub


Private Sub grd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
      Dim s As String

10    If grd.MouseCol <> 1 Or grd.MouseCol = 0 Then
20      s = ""
30    Else
40      Select Case grd.TextMatrix(grd.MouseRow, 1)
          Case "RS": s = "Reserve Stock"
50        Case "SM": s = "Stock Movement"
60        Case "RTS": s = "Return to Stock"
70        Case "FT": s = "Fate of Unit"
80        Case "SU3": s = "Stock Update"
90        Case "ST": s = "Stock Transfer"
100       Case Else: s = ""
110     End Select
120   End If

130   grd.ToolTipText = s

End Sub



Private Sub txtUnitNumber_LostFocus()
          Dim s As String
          
10        txtUnitNumber = UCase(txtUnitNumber)
          
20        If Left$(txtUnitNumber, 1) = "=" Then    'Barcode scanning entry
30            s = ISOmod37_2(Mid$(txtUnitNumber, 2, 13))
40            txtUnitNumber = Mid$(txtUnitNumber, 2, 13) & " " & s
50        End If

End Sub
