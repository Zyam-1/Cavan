VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAutoValHaem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Autovalidation"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Height          =   885
      Left            =   5250
      Picture         =   "frmAutoValHaem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2670
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   885
      Left            =   5250
      Picture         =   "frmAutoValHaem.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4170
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   885
      Left            =   5250
      Picture         =   "frmAutoValHaem.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5580
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6165
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   "<Parameter         |<Low       |<High     |^Include "
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
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   5220
      TabIndex        =   5
      Top             =   3570
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAutoValHaem.frx":2C5E
      Height          =   1065
      Left            =   4620
      TabIndex        =   1
      Top             =   300
      Width           =   2235
   End
End
Attribute VB_Name = "frmAutoValHaem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim sql As String
          Dim tb As Recordset

64100     On Error GoTo FillG_Error

64110     g.Rows = 2
64120     g.AddItem ""
64130     g.RemoveItem 1

64140     sql = "SELECT Parameter + CHAR(9) + " & _
              "CAST(Low AS nvarchar(50)) + CHAR(9) + " & _
              "CAST(High AS nvarchar(50)) + CHAR(9) + " & _
              "CASE Include WHEN 1 THEN 'Yes' ELSE 'No' END Rec " & _
              "FROM HaemAutoVal " & _
              "ORDER BY ListOrder"
64150     Set tb = New Recordset
64160     RecOpenServer 0, tb, sql
64170     Do While Not tb.EOF
64180         g.AddItem tb!Rec
64190         tb.MoveNext
64200     Loop

64210     If g.Rows > 2 Then
64220         g.RemoveItem 1
64230     End If

64240     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

64250     intEL = Erl
64260     strES = Err.Description
64270     LogError "frmAutoValHaem", "FillG", intEL, strES, sql


End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim Parameter As String
          Dim l As Single
          Dim H As Single
          Dim Inc As Integer

64280     On Error GoTo cmdSave_Click_Error

64290     For n = 1 To g.Rows - 1
64300         Parameter = g.TextMatrix(n, 0)
64310         l = Val(g.TextMatrix(n, 1))
64320         H = Val(g.TextMatrix(n, 2))
64330         Inc = IIf(g.TextMatrix(n, 3) = "Yes", 1, 0)
        
64340         sql = "SELECT * FROM HaemAutoVal " & _
                  "WHERE Parameter = '" & Parameter & "'"
64350         Set tb = New Recordset
64360         RecOpenServer 0, tb, sql
        
64370         If Not tb.EOF Then
64380             tb!Low = l
64390             tb!High = H
64400             tb!Include = Inc
64410             tb.Update
64420         End If
64430     Next

64440     cmdSave.Enabled = False

64450     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

64460     intEL = Erl
64470     strES = Err.Description
64480     LogError "frmAutoValHaem", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

64490     Unload Me

End Sub

Private Sub cmdExport_Click()

64500     ExportFlexGrid g, Me, "Auto Validation Ranges" & vbCr

End Sub

Private Sub Form_Load()

64510     FillG

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

64520     If cmdSave.Enabled Then
64530         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
64540             Cancel = True
64550         End If
64560     End If

End Sub


Private Sub g_Click()

64570     If g.MouseRow = 0 Then Exit Sub
64580     If g.MouseCol = 0 Then Exit Sub

64590     With g
64600         Select Case .Col
                  Case 1:
64610                 .TextMatrix(.row, 1) = Format$(Val(iBOX("Auto Validation Low", , .TextMatrix(.row, 1))))
64620             Case 2:
64630                 .TextMatrix(.row, 2) = Format$(Val(iBOX("Auto Validation High", , .TextMatrix(.row, 2))))
64640             Case 3:
64650                 .TextMatrix(.row, 3) = IIf(.TextMatrix(.row, 3) = "Yes", "No", "Yes")
64660         End Select
64670     End With

64680     cmdSave.Enabled = True
        
End Sub


