VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBGARangeHistory 
   Caption         =   "NetAcquire - Blood Gas - Normal Ranges Change History"
   ClientHeight    =   4155
   ClientLeft      =   135
   ClientTop       =   1050
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   11445
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   705
      Left            =   10140
      Picture         =   "frmBGARangeHistory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2505
      Left            =   120
      TabIndex        =   0
      Top             =   750
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   16
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmBGARangeHistory.frx":066A
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tot CO2"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   7
      Left            =   9720
      TabIndex        =   8
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O2Sat"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   6
      Left            =   8505
      TabIndex        =   7
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BE"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   5
      Left            =   7290
      TabIndex        =   6
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HCO3"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   4
      Left            =   6075
      TabIndex        =   5
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "pO2"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   3
      Left            =   4860
      TabIndex        =   4
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "pCO2"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   2
      Left            =   3645
      TabIndex        =   3
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "pH"
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   2430
      TabIndex        =   2
      Top             =   450
      Width           =   1200
   End
End
Attribute VB_Name = "frmBGARangeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

750       Unload Me

End Sub


Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

760       On Error GoTo Form_Load_Error

770       g.Rows = 2
780       g.AddItem ""
790       g.RemoveItem 1

800       sql = "Select * from BGADefinitions " & _
              "Order by DateTimeAmended Desc"
810       Set tb = New Recordset
820       RecOpenServer 0, tb, sql

830       Do While Not tb.EOF
840           s = Format$(tb!DateTimeAmended, "dd/mm/yy hh:mm") & vbTab & _
                  tb!AmendedBy & vbTab & _
                  tb!pHLow & vbTab & tb!phhigh & vbTab & _
                  tb!pCO2Low & vbTab & tb!pCO2High & vbTab & _
                  tb!PO2Low & vbTab & tb!po2high & vbTab & _
                  tb!HCO3Low & vbTab & tb!hco3high & vbTab & _
                  tb!BELow & vbTab & tb!BEHigh & vbTab & _
                  tb!O2SATLow & vbTab & tb!O2SatHigh & vbTab & _
                  tb!totCO2Low & vbTab & tb!TotCO2High & ""
850           g.AddItem s
860           tb.MoveNext
870       Loop

880       If g.Rows > 2 Then
890           g.RemoveItem 1
900       End If

910       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

920       intEL = Erl
930       strES = Err.Description
940       LogError "frmBGARangeHistory", "Form_Load", intEL, strES, sql


End Sub



