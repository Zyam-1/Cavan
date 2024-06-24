VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatientWorkList 
   Caption         =   "NetAcquire - Work List"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   630
   ClientWidth     =   11625
   Icon            =   "7frmPatientWorkList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   11625
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4905
      Left            =   330
      TabIndex        =   2
      Top             =   900
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   8652
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmPatientWorkList.frx":08CA
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
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   10155
      TabIndex        =   1
      Top             =   225
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   315
      Left            =   330
      TabIndex        =   0
      Top             =   285
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   37187
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   375
      TabIndex        =   3
      Top             =   6030
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmPatientWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

      Dim sn As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    sql = "select * from patientdetails where " & _
            "daterequired = '" & Format(DTPicker, "dd/mmm/yyyy") & "'"
30    Set sn = New Recordset
40    RecOpenServerBB 0, sn, sql

50    g.Rows = 2
60    g.AddItem ""
70    g.RemoveItem 1

80    Do While Not sn.EOF
90      s = sn!Patnum & vbTab & _
            sn!Name & vbTab & _
            sn!Clinician & vbTab & _
            sn!Ward & ""
100     g.AddItem s
110     sn.MoveNext
120   Loop

130   If g.Rows = 2 Then
140     s = vbTab & "No Entries"
150     g.AddItem s
160   End If

170   g.RemoveItem 1

180   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmPatientWorkList", "FillG", intEL, strES, sql


End Sub

Private Sub btnCancel_Click()

10    Unload Me

End Sub


Private Sub DTPicker_CloseUp()

10    FillG

End Sub




Private Sub Form_Load()

10    DTPicker = Format(Now, "dd/mmm/yyyy")

20    Activated = False

30    g.Font.Bold = True

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
40        FillG
      '**************************************
End Sub


