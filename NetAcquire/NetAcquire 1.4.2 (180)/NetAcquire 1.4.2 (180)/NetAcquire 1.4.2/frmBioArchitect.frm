VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioArchitect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Architect Codes"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      HelpContextID   =   10026
      Left            =   5310
      Picture         =   "frmBioArchitect.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1485
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   615
      Left            =   5310
      Picture         =   "frmBioArchitect.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4530
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid gCodes 
      Height          =   6405
      HelpContextID   =   10090
      Left            =   270
      TabIndex        =   0
      Top             =   330
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   11298
      _Version        =   393216
      Cols            =   3
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Long Name                       |<Short Name        |<Architect Code   "
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5430
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmBioArchitect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

5320      On Error GoTo FillG_Error

5330      With gCodes
5340          .Rows = 2
5350          .AddItem ""
5360          .RemoveItem 1
5370      End With

5380      sql = "SELECT LongName, ShortName, ArchitectCode FROM BioTestDefinitions " & _
              "GROUP BY LongName, ShortName, ArchitectCode " & _
              "ORDER BY LongName"
5390      Set tb = New Recordset
5400      RecOpenServer 0, tb, sql
5410      Do While Not tb.EOF
5420          s = tb!LongName & vbTab & _
                  tb!ShortName & vbTab & _
                  tb!ArchitectCode & ""
5430          gCodes.AddItem s
5440          tb.MoveNext
5450      Loop

5460      If gCodes.Rows > 2 Then
5470          gCodes.RemoveItem 1
5480      End If

5490      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

5500      intEL = Erl
5510      strES = Err.Description
5520      LogError "frmBioArchitect", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

5530      Unload Me

End Sub

Private Sub cmdXL_Click()

5540      ExportFlexGrid gCodes, Me

End Sub


Private Sub Form_Load()

5550      EnsureColumnExists "BioTestDefinitions", "ArchitectCode", "nvarchar(50)"

5560      FillG

End Sub

Private Sub gCodes_Click()

          Dim s As String
          Dim strIP As String
          Dim sql As String

5570      On Error GoTo gCodes_Click_Error

5580      If gCodes.MouseRow = 0 Then Exit Sub

5590      gCodes.Enabled = False

5600      s = "Enter Architect Code" & vbCrLf & _
              "for " & gCodes.TextMatrix(gCodes.row, 0)
5610      strIP = gCodes.TextMatrix(gCodes.row, 2)
5620      strIP = iBOX(s, , strIP)

5630      sql = "UPDATE BioTestDefinitions " & _
              "SET ArchitectCode = '" & strIP & "' " & _
              "WHERE LongName = '" & gCodes.TextMatrix(gCodes.row, 0) & "'"
5640      Cnxn(0).Execute sql

5650      gCodes.Enabled = True

5660      FillG

5670      Exit Sub

gCodes_Click_Error:

          Dim strES As String
          Dim intEL As Integer

5680      intEL = Erl
5690      strES = Err.Description
5700      LogError "frmBioArchitect", "gCodes_Click", intEL, strES, sql


End Sub



