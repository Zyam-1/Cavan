VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBatchOccult 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetAcquire Batch Entry - Occult Blood"
   ClientHeight    =   6015
   ClientLeft      =   345
   ClientTop       =   495
   ClientWidth     =   9120
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
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   9120
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Enabled         =   0   'False
      Height          =   975
      Left            =   7890
      Picture         =   "frmBatchOccult.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   990
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdOccult 
      Height          =   4845
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   990
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   8546
      _Version        =   393216
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
      ScrollBars      =   2
      FormatString    =   "<Sample ID     |<Result       "
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
   Begin VB.CommandButton cmdSetAll 
      Caption         =   "S e t    1,  2,  a n d  3    A l l    N e g a t i v e"
      Height          =   405
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   7635
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All Negative"
      Height          =   400
      Index           =   2
      Left            =   5490
      TabIndex        =   3
      Top             =   540
      Width           =   2300
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All Negative"
      Height          =   400
      Index           =   1
      Left            =   2820
      TabIndex        =   2
      Top             =   540
      Width           =   2300
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All Negative"
      Height          =   400
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   2300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   7890
      Picture         =   "frmBatchOccult.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdOccult 
      Height          =   4845
      Index           =   1
      Left            =   2820
      TabIndex        =   6
      Top             =   990
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   8546
      _Version        =   393216
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
      ScrollBars      =   2
      FormatString    =   "<Sample ID     |<Result       "
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
   Begin MSFlexGridLib.MSFlexGrid grdOccult 
      Height          =   4845
      Index           =   2
      Left            =   5490
      TabIndex        =   7
      Top             =   990
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   8546
      _Version        =   393216
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
      ScrollBars      =   2
      FormatString    =   "<Sample ID     |<Result       "
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
Attribute VB_Name = "frmBatchOccult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim SID As Long

65390     On Error GoTo FillGrid_Error

65400     For n = 0 To 2
65410         With grdOccult(n)
65420             .Rows = 2
65430             .AddItem ""
65440             .RemoveItem 1
65450         End With
65460     Next

65470     sql = "select sampleid, request from faecesrequests50 where " & _
              "request = 'ob0' " & _
              "and sampleid not in(select sampleid from faecesresults50 where testname = 'ob0') " & _
              "Union " & _
              "select sampleid, request from faecesrequests50 where " & _
              "request = 'ob1' " & _
              "and sampleid not in(select sampleid from faecesresults50 where testname = 'ob1') " & _
              "Union " & _
              "select sampleid, request from faecesrequests50 where " & _
              "request = 'ob2' " & _
              "and sampleid not in(select sampleid from faecesresults50 where testname = 'ob2') " & _
              "ORDER BY SampleID"

65480     Set tb = New Recordset
65490     RecOpenServer 0, tb, sql

65500     Do While Not tb.EOF
        
65510         SID = tb!SampleID ' - sysOptMicroOffset(0)
65520         If tb!Request = "OB0" Then
65530             grdOccult(0).AddItem SID
10            ElseIf tb!Request = "OB1" Then
20                grdOccult(1).AddItem SID
30            ElseIf tb!Request = "OB2" Then
40                grdOccult(2).AddItem SID
50            End If
60            tb.MoveNext
70        Loop

80        For n = 0 To 2
90            With grdOccult(n)
100               If .Rows > 2 Then .RemoveItem 1
110           End With
120       Next

130       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBatchOccult", "FillGrid", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

170       If cmdSave.Enabled Then
180           If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
190               Exit Sub
200           End If
210       End If

220       Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim grd As Integer
          Dim Y As Integer
          Dim Result As String
          Dim Fx As FaecesResult
          Dim Fxs As New FaecesResults

230       On Error GoTo cmdSave_Click_Error

240       For grd = 0 To 2
250           With grdOccult(grd)
260               For Y = 1 To .Rows - 1
270                   If .TextMatrix(Y, 1) <> "" Then
280                       Set Fx = New FaecesResult
290                       Fx.SampleID = Val(.TextMatrix(Y, 0)) ' + sysOptMicroOffset(0)
300                       Fx.TestName = "OB" & Format$(grd)
310                       Fx.Result = IIf(.TextMatrix(Y, 1) = "Positive", "P", "N")
320                       Fx.UserName = UserName
330                       Fxs.Save Fx
340                   End If
350               Next
360           End With
370       Next

380       FillGrid
390       cmdSave.Enabled = False

400       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmBatchOccult", "cmdSave_Click", intEL, strES

End Sub

Private Sub cmdSet_Click(Index As Integer)

          Dim n As Integer

440       With grdOccult(Index)
450           .Col = 1
460           For n = 1 To .Rows - 1
470               .row = n
480               If .TextMatrix(n, 0) <> "" Then
490                   .Text = "Negative"
500                   .CellForeColor = vbGreen
510               End If
520           Next
530       End With

540       cmdSave.Enabled = True

End Sub

Private Sub cmdSetAll_Click()

          Dim n As Integer

550       For n = 0 To 2
560           cmdSet_Click (n)
570       Next

End Sub

Private Sub Form_Load()

580       FillGrid

End Sub

Private Sub grdOccult_Click(Index As Integer)

590       With grdOccult(Index)
600           If .TextMatrix(.row, 0) <> "" Then
610               .Col = 1
620               Select Case .Text
                      Case "":
630                       .Text = "Negative"
640                       .CellForeColor = vbGreen
650                   Case "Negative":
660                       .Text = "Positive"
670                       .CellForeColor = vbRed
680                   Case "Positive":
690                       .Text = ""
700                       .CellForeColor = vbBlack
710               End Select
720           End If
730       End With

740       cmdSave.Enabled = True
        
End Sub


