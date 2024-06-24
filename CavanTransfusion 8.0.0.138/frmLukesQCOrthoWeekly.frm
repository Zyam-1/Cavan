VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLukesQCOrthoWeekly 
   Caption         =   "NetAcquire --- Forward Grouping Cards (ADD) Quality Assurance"
   ClientHeight    =   7905
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   6495
   Icon            =   "frmLukesQCOrthoWeekly.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   6495
   Begin VB.TextBox txtAffirmagenLotNo 
      Height          =   285
      Left            =   1950
      MaxLength       =   50
      TabIndex        =   16
      Top             =   4425
      Width           =   2055
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   990
      TabIndex        =   11
      Top             =   6270
      Width           =   5235
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Previous"
      Enabled         =   0   'False
      Height          =   765
      Left            =   990
      Picture         =   "frmLukesQCOrthoWeekly.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6690
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   765
      Left            =   3060
      Picture         =   "frmLukesQCOrthoWeekly.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6690
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   5070
      Picture         =   "frmLukesQCOrthoWeekly.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6690
      Width           =   1155
   End
   Begin VB.TextBox txtCardLotNumber 
      Height          =   285
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   4
      Top             =   150
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grdLotNos 
      Height          =   2025
      Left            =   270
      TabIndex        =   2
      Top             =   900
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3572
      _Version        =   393216
      Rows            =   8
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   $"frmLukesQCOrthoWeekly.frx":1C08
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
   Begin MSFlexGridLib.MSFlexGrid grdSeraReactions 
      Height          =   1335
      Left            =   270
      TabIndex        =   1
      Top             =   4770
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   5
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
      FormatString    =   $"frmLukesQCOrthoWeekly.frx":1C98
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
   Begin MSFlexGridLib.MSFlexGrid grdCardReactions 
      Height          =   1335
      Left            =   270
      TabIndex        =   0
      Top             =   3030
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   5
      Cols            =   7
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   $"frmLukesQCOrthoWeekly.frx":1D19
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   780
      TabIndex        =   13
      Top             =   7590
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdAffirmagenCards 
      Height          =   1335
      Left            =   270
      TabIndex        =   14
      Top             =   4770
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   $"frmLukesQCOrthoWeekly.frx":1D95
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
   Begin VB.Label lblAffirmagen 
      AutoSize        =   -1  'True
      Caption         =   "Affirmagen Lot number"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   4470
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   270
      TabIndex        =   12
      Top             =   6300
      Width           =   660
   End
   Begin VB.Label lblCardExpiry 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1470
      TabIndex        =   10
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   540
      Width           =   420
   End
   Begin VB.Label lblLastEntered 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   150
      Width           =   2475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Card Lot Number"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   1200
   End
End
Attribute VB_Name = "frmLukesQCOrthoWeekly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdLoad_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim X As Integer
      Dim Y As Integer

10    On Error GoTo cmdLoad_Click_Error

20    sql = "Select top 1 * from StLukesGroupingCards " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      txtComment = tb!Comment & ""
70      With grdLotNos
80        .TextMatrix(1, 1) = tb!A1rrLot & ""
90        .TextMatrix(1, 2) = Format$(tb!A1rrExpiry, "dd/mm/yyyy")
100       .TextMatrix(2, 1) = tb!A2rrLot & ""
110       .TextMatrix(2, 2) = Format$(tb!A2rrExpiry, "dd/mm/yyyy")
120       .TextMatrix(3, 1) = tb!BrrLot & ""
130       .TextMatrix(3, 2) = Format$(tb!BrrExpiry, "dd/mm/yyyy")
140       .TextMatrix(4, 1) = tb!OR1wR1Lot & ""
150       .TextMatrix(4, 2) = Format$(tb!OR1wR1Expiry, "dd/mm/yyyy")
160       .TextMatrix(5, 1) = tb!AntiALot & ""
170       .TextMatrix(5, 2) = Format$(tb!AntiAExpiry, "dd/mm/yyyy")
180       .TextMatrix(6, 1) = tb!AntiBLot & ""
190       .TextMatrix(6, 2) = Format$(tb!AntiBExpiry, "dd/mm/yyyy")
200       .TextMatrix(7, 1) = tb!AntiDLot & ""
210       .TextMatrix(7, 2) = Format$(tb!AntiDExpiry, "dd/mm/yyyy")
220     End With
230     With grdCardReactions
240       For Y = 1 To 4
250         For X = 1 To 6
260           .TextMatrix(Y, X) = tb("C" & Format$(Y) & Format$(X)) & ""
270         Next
280       Next
290     End With
300     With grdSeraReactions
310       For Y = 1 To 4
320         For X = 1 To 3
330           .TextMatrix(Y, X) = tb("S" & Format$(Y) & Format$(X)) & ""
340         Next
350       Next
360     End With
370     txtCardLotNumber = tb!CardLotNumber
380     lblCardExpiry = Format$(tb!CardExpiry, "dd/mm/yyyy")
390   Else
400     With grdLotNos
410       For Y = 1 To 7
420         For X = 1 To 2
430           .TextMatrix(Y, X) = ""
440         Next
450       Next
460     End With
470     With grdCardReactions
480       For Y = 1 To 4
490         For X = 1 To 6
500           .TextMatrix(Y, X) = ""
510         Next
520       Next
530     End With
540     With grdSeraReactions
550       For Y = 1 To 4
560         For X = 1 To 3
570           .TextMatrix(Y, X) = ""
580         Next
590       Next
600     End With
610     txtCardLotNumber = ""
620     lblCardExpiry = ""
630   End If
 
640       With grdAffirmagenCards
650           .TextMatrix(1, 1) = tb!A1rrA1 & ""
660           .TextMatrix(1, 2) = tb!A1rrB & ""
670           .TextMatrix(2, 1) = tb!A2rrA1 & ""
680           .TextMatrix(2, 2) = tb!A2rrB & ""
690           .TextMatrix(3, 1) = tb!BA1 & ""
700           .TextMatrix(3, 2) = tb!BB & ""
710           .TextMatrix(4, 1) = tb!OA1 & ""
720           .TextMatrix(4, 2) = tb!OB & ""
730       End With
740       txtAffirmagenLotNo = tb!AffirmagenLotNo & ""
 
750   cmdSave.Enabled = True

760   Exit Sub

cmdLoad_Click_Error:

      Dim strES As String
      Dim intEL As Integer

770   intEL = Erl
780   strES = Err.Description
790   LogError "frmLukesQCOrthoWeekly", "cmdLoad_Click", intEL, strES, sql

 
End Sub

Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim X As Integer
      Dim Y As Integer

      'For Y = 1 To 7
      '  If Trim$(grdLotNos.TextMatrix(Y, 1)) = "" Then
      '    iMsg "All Lot Numbers must be filled.", vbCritical
      '    Exit Sub
      '  End If
      '  If Not IsDate(grdLotNos.TextMatrix(Y, 2)) Then
      '    iMsg "All Expiry Dates must be filled.", vbCritical
      '    Exit Sub
      '  End If
      'Next

10    On Error GoTo cmdSave_Click_Error

20    sql = "Select  * from StLukesGroupingCards where " & _
            "DateTime = '01/01/2000'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    tb.AddNew
60    tb!Comment = txtComment
70    With grdLotNos
80      tb!A1rrLot = .TextMatrix(1, 1)
90      If IsDate(.TextMatrix(1, 2)) Then
100       tb!A1rrExpiry = Format$(.TextMatrix(1, 2), "dd/mm/yyyy")
110     Else
120       tb!A1rrExpiry = Null
130     End If
  
140     tb!A2rrLot = .TextMatrix(2, 1)
150     If IsDate(.TextMatrix(2, 2)) Then
160       tb!A2rrExpiry = Format$(.TextMatrix(2, 2), "dd/mm/yyyy")
170     Else
180       tb!A2rrExpiry = Null
190     End If
  
200     tb!BrrLot = .TextMatrix(3, 1)
210     If IsDate(.TextMatrix(3, 2)) Then
220       tb!BrrExpiry = Format$(.TextMatrix(3, 2), "dd/mm/yyyy")
230     Else
240       tb!BrrExpiry = Null
250     End If
  
260     tb!OR1wR1Lot = .TextMatrix(4, 1)
270     If IsDate(.TextMatrix(4, 2)) Then
280       tb!OR1wR1Expiry = Format$(.TextMatrix(4, 2), "dd/mm/yyyy")
290     Else
300       tb!OR1wR1Expiry = Null
310     End If
  
320     tb!AntiALot = .TextMatrix(5, 1)
330     If IsDate(.TextMatrix(5, 2)) Then
340       tb!AntiAExpiry = Format$(.TextMatrix(5, 2), "dd/mm/yyyy")
350     Else
360       tb!AntiAExpiry = Null
370     End If
  
380     tb!AntiBLot = .TextMatrix(6, 1)
390     If IsDate(.TextMatrix(6, 2)) Then
400       tb!AntiBExpiry = Format$(.TextMatrix(6, 2), "dd/mm/yyyy")
410     Else
420       tb!AntiBExpiry = Null
430     End If
  
440     tb!AntiDLot = .TextMatrix(7, 1)
450     If IsDate(.TextMatrix(7, 2)) Then
460       tb!AntiDExpiry = Format$(.TextMatrix(7, 2), "dd/mm/yyyy")
470     Else
480       tb!AntiDExpiry = Null
490     End If
500   End With
510   With grdCardReactions
520     For Y = 1 To 4
530       For X = 1 To 6
540         tb("C" & Format$(Y) & Format$(X)) = Left$(.TextMatrix(Y, X), 1)
550       Next
560     Next
570   End With
580   With grdSeraReactions
590     For Y = 1 To 4
600       For X = 1 To 3
610         tb("S" & Format$(Y) & Format$(X)) = Left$(.TextMatrix(Y, X), 1)
620       Next
630     Next
640   End With
650       With grdAffirmagenCards
660           tb!A1rrA1 = .TextMatrix(1, 1)
670           tb!A1rrB = .TextMatrix(1, 2)
680           tb!A2rrA1 = .TextMatrix(2, 1)
690           tb!A2rrB = .TextMatrix(2, 2)
700           tb!BA1 = .TextMatrix(3, 1)
710           tb!BB = .TextMatrix(3, 2)
720           tb!OA1 = .TextMatrix(4, 1)
730           tb!OB = .TextMatrix(4, 2)
740           tb!AffirmagenLotNo = txtAffirmagenLotNo
750       End With
760   tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
770   tb!CardLotNumber = txtCardLotNumber
780   If IsDate(lblCardExpiry) Then
790     tb!CardExpiry = Format$(lblCardExpiry, "dd/mmm/yyyy")
800   Else
810     tb!CardExpiry = Null
820   End If
830   tb!Operator = UserName
840   tb.Update

850   cmdSave.Enabled = False
860   Unload Me

870   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

880   intEL = Erl
890   strES = Err.Description
900   LogError "frmLukesQCOrthoWeekly", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo Form_Load_Error

20    sql = "Select top 1 * from StLukesGroupingCards " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      s = "Last Entered by " & tb!Operator & vbCrLf & _
            "on " & Format$(tb!DateTime, "dd/mm/yyyy") & _
            " at " & Format$(tb!DateTime, "hh:mm:ss")
70      lblLastEntered = s
80    End If

90      grdSeraReactions.Visible = False
100     grdLotNos.Height = 1305
110     cmdLoad.Enabled = True
120     grdAffirmagenCards.Visible = True
130     lblAffirmagen.Visible = True
140     txtAffirmagenLotNo.Visible = True

150   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmLukesQCOrthoWeekly", "Form_Load", intEL, strES, sql


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Enabled Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub grdAffirmagenCards_Click()
10    On Error GoTo grdAffirmagenCards_Click_Error

20    With grdAffirmagenCards
30        If .MouseRow <> 0 And .MouseCol <> 0 Then
40                Select Case .TextMatrix(.Row, .Col)
                      Case "": .TextMatrix(.Row, .Col) = "0"
50                    Case "0": .TextMatrix(.Row, .Col) = "4"
60                    Case "4": .TextMatrix(.Row, .Col) = "3"
70                    Case "3": .TextMatrix(.Row, .Col) = "2"
80                    Case "2": .TextMatrix(.Row, .Col) = "1"
90                    Case "1": .TextMatrix(.Row, .Col) = "+"
100                   Case "+": .TextMatrix(.Row, .Col) = "NT"
110                   Case Else: .TextMatrix(.Row, .Col) = ""
120               End Select
130           cmdSave.Enabled = True
140       End If
150   End With

160   Exit Sub

grdAffirmagenCards_Click_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmLukesQCOrthoWeekly", "grdAffirmagenCards_Click", intEL, strES

End Sub

Private Sub grdCardReactions_Click()

10    With grdCardReactions
20      If .MouseRow <> 0 And .MouseCol <> 0 Then
30          Select Case .TextMatrix(.Row, .Col)
              Case "": .TextMatrix(.Row, .Col) = "0"
40            Case "0": .TextMatrix(.Row, .Col) = "4"
50            Case "4": .TextMatrix(.Row, .Col) = "3"
60            Case "3": .TextMatrix(.Row, .Col) = "2"
70            Case "2": .TextMatrix(.Row, .Col) = "1"
80            Case "1": .TextMatrix(.Row, .Col) = "+"
90            Case "+": .TextMatrix(.Row, .Col) = "NT"
100           Case Else: .TextMatrix(.Row, .Col) = ""
110         End Select
120       cmdSave.Enabled = True
130     End If
140   End With

End Sub

Private Sub grdLotNos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim strIP As String
      Dim s As String
      Dim f As Form
      Dim lot As String
      Dim Exp As String

10    If grdLotNos.MouseRow <> 0 And grdLotNos.MouseCol <> 0 Then
20      If grdLotNos.Col = 2 Then

30        Set f = New frmAskDate
40        With f
50          .Caption = "NetAcquire"
60          .DisplayDate = grdLotNos.TextMatrix(grdLotNos.Row, 2)
70          .Show 1
80          grdLotNos.TextMatrix(grdLotNos.Row, 2) = .DisplayDate
90        End With
100       Set f = Nothing
    
110     Else
    
120       s = "Enter Lot Number for " & grdLotNos.TextMatrix(grdLotNos.Row, 0)
130       strIP = iBOX(s, , grdLotNos.TextMatrix(grdLotNos.Row, 1))
140       If TimedOut Then Unload Me: Exit Sub
150       grdLotNos.TextMatrix(grdLotNos.Row, 1) = strIP
160       If grdLotNos.Row = 4 And grdLotNos.Col = 1 And Len(strIP) = 10 Then
170         strIP = "8SS" & Right$(strIP, 3)
180         grdLotNos.TextMatrix(4, 1) = strIP
190       ElseIf Len(strIP) = 14 Then
200         Exp = Right$(strIP, 2) & "/" & Mid$(strIP, 11, 2) & "/" & Mid$(strIP, 9, 2)
210         If IsDate(Exp) Then
220           Exp = Format$(Exp, "dd/MMM/yyyy")
230           grdLotNos.TextMatrix(grdLotNos.Row, 2) = Exp
  
240           lot = Left$(strIP, 5) & "." & Mid$(strIP, 6, 2) & "." & Mid$(strIP, 8, 1)
250           grdLotNos.TextMatrix(grdLotNos.Row, 1) = lot
260         Else
270           grdLotNos.TextMatrix(grdLotNos.Row, 1) = strIP
280         End If
290       Else
300         grdLotNos.TextMatrix(grdLotNos.Row, 1) = strIP
310       End If
320     End If
  
330     cmdSave.Enabled = True

340   End If

End Sub


Private Sub grdSeraReactions_Click()

10    With grdSeraReactions
20      If .MouseRow <> 0 And .MouseCol <> 0 Then
30        Select Case .TextMatrix(.Row, .Col)
            Case "": .TextMatrix(.Row, .Col) = "0"
40          Case "0": .TextMatrix(.Row, .Col) = "+"
      '      Case "1": .TextMatrix(.Row, .Col) = "2"
      '      Case "2": .TextMatrix(.Row, .Col) = "3"
      '      Case "3": .TextMatrix(.Row, .Col) = "4"
      '      Case "4": .TextMatrix(.Row, .Col) = "+"
50          Case "+": .TextMatrix(.Row, .Col) = "Not Tested"
60          Case Else: .TextMatrix(.Row, .Col) = ""
70        End Select
80        cmdSave.Enabled = True
90      End If
100   End With

End Sub


Private Sub lblCardExpiry_Click()

      Dim f As Form

10    Set f = New frmAskDate
20    With f
30      .Caption = "Expiry Date"
40      .DisplayDate = lblCardExpiry
50      .Show 1
60      lblCardExpiry = .DisplayDate
70    End With
80    Set f = Nothing

End Sub


Private Sub txtCardLotNumber_LostFocus()

      Dim Expiry As String
      Dim lot As String

10    If Len(txtCardLotNumber) = 20 Then

20      If Mid$(txtCardLotNumber, 7, 2) <> "48" Then
30        iMsg "This is not an ABO DD Blood Grouping Card", vbCritical
40        If TimedOut Then Unload Me: Exit Sub
50        Exit Sub
60      End If
  
70      Expiry = Left$(txtCardLotNumber, 2) & "/" & Mid$(txtCardLotNumber, 3, 2) & "/" & Mid$(txtCardLotNumber, 5, 2)
80      Expiry = Format$(Expiry, "dd/MMM/yyyy")
90      lblCardExpiry = Expiry
  
100     lot = Mid$(txtCardLotNumber, 15, 3)
110     txtCardLotNumber = "ADD" & lot & "A"
120   End If

End Sub


