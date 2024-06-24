VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPatHistoryChart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1100
      Left            =   4860
      Picture         =   "frmPatHistoryChart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5850
      Width           =   1200
   End
   Begin VB.CommandButton bOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   1100
      Left            =   3450
      Picture         =   "frmPatHistoryChart.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5850
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5415
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   240
      Picture         =   "frmPatHistoryChart.frx":0AAC
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmPatHistoryChart.frx":0D82
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmPatHistoryChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sPatientHistory As String
Private m_sLabNoUpd As String

Private m_objEditScreen As Form

Private Sub Form_Unload(Cancel As Integer)

30630     Set m_objEditScreen = Nothing

End Sub





'---------------------------------------------------------------------------------------
' Procedure : bOK_Click
' Author    : XPMUser
' Date      : 19/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bOK_Click()
30640 On Error GoTo bOK_Click_Error

      Dim i As Integer
      Dim PatientSelected As String
      Dim s As String
      Dim sql As String
      Dim MultiSelectedDemoForLabNoUpdate As String

30650 PatientSelected = "'-9'"

30660 With m_objEditScreen
30670     .gMDemoLabNoUpd.Cols = 6
30680     .gMDemoLabNoUpd.Clear
30690     .gMDemoLabNoUpd.FormatString = "<Chart     |<Name                |<Address        |<Dob        |<Sex  |<    "
30700     .txtMultiSeltdDemoForLabNoUpd = ""
30710     .gMDemoLabNoUpd.Rows = 1
30720 End With

30730 With g
30740     For i = 0 To .Rows - 1
30750         .row = i
30760         .Col = 5
30770         If .CellPicture = imgSquareTick Then
                  ' PatientSelected = PatientSelected & " OR " & " ( PatName = '" & .TextMatrix(i, 1) & "' AND Addr0= '" & .TextMatrix(i, 2) & "' AND DoB= '" & Format(.TextMatrix(i, 3), "dd/mm/yyyy") & "' AND Sex= '" & .TextMatrix(i, 4) & "' )"
30780             PatientSelected = PatientSelected & " OR " & " ( upper(PatName) = '" & AddTicks((UCase(.TextMatrix(i, 1)))) & "' AND DoB= '" & Format(.TextMatrix(i, 3), "dd/MMM/yyyy") & "' AND upper(Sex)= '" & UCase(.TextMatrix(i, 4)) & "'  AND upper(Chart)= '" & UCase(.TextMatrix(i, 0)) & "' )"
30790             s = .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 2) & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 4)
30800             m_objEditScreen.gMDemoLabNoUpd.AddItem s
30810         End If
30820     Next i
30830 End With
30840 MultiSelectedDemoForLabNoUpdate = Replace(Replace(PatientSelected, "'-9' OR ", ""), "'-9'", "")

30850 If MultiSelectedDemoForLabNoUpdate <> "" Then
30860     sql = "UPDATE demographics "
30870     sql = sql & " SET LabNo ='" & m_sLabNoUpd & "'"
30880     sql = sql & " WHERE " & MultiSelectedDemoForLabNoUpdate
30890     m_objEditScreen.txtMultiSeltdDemoForLabNoUpd = sql
30900 Else
30910     m_objEditScreen.txtMultiSeltdDemoForLabNoUpd = ""

30920 End If

30930 Unload Me


30940 Exit Sub


bOK_Click_Error:

      Dim strES As String
      Dim intEL As Integer

30950 intEL = Erl
30960 strES = Err.Description
30970 LogError "frmPatHistoryChart", "bOK_Click", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
30980 On Error GoTo cmdCancel_Click_Error


30990 Unload Me


31000 Exit Sub


cmdCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

31010 intEL = Erl
31020 strES = Err.Description
31030 LogError "frmPatHistoryChart", "cmdCancel_Click", intEL, strES
End Sub

Private Sub Form_Load()
31040 Me.Caption = "Matching Demographics"
31050 FillGrid (m_sPatientHistory)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FillGrid
' Author    : XPMUser
' Date      : 19/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FillGrid(PateintHistroy As String)

31060 On Error GoTo FillGrid_Error
      Dim sql As String
      Dim s As String
      Dim tb As New ADODB.Recordset
      Dim i As Integer

31070 sql = "select D.PatName,D.Chart,D.DoB,D.Sex, D.LabNo from Demographics D " & _
            " WHERE D.SAMPLEID <> '-9' " & PateintHistroy & " " & _
            "  GROUP BY D.PatName,D.Chart,D.DoB,D.Sex, D.LabNo " & _
            "HAVING COALESCE(LabNo, '') = ''"

31080 With g
31090     .Cols = 7
31100     .ColAlignment(5) = 2
          
31110     .FormatString = "<Chart          |<Name                                       |<Address                            |<Dob               |<Sex  |<   |<LabNo  "
31120     .ColWidth(2) = 0
31130     Set tb = New Recordset
31140     RecOpenClient 0, tb, sql
31150     Do While Not tb.EOF
31160         s = tb!Chart & "" & vbTab & tb!PatName & "" & vbTab & "" & vbTab & tb!DoB & vbTab & tb!Sex & "" & vbTab & vbTab & tb!LabNo & ""
31170         g.AddItem s
31180         .row = g.Rows - 1
31190         .Col = 5

31200         Set .CellPicture = imgSquareCross.Picture
31210         tb.MoveNext
31220     Loop
31230     If g.row > 2 Then
31240         .RemoveItem (1)
31250     End If
31260 End With

31270 With m_objEditScreen.gMDemoLabNoUpd
31280     If .Rows > 1 Then
31290         For i = 1 To m_objEditScreen.gMDemoLabNoUpd.Rows - 1
                  '.TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 2) & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 4)
31300             Call FndValueInGridtoTick(.TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4))
31310         Next i
31320     End If
31330 End With

31340 Exit Sub


FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

31350 intEL = Erl
31360 strES = Err.Description
31370 LogError "frmPatHistoryChart", "FillGrid", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FndValueInGridtoTick
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FndValueInGridtoTick(Chart As String, Name As String, Address As String, DoB As String, Sex As String)

31380 On Error GoTo FndValueInGridtoTick_Error
      Dim R As Integer

31390 With g
31400     For R = 1 To .Rows - 1
            '.TextMatrix(R, 2) = Address And
31410         If .TextMatrix(R, 0) = Chart And .TextMatrix(R, 1) = Name And .TextMatrix(R, 3) = DoB And .TextMatrix(R, 4) = Sex Then
31420             .row = R
31430             .Col = 5
31440             Set .CellPicture = imgSquareTick.Picture
31450         End If
31460     Next R
31470 End With



31480 Exit Sub


FndValueInGridtoTick_Error:

      Dim strES As String
      Dim intEL As Integer

31490 intEL = Erl
31500 strES = Err.Description
31510 LogError "frmPatHistoryChart", "FndValueInGridtoTick", intEL, strES
End Sub


Public Property Get PatientHistory() As String

31520 PatientHistory = m_sPatientHistory

End Property

Public Property Let PatientHistory(ByVal sPatientHistory As String)

31530 m_sPatientHistory = sPatientHistory

End Property

'---------------------------------------------------------------------------------------
' Procedure : g_MouseUp
' Author    : XPMUser
' Date      : 19/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

31540 On Error GoTo g_MouseUp_Error

31550 With g

31560     If .MouseRow > 0 And .MouseCol = 5 Then
31570         .row = .MouseRow
31580         .Col = 5
31590         If .CellPicture = imgSquareCross.Picture Then
31600             Set .CellPicture = imgSquareTick.Picture
31610         Else
31620             Set .CellPicture = imgSquareCross.Picture
31630         End If
31640         .CellPictureAlignment = flexAlignCenterCenter
31650     End If

31660 End With




31670 Exit Sub


g_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

31680 intEL = Erl
31690 strES = Err.Description
31700 LogError "frmPatHistoryChart", "g_MouseUp", intEL, strES
End Sub

Public Property Get LabNoUpd() As String

31710 LabNoUpd = m_sLabNoUpd

End Property

Public Property Let LabNoUpd(ByVal sLabNoUpd As String)

31720 m_sLabNoUpd = sLabNoUpd

End Property


Public Property Get EditScreen() As Form

31730     Set EditScreen = m_objEditScreen

End Property

Public Property Set EditScreen(objEditScreen As Form)

31740     Set m_objEditScreen = objEditScreen

End Property
