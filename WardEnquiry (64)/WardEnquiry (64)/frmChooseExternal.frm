VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmChooseExternal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - External Tests"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   1065
      Left            =   10800
      Picture         =   "frmChooseExternal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3750
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4185
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   0
      SelectionMode   =   1
      FormatString    =   $"frmChooseExternal.frx":0ECA
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Sample of interest"
      ForeColor       =   &H80000018&
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   2220
   End
End
Attribute VB_Name = "frmChooseExternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SampleID As String
Private m_SampleDate As String

Private m_DateSID As Collection

Private Sub cmdCancel_Click()

10    m_SampleID = ""

20    Me.Hide

End Sub

Private Sub Form_Activate()

Dim S() As String
Dim Item As Variant
Dim sql As String
Dim tb As Recordset
Dim Rq As String
Dim ResultStatus As String

10    On Error GoTo Form_Activate_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    For Each Item In m_DateSID
60    Rq = ""
70    S = Split(Item)
80    sql = "SELECT Analyte FROM ExtResults WHERE SampleID = '" & S(0) & "'"
90    Set tb = New Recordset
100   RecOpenServer 0, tb, sql
110   Do While Not tb.EOF
120     Rq = Rq & tb!Analyte & " "
130     tb.MoveNext
140   Loop
150   sql = "SELECT Count(*) Tot FROM BioResults WHERE SampleID = '" & S(0) & "' AND  ((code='SYS001') OR(analyser ='BIOMNIS' OR Analyser = 'MATER' OR analyser ='NVRL' OR analyser = 'BEAUMONT'))"
160   Set tb = New Recordset
170   RecOpenServer 0, tb, sql
180   If tb!Tot > 0 Then
190     ResultStatus = "Results Available"
200   Else
210     ResultStatus = "Tests Requested"
220   End If
230   g.AddItem S(0) & vbTab & S(1) & vbTab & Trim$(Rq) & vbTab & ResultStatus
240   Next
      
250   If g.Rows > 2 Then
260   g.RemoveItem 1
270   End If
      g.col = 1
      g.Sort = 9
       
       
       
280   Exit Sub

Form_Activate_Error:

Dim strES As String
Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmChooseExternal", "Form_Activate", intEL, strES, sql


End Sub


Public Property Get SampleID() As String

10    SampleID = m_SampleID

End Property


Public Property Get SampleDate() As String

10    SampleDate = m_SampleDate

End Property


Public Property Set DateSID(colDateSID As Collection)

10    Set m_DateSID = colDateSID

End Property


'
'Private Sub Form_Load()
'    'Zyam 1-3-24
'      SortFlexGridByDate g, 1
'    'Zyam 1-3-24
'End Sub

Private Sub Form_Unload(Cancel As Integer)

10    Set m_DateSID = Nothing

End Sub

Private Sub g_Click()

10    If g.MouseRow = 0 Then
20        g.Row = 1
30    End If
      
40    If g.TextMatrix(g.Row, 3) <> "Results Available" Then
50        m_SampleID = ""
60    Else
70        m_SampleID = g.TextMatrix(g.Row, 0)
80    End If
90    Me.Hide

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    With g
        Cmp = IIf(CDate(.TextMatrix(Row1, 1)) > CDate(.TextMatrix(Row2, 1)), 1, -1)
    End With
End Sub

'Zyam 1-3-24
'Private Sub SortFlexGridByDate(flexGrid As MSFlexGrid, Column As Integer)
'    Dim i As Integer
'    Dim J As Integer
'    Dim Temp As Integer
'    Dim numRows As Integer
'    Dim dates() As Date
'    'MsgBox (g.TextMatrix(2, 1))
'    numRows = flexGrid.Rows - 1
'    ReDim dates(0 To numRows)
'
'    ' Store dates from the specified column
'    For i = 1 To numRows
'
'        dates(i) = ParseDate(flexGrid.TextMatrix(i, Column))
'    Next i
'
'    ' Perform a simple bubble sort based on dates
'    For i = 1 To numRows - 1
'        For J = i + 1 To numRows
'            If dates(J) < dates(i) Then
'                ' Swap row indices
'                Temp = flexGrid.Row
'                flexGrid.Row = J
'                flexGrid.TextMatrix(0, 0) = Temp ' Hide flickering
'                Temp = flexGrid.Row
'                flexGrid.Row = i
'                flexGrid.TextMatrix(0, 0) = Temp ' Hide flickering
'            End If
'        Next J
'    Next i
'End Sub
'
'Private Function ParseDate(dateString As String) As Date
'    Dim parts() As String
'    parts = Split(dateString, "/")
'    ParseDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
'End Function
''Zyam 1-3-24
