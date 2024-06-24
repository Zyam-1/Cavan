VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmAntibodyList 
   Caption         =   "NetAcquire - Positive Antibody List"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   Icon            =   "frmAntibodyList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bprint 
      Caption         =   "Print"
      Height          =   825
      Left            =   10530
      Picture         =   "frmAntibodyList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   10530
      Picture         =   "frmAntibodyList.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "NEQAS"
      Height          =   915
      Left            =   10485
      TabIndex        =   2
      Top             =   60
      Width           =   1275
      Begin VB.OptionButton optExclude 
         Caption         =   "Exclude"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   570
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optInclude 
         Caption         =   "Include"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   10530
      Picture         =   "frmAntibodyList.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6570
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6885
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   12144
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmAntibodyList.frx":18A8
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   7020
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   10545
      TabIndex        =   6
      Top             =   1950
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmAntibodyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "SELECT AIDR, AandE, PatNum, Name, DoB, LabNumber, DateTime FROM PatientDetails WHERE " & _
            "AIDR LIKE '%pos%' "
60    If optExclude Then
70      sql = sql & "AND Name NOT LIKE '%NEQAS%' "
80    End If
90    sql = sql & "ORDER BY DateTime DESC"
100   Set tb = New Recordset
110   RecOpenServerBB 0, tb, sql
120   Do While Not tb.EOF
130     s = tb!AIDR & vbTab & _
            tb!Patnum & vbTab & _
            tb!AandE & vbTab & _
            tb!Name & vbTab & _
            tb!DoB & vbTab & _
            tb!LabNumber & vbTab & _
            tb!DateTime
140     g.AddItem s
150     tb.MoveNext
160   Loop

170   If g.Rows > 2 Then
180     g.RemoveItem 1
190   End If

200   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmAntibodyList", "FillG", intEL, strES, sql


End Sub

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName
20        If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.Font.Size = 9
50    Printer.Orientation = vbPRORPortrait
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "                                             Posative Antibodies List";
90    Printer.Print

100   If optInclude Then
110     Printer.Print "                                               (NEQAS Included)";
120   ElseIf optExclude Then
130     Printer.Print "                                               (NEQAS Excluded) ";

140   End If
150   Printer.Print

      'Heading Section

160   Printer.Font.Bold = True
170   For i = 1 To 108
180       Printer.Print "_";
190   Next i
200   Printer.Print
210   Printer.Print FormatString("Chart #", 10, "|"); 'chart
220   Printer.Print FormatString("A and E", 10, "|"); 'a and e
230   Printer.Print FormatString("Name", 40, "|"); 'name
240   Printer.Print FormatString("D.O.B.", 16, "|"); 'date of birth
250   Printer.Print FormatString("SampleID", 10, "|"); 'sample id
260   Printer.Print FormatString("Entry DateTime", 16, "|") 'date time of entry

270   For i = 1 To 108
280       Printer.Print "-";
290   Next i
300   Printer.Print
310   For Y = 1 To g.Rows - 1
    
        'Detail section
320     Printer.Font.Bold = False
330     Printer.Print FormatString(g.TextMatrix(Y, 1), 10, "|"); 'chart
340     Printer.Print FormatString(g.TextMatrix(Y, 2), 10, "|"); 'a and e
350     Printer.Print FormatString(g.TextMatrix(Y, 3), 40, "|"); 'name
360     Printer.Print FormatString(g.TextMatrix(Y, 4), 16, "|"); 'date of birth
370     Printer.Print FormatString(g.TextMatrix(Y, 5), 10, "|"); 'sample id
380     Printer.Print FormatString(g.TextMatrix(Y, 6), 16, "|") 'date time of entry
  
390   Next

400   Printer.EndDoc

410   For Each Px In Printers
420     If Px.DeviceName = OriginalPrinter Then
430       Set Printer = Px
440       Exit For
450     End If
460   Next

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Posative Antibodies List" & vbCr
20    If optInclude Then
30        strHeading = strHeading & "(NEQAS Included)" & vbCrLf
40    ElseIf optExclude Then
50        strHeading = strHeading & "(NEQAS Excluded)" & vbCrLf
60    End If
70    strHeading = strHeading & " " & vbCr
80    ExportFlexGrid g, Me, strHeading

End Sub

Private Sub Form_Load()

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************

End Sub


Private Sub g_Click()

10    If g.MouseRow <> 0 Then Exit Sub

20    If InStr(g.TextMatrix(0, g.col), "Date") <> 0 Then
30      g.Sort = 9
40    Else
50      If SortOrder Then
60        g.Sort = flexSortGenericAscending
70      Else
80        g.Sort = flexSortGenericDescending
90      End If
100   End If

110   SortOrder = Not SortOrder

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(g.TextMatrix(Row1, g.col)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(g.TextMatrix(Row2, g.col)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(g.TextMatrix(Row1, g.col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(g.TextMatrix(Row2, g.col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("D", d1, d2))
130   Else
140     Cmp = -Sgn(DateDiff("D", d1, d2))
150   End If

End Sub


Private Sub optExclude_Click()

10    FillG

End Sub


Private Sub optInclude_Click()

10    FillG

End Sub


