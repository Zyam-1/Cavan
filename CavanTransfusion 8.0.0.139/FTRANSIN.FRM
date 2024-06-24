VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ftransindex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfusion Index"
   ClientHeight    =   6690
   ClientLeft      =   225
   ClientTop       =   585
   ClientWidth     =   8400
   ControlBox      =   0   'False
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
   Icon            =   "FTRANSIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6690
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   7200
      Picture         =   "FTRANSIN.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5205
      Left            =   150
      TabIndex        =   4
      Top             =   1020
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   9181
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5940
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton bstart 
      Appearance      =   0  'Flat
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3150
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   510
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   128778241
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   510
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   128778241
      CurrentDate     =   36963
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   360
      TabIndex        =   8
      Top             =   6390
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   291
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
      Left            =   7080
      TabIndex        =   9
      Top             =   1860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Between Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   240
      Width           =   2925
   End
End
Attribute VB_Name = "ftransindex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bprint_Click()

      Dim X As Integer
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName

20        If Not SetFormPrinter() Then Exit Sub

30    Printer.Print "Transfusion Index Report."
40    Printer.Print dtFrom; " to "; dtTo
50    Printer.Print
60    For Y = 0 To g.Rows - 1
70      g.Row = Y
80      For X = 0 To g.Cols - 1
90        g.Col = X
100       Printer.Print Tab(Choose(X + 1, 1, 30, 37, 43, 47, 51));
110       Printer.Print g;
120     Next
130     Printer.Print
140   Next
150   Printer.EndDoc

160   For Each Px In Printers
170     If Px.DeviceName = OriginalPrinter Then
180       Set Printer = Px
190       Exit For
200     End If
210   Next

End Sub

Private Sub bstart_Click()


      Dim xmsn As Recordset
      Dim esn As Recordset
      Dim sql As String
      Dim strFromTime As String, strToTime As String
      Dim procnumber As Integer
      Dim proctotal As Integer
      Dim procmin As Integer
      Dim procmax As Integer
      Dim procaverage As Single
      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim Criteria As String
      Dim blank As Integer

10    On Error GoTo bstart_Click_Error

20    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
30    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"
40    g.Rows = 1

50    For X = 0 To List1.ListCount - 1
60      blank = True
70      procnumber = 0: proctotal = 0: procmin = 999: procmax = 0
80      sql = "select distinct name, patnum " & _
              "from patientdetails where " & _
              "datetime between '" & _
              strFromTime & "' and '" & strToTime & "' " & _
              "and [procedure] = '" & List1.List(X) & "'"
90      Set xmsn = New Recordset
100     RecOpenClientBB 0, xmsn, sql
  
110     Do While Not xmsn.EOF
120       procnumber = procnumber + 1
130       Criteria = "SELECT ISBT128 FROM product WHERE " & _
                     "patname = '" & AddTicks(xmsn!Name & "") & "' " & _
                     "AND patid = '" & xmsn("patnum") & "' " & _
                     "AND event = 'S'"
140       Set esn = New Recordset
150       RecOpenClientBB 0, esn, Criteria
160       If Not esn.EOF Then
170         esn.MoveLast
180         n = esn.RecordCount
190         proctotal = proctotal + n
200         If procmin > n Then procmin = n
210         If procmax < n Then procmax = n
220       End If
230       xmsn.MoveNext
240     Loop

250     s = List1.List(X) & vbTab
260     If procnumber <> 0 Then
270       blank = False
280       s = s & Format(procnumber) & vbTab
290       s = s & Format(proctotal) & vbTab
300       If procmin <> 999 Then s = s & Format(procmin, "##;;")
310       s = s & vbTab
320       s = s & Format(procmax, "##;;") & vbTab
330       procaverage = proctotal / procnumber
340       s = s & Format(procaverage, "#0.##;;")
350     End If
360     If Not blank Then g.AddItem s
370     g.Refresh
380   Next

390   Exit Sub

bstart_Click_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "ftransindex", "bstart_Click", intEL, strES, sql

End Sub



Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Transfusion Index" & vbCr
20    strHeading = strHeading & "From " & dtFrom.Value & " To " & dtTo.Value & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid g, Me, strHeading
End Sub

Private Sub Form_Load()

      Dim n As Integer

10    dtTo = Format(Now, "dd/mmm/yyyy")
20    dtFrom = Format(Now - 7, "dd/mmm/yyyy")

30    g.Row = 0
40    For n = 0 To 5
50      g.FixedAlignment(n) = 2
60      g.ColAlignment(n) = 2
70      g.Col = n
80      g.ColWidth(n) = Choose(n + 1, 2150, 900, 900, 900, 900, 900)
90      g.Text = Choose(n + 1, "Procedure", "Number", "Total", "Min", "Max", "Average")
100   Next

      '*****NOTE
          'loadlist might be dependent on many components so for any future
          'update in code try to keep loadlist on bottom most line of form load.
110       loadlist
      '**************************************

End Sub

Private Sub loadlist()

      Dim tb As Recordset
      Dim s As String
      Dim sql As String

10    On Error GoTo loadlist_Error

20    List1.Clear
30    sql = "Select * from Lists where ListType = 'P'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF
70      s = tb!Text & ""
80      List1.AddItem s
90      tb.MoveNext
100   Loop

110   Exit Sub

loadlist_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "ftransindex", "loadlist", intEL, strES, sql


End Sub

