VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfused 
   Caption         =   "NetAcquire - Transfused Products"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12435
   Icon            =   "frmTransfused.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   7140
      Picture         =   "frmTransfused.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   795
      Left            =   3510
      Picture         =   "frmTransfused.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   8370
      Picture         =   "frmTransfused.frx":1376
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   795
      Left            =   4590
      Picture         =   "frmTransfused.frx":19E0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4635
      Left            =   150
      TabIndex        =   4
      Top             =   1020
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   8176
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"frmTransfused.frx":1CEA
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   450
      TabIndex        =   5
      Top             =   450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   147324929
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Top             =   450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   147324929
      CurrentDate     =   36963
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   5730
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Between Dates"
      Height          =   255
      Left            =   450
      TabIndex        =   8
      Top             =   180
      Width           =   2925
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   5745
      TabIndex        =   7
      Top             =   300
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmTransfused"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private SortOrder As Boolean

Private Sub FillG()

          Dim sn As Recordset
          Dim sql As String
          Dim s As String
          Dim fdate As String
          Dim tdate As String
          Dim BP As BatchProduct
          Dim BPs As New BatchProducts

10        On Error GoTo FillG_Error

20        Grid1.Rows = 2
30        Grid1.AddItem ""
40        Grid1.RemoveItem 1

50        fdate = Format(dtFrom, "dd/MMM/yyyy")
60        tdate = Format(DateAdd("d", 1, dtTo), "dd/MMM/yyyy")

70        sql = "Select D.*, P.Typenex from Product as D, PatientDetails as P where " & _
                "(D.DateTime between '" & fdate & "' and '" & tdate & "') " & _
                "and D.LabNumber = P.LabNumber " & _
                "and Event = 'S' Order By D.EventStart"
80        Set sn = New Recordset
90        RecOpenServerBB 0, sn, sql

100       Do While Not sn.EOF
110           If sn!ISBT128 & "" <> "" Then
120               s = sn!ISBT128 & ""
130           Else
140               s = sn!Number & ""
150           End If

160           s = s & vbTab & _
                  Format(sn!DateExpiry, "dd/mmm/yyyy HH:mm") & vbTab & _
                  ProductWordingFor(sn!BarCode & "") & vbTab & _
                  Bar2Group(sn!GroupRh) & vbTab & _
                  sn!LabNumber & vbTab & _
                  sn!Patid & vbTab & _
                  sn!PatName & vbTab & _
                  sn!Typenex & vbTab & _
                  sn!Operator & vbTab & _
                  sn!EventStart & vbTab & _
                  sn!DateTime
170           Grid1.AddItem s
180           sn.MoveNext
190       Loop

200       BPs.LoadBetweenDates fdate, tdate, "S"
210       For Each BP In BPs
220           s = BP.BatchNumber & vbTab & _
                  BP.DateExpiry & vbTab & _
                  BP.Product & " (" & BP.Identifier & ")" & vbTab & _
                  Bar2Group(BP.PatientGroup & "") & vbTab & _
                  BP.SampleID & vbTab & _
                  BP.Chart & vbTab & _
                  BP.PatName & vbTab & _
                  BP.Typenex & vbTab & _
                  BP.UserName & vbTab & _
                  BP.RecordDateTime
230           Grid1.AddItem s
240       Next

250       If Grid1.Rows > 2 Then
260           Grid1.RemoveItem 1
270       End If

280       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmTransfused", "FillG", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdPrint_Click()

          Dim Y As Integer
          Dim Px As Printer
          Dim OriginalPrinter As String
          Dim i As Integer

10        OriginalPrinter = Printer.DeviceName

20        If Not SetFormPrinter() Then Exit Sub

30        Printer.FontName = "Courier New"
40        Printer.FontSize = 9
50        Printer.Orientation = vbPRORLandscape

60        Printer.Print

          '****Report heading
70        For i = 1 To 152
80            Printer.Print "-";
90        Next i
100       Printer.Print
110       Printer.Font.Bold = True
120       Printer.Print "                                                                 Transfused Products"
130       Printer.Print
140       Printer.Print "Search Results for transfused products From "; dtFrom.Value; " To "; dtTo.Value

          '****Report body

150       Printer.Print
160       Printer.Print FormatString("Unit", 16, "|", AlignCenter);
170       Printer.Print FormatString("Exp.Date", 10, "|", AlignCenter);
180       Printer.Print FormatString("Product", 29, "|", AlignCenter);
190       Printer.Print FormatString("Group", 5, "|", AlignCenter);
200       Printer.Print FormatString("S.I.D.", 8, "|", AlignCenter);
210       Printer.Print FormatString("Pat.ID", 10, "|", AlignCenter);
220       Printer.Print FormatString("Name", 19, "|", AlignCenter);
230       Printer.Print FormatString("Typenex", 10, "|", AlignCenter);
240       Printer.Print FormatString("OP", 5, "|", AlignCenter);
250       Printer.Print FormatString("Start Date/Time", 16, "|", AlignCenter);
260       Printer.Print FormatString("End Date/Time", 16, "|", AlignCenter)

270       Printer.Font.Bold = False
280       For i = 1 To 152
290           Printer.Print "-";
300           If i = 152 Then
310               Printer.Print
320           End If
330       Next i

340       For Y = 1 To Grid1.Rows - 1
350           Printer.Print FormatString(Grid1.TextMatrix(Y, 0), 16, "|", AlignRight);
360           Printer.Print FormatString(Grid1.TextMatrix(Y, 1), 10, "|");
370           Printer.Print FormatString(Grid1.TextMatrix(Y, 2), 29, "|");
380           Printer.Print FormatString(Grid1.TextMatrix(Y, 3), 5, "|", AlignCenter);
390           Printer.Print FormatString(Grid1.TextMatrix(Y, 4), 8, "|");
400           Printer.Print FormatString(Grid1.TextMatrix(Y, 5), 10, "|");
410           Printer.Print FormatString(Grid1.TextMatrix(Y, 6), 19, "|");
420           Printer.Print FormatString(Grid1.TextMatrix(Y, 7), 10, "|");
430           Printer.Print FormatString(Grid1.TextMatrix(Y, 8), 5, "|", AlignCenter);
440           Printer.Print FormatString(Grid1.TextMatrix(Y, 9), 16, "|");
450           Printer.Print FormatString(Grid1.TextMatrix(Y, 10), 16, "|")

460       Next

          'For Y = 0 To Grid1.Rows - 1
          '  For x = 0 To 10
          '    Printer.Print Grid1.TextMatrix(Y, x);
          '    Printer.Print Tab(Choose(x + 1, 10, 23, 50, 58, 68, 78, 110, 125, 130, 152, 160));
          '  Next
          '  Printer.Print
          'Next

470       Printer.EndDoc

480       Printer.Orientation = vbPRORPortrait

490       For Each Px In Printers
500           If Px.DeviceName = OriginalPrinter Then
510               Set Printer = Px
520               Exit For
530           End If
540       Next

End Sub


Private Sub cmdSearch_Click()

10        FillG
20        If Grid1.Rows > 2 Then
30            Call AutosizeGridColumns(Grid1)
40        End If
End Sub


Private Sub cmdXL_Click()
          Dim strHeading As String
10        strHeading = "Transfused Products" & vbCr
20        strHeading = strHeading & "Search Results for transfused products From " & dtFrom.Value & " To " & dtTo.Value & vbCr
30        strHeading = strHeading & " " & vbCr
40        ExportFlexGrid Grid1, Me, strHeading

End Sub


Private Sub Form_Load()

10        dtFrom = Format(Now - 7, "dd/mm/yyyy")
20        dtTo = Format(Now, "dd/mm/yyyy")

End Sub


Private Sub Grid1_Click()

10        If Grid1.MouseRow = 0 Then
20            If InStr(Grid1.TextMatrix(0, Grid1.col), "Date") = 0 Then
30                If SortOrder Then
40                    Grid1.Sort = flexSortGenericAscending
50                Else
60                    Grid1.Sort = flexSortGenericDescending
70                End If
80            Else
90                Grid1.Sort = 9
100           End If
110           SortOrder = Not SortOrder
120           Exit Sub
130       End If

140       If Trim$(Grid1.TextMatrix(Grid1.row, 0)) = "" Then Exit Sub

150       With frmUnitHistory
160           .UnitNumber = Grid1.TextMatrix(Grid1.row, 0)
170           .lblExpiry = Grid1.TextMatrix(Grid1.row, 1)
180           .ProductName = Grid1.TextMatrix(Grid1.row, 2)
190           .cmdSearch = True
200           .Show 1
210       End With

End Sub


Private Sub Grid1_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String
          Dim GC As Integer

10        GC = Grid1.col

20        If Not IsDate(Grid1.TextMatrix(Row1, GC)) Then
30            Cmp = 0
40            Exit Sub
50        End If

60        If Not IsDate(Grid1.TextMatrix(Row2, GC)) Then
70            Cmp = 0
80            Exit Sub
90        End If

100       d1 = Format(Grid1.TextMatrix(Row1, GC), "dd/mmm/yyyy hh:mm:ss")
110       d2 = Format(Grid1.TextMatrix(Row2, GC), "dd/mmm/yyyy hh:mm:ss")

120       If SortOrder Then
130           Cmp = Sgn(DateDiff("D", d1, d2))
140       Else
150           Cmp = Sgn(DateDiff("D", d2, d1))
160       End If

End Sub


