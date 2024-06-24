VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   6120
   ClientLeft      =   330
   ClientTop       =   465
   ClientWidth     =   12210
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
   Icon            =   "7frmSearch.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   12210
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4500
      Picture         =   "7frmSearch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   90
      Width           =   1065
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
      Height          =   795
      Left            =   8280
      Picture         =   "7frmSearch.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3420
      Picture         =   "7frmSearch.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrint 
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
      Height          =   795
      Left            =   7050
      Picture         =   "7frmSearch.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4635
      Left            =   60
      TabIndex        =   1
      Top             =   990
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8176
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmSearch.frx":1CEA
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
      Left            =   360
      TabIndex        =   2
      Top             =   420
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   147324929
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   1890
      TabIndex        =   3
      Top             =   420
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   147324929
      CurrentDate     =   36963
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   60
      TabIndex        =   10
      Top             =   5670
      Width           =   12015
      _ExtentX        =   21193
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
      Left            =   5520
      TabIndex        =   9
      Top             =   270
      Visible         =   0   'False
      Width           =   1305
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
      Left            =   360
      TabIndex        =   4
      Top             =   150
      Width           =   2925
   End
   Begin VB.Label lblsearchfor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblsearchfor event= 'A'"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmSearch"
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

10    On Error GoTo FillG_Error

20    Grid1.Rows = 2
30    Grid1.AddItem ""
40    Grid1.RemoveItem 1

50    fdate = Format(dtFrom, "dd/MMM/yyyy")
60    tdate = Format(DateAdd("d", 1, dtTo), "dd/MMM/yyyy")

70    Select Case lblsearchfor.Caption
        Case "F", "S"
80        Grid1.FormatString = "<Unit             |<Expiry Date              |<Product            |<Group |<Sample ID.     |<Pat ID          |<Name                        |<Event            |<Op      |<Date                   |<Typenex  "
    
90        sql = "Select D.*, P.Typenex from Product as D, PatientDetails as P where " & _
                "(D.DateTime between '" & fdate & "' and '" & tdate & "') " & _
                "and D.LabNumber = P.LabNumber " & _
                "and Event = '" & lblsearchfor.Caption & "' Order By EventStart"
100       Set sn = New Recordset
110       RecOpenServerBB 0, sn, sql
120     Case Else
130       Grid1.FormatString = "<Unit                |<Expiry Date              |<Product                                                                  |<Group    ||||<Event                                   |<Op        |<Date                              |"
  
      '160       Grid1.ColWidth(0) = 90
      '170       Grid1.ColWidth(1) = 90
      '180       Grid1.ColWidth(2) = 90
      '190       Grid1.ColWidth(3) = 90
140       Grid1.ColWidth(4) = 0
150       Grid1.ColWidth(5) = 0
160       Grid1.ColWidth(6) = 0
      '190       Grid1.ColWidth(7) = 0
      '200       Grid1.ColWidth(8) = 90
      '200       Grid1.ColWidth(9) = 90
170       Grid1.ColWidth(10) = 0
    
180       sql = "Select * from Product where " & _
                "(DateTime between '" & fdate & "' and '" & tdate & "') " & _
                "and Event = '" & lblsearchfor.Caption & "' Order By EventStart"
190       Set sn = New Recordset
200       RecOpenServerBB 0, sn, sql
210   End Select

220   Do While Not sn.EOF
230     If sn!ISBT128 & "" <> "" Then
240         s = sn!ISBT128 & "" & vbTab
250     Else
260         s = sn!Number & "" & vbTab
270     End If
280     s = s & Format(sn!DateExpiry, "dd/mm/yyyy HH:mm") & vbTab
290     s = s & ProductWordingFor(sn!BarCode & "") & vbTab
300     s = s & Bar2Group(sn!GroupRh) & vbTab
310     s = s & sn!LabNumber & vbTab
320     s = s & sn!Patid & vbTab
330     s = s & sn!PatName & vbTab
340     s = s & gEVENTCODES(sn!Event).Text & vbTab
350     s = s & sn!Operator & vbTab
360     s = s & sn!DateTime & vbTab
370     If lblsearchfor.Caption = "F" Or lblsearchfor.Caption = "S" Then
380       s = s & sn!Typenex & ""
390     Else
400       s = s & Trim$(sn!Reason & "")
410     End If
420     Grid1.AddItem s
430     sn.MoveNext
440   Loop

Dim BP As BatchProduct
Dim BPs As New BatchProducts
450   BPs.LoadBetweenDates fdate, tdate, lblsearchfor.Caption
460   For Each BP In BPs
470     s = BP.BatchNumber & vbTab & _
        BP.DateExpiry & vbTab & _
        BP.Product & " (" & BP.Identifier & ")" & vbTab & _
        Bar2Group(BP.PatientGroup & "") & vbTab & _
        BP.SampleID & vbTab & _
        BP.Chart & vbTab & _
        BP.PatName & vbTab & _
        BP.Typenex & vbTab & _
        BP.UserName & vbTab & _
        BP.RecordDateTime
480     Grid1.AddItem s
490   Next

500   If Grid1.Rows > 2 Then
510     Grid1.RemoveItem 1
520   End If

530   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "frmSearch", "FillG", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 9
50    Select Case lblsearchfor.Caption
          Case "F", "S"
60            Printer.Orientation = vbPRORLandscape
70        Case Else
80            Printer.Orientation = vbPRORLandscape
90    End Select


      '****Report heading
100   Printer.Font.Bold = True
110   Printer.Print
120   Printer.Print "Search Results For "; Me.Caption
130   Printer.Print "From "; dtFrom.Value; " To "; dtTo.Value

      '****Report body

140   Select Case lblsearchfor.Caption
          Case "F", "S"
150           For i = 1 To 152
160               Printer.Print "-";
170           Next i
180           Printer.Print
190           Printer.Print FormatString("Unit", 16, "|", AlignCenter);
200           Printer.Print FormatString("Exp.Date", 10, "|", AlignCenter);
210           Printer.Print FormatString("Product", 40, "|", AlignCenter);
220           Printer.Print FormatString("Group", 8, "|", AlignCenter);
230           Printer.Print FormatString("Samp. ID", 10, "|", AlignCenter);
240           Printer.Print FormatString("Pat.ID", 10, "|", AlignCenter);
250           Printer.Print FormatString("Name", 30, "|", AlignCenter);
260           Printer.Print FormatString("Event", 25, "|")
270           Printer.Font.Bold = False
280           For i = 1 To 152
290               Printer.Print "-";
300           Next i
310           Printer.Print
320           For Y = 1 To Grid1.Rows - 1
330               Printer.Print FormatString(Grid1.TextMatrix(Y, 0), 16, "|", AlignRight);
340               Printer.Print FormatString(Grid1.TextMatrix(Y, 1), 10, "|");
350               Printer.Print FormatString(Grid1.TextMatrix(Y, 2), 40, "|");
360               Printer.Print FormatString(Grid1.TextMatrix(Y, 3), 8, "|", AlignCenter);
370               Printer.Print FormatString(Grid1.TextMatrix(Y, 4), 10, "|");
380               Printer.Print FormatString(Grid1.TextMatrix(Y, 5), 10, "|");
390               Printer.Print FormatString(Grid1.TextMatrix(Y, 6), 30, "|");
400               Printer.Print FormatString(Grid1.TextMatrix(Y, 7), 25, "|")
410           Next
420       Case Else
430           For i = 1 To 152
440               Printer.Print "-";
450           Next i
460           Printer.Print
470           Printer.Print FormatString("Unit", 16, "|", AlignCenter);
480           Printer.Print FormatString("Exp.Date", 12, "|", AlignCenter);
490           Printer.Print FormatString("Product", 68, "|", AlignCenter);
500           Printer.Print FormatString("Group", 5, "|", AlignCenter);
510           Printer.Print FormatString("Event", 22, "|", AlignCenter);
520           Printer.Print FormatString("Op", 6, "|", AlignCenter);
530           Printer.Print FormatString("Date", 20, "|", AlignCenter)
540           Printer.Font.Bold = False
550           For i = 1 To 152
560               Printer.Print "-";
570           Next i
580           Printer.Print
    
590           For Y = 1 To Grid1.Rows - 1
600               Printer.Print FormatString(Grid1.TextMatrix(Y, 0), 16, "|", AlignRight);
610               Printer.Print FormatString(Grid1.TextMatrix(Y, 1), 12, "|");
620               Printer.Print FormatString(Grid1.TextMatrix(Y, 2), 68, "|");
630               Printer.Print FormatString(Grid1.TextMatrix(Y, 3), 5, "|", AlignCenter);
640               Printer.Print FormatString(Grid1.TextMatrix(Y, 7), 22, "|");
650               Printer.Print FormatString(Grid1.TextMatrix(Y, 8), 6, "|", AlignCenter);
660               Printer.Print FormatString(Grid1.TextMatrix(Y, 9), 20, "|")
670           Next
680   End Select

690   Printer.EndDoc

700   For Each Px In Printers
710     If Px.DeviceName = OriginalPrinter Then
720       Set Printer = Px
730       Exit For
740     End If
750   Next

End Sub

Private Sub cmdSearch_Click()

10    FillG
20    If Grid1.Rows > 2 Then
30        Call AutosizeGridColumns(Grid1)
40    End If
End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String

10    strHeading = Me.Caption & vbCr
20    strHeading = strHeading & "Search Results for " & Me.Caption & " From " & dtFrom.Value & " To " & dtTo.Value & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid Grid1, Me, strHeading

End Sub

Private Sub Form_Load()

10    dtFrom = Format(Now - 7, "dd/mm/yyyy")
20    dtTo = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub Grid1_Click()

10    If Grid1.MouseRow = 0 Then
20      If InStr(Grid1.TextMatrix(0, Grid1.col), "Date") = 0 Then
30        If SortOrder Then
40          Grid1.Sort = flexSortGenericAscending
50        Else
60          Grid1.Sort = flexSortGenericDescending
70        End If
80      Else
90        Grid1.Sort = 9
100     End If
110     SortOrder = Not SortOrder
120     Exit Sub
130   End If

140   If Trim$(Grid1.TextMatrix(Grid1.row, 0)) = "" Then Exit Sub

150   With frmUnitHistory
160     .UnitNumber = Grid1.TextMatrix(Grid1.row, 0)
170     .lblExpiry = Grid1.TextMatrix(Grid1.row, 1)
180     .ProductName = Grid1.TextMatrix(Grid1.row, 2)
190     .cmdSearch = True
200     .Show 1
210   End With

End Sub

Private Sub Grid1_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String
      Dim GC As Integer

10    GC = Grid1.col

20    If Not IsDate(Grid1.TextMatrix(Row1, GC)) Then
30      Cmp = 0
40      Exit Sub
50    End If

60    If Not IsDate(Grid1.TextMatrix(Row2, GC)) Then
70      Cmp = 0
80      Exit Sub
90    End If

100   d1 = Format(Grid1.TextMatrix(Row1, GC), "dd/mmm/yyyy hh:mm:ss")
110   d2 = Format(Grid1.TextMatrix(Row2, GC), "dd/mmm/yyyy hh:mm:ss")

120   If SortOrder Then
130     Cmp = Sgn(DateDiff("D", d1, d2))
140   Else
150     Cmp = Sgn(DateDiff("D", d2, d1))
160   End If

End Sub


