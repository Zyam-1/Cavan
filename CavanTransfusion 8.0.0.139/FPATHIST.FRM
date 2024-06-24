VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fpathistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient History"
   ClientHeight    =   8610
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13335
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
   Icon            =   "FPATHIST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8610
   ScaleWidth      =   13335
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
      Height          =   900
      Left            =   6750
      Picture         =   "FPATHIST.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   60
      Width           =   900
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   5730
      TabIndex        =   9
      Top             =   2550
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin VB.PictureBox SSPanel1 
      Height          =   765
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   1905
      TabIndex        =   5
      Top             =   128
      Width           =   1965
      Begin VB.CheckBox chkSoundex 
         Alignment       =   1  'Right Justify
         Caption         =   "Soundex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1050
         TabIndex        =   8
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optChart 
         Caption         =   "Chart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1050
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optDoB 
         Caption         =   "D.o.B."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1050
         TabIndex        =   6
         Top             =   480
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7155
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12621
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
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
      AllowUserResizing=   1
      FormatString    =   $"FPATHIST.frx":0BD4
   End
   Begin VB.CommandButton bprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   9300
      Picture         =   "FPATHIST.frx":0D09
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   12210
      Picture         =   "FPATHIST.frx":1373
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5100
      Picture         =   "FPATHIST.frx":19DD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2130
      MaxLength       =   20
      TabIndex        =   0
      Top             =   368
      Width           =   2895
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   8190
      Width           =   13125
      _ExtentX        =   23151
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
      Left            =   7860
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "fpathistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNoPrevious As Boolean

Private LocalRecords As Recordset

Private strCurrentLabNumber As String

Private SortOrder As Boolean

Private Sub FillEither(ByVal tb As Recordset)

    Dim Product As String
    Dim s As String
    Dim sql As String
    Dim XMH As Recordset
    Dim TopLineFilled As Boolean
    Dim tbHistory As Recordset

10  On Error GoTo FillEither_Error

20  Product = ""

30  TopLineFilled = False

40  s = tb!LabNumber & vbTab & _
        tb!AIDR & vbTab & _
        tb!Name & vbTab & _
        tb!Patnum & vbTab & _
        tb!DoB & vbTab & _
        tb!fGroup & vbTab & _
        tb!Kell & vbTab & _
        tb!Addr1 & " " & tb!Addr2 & " " & _
        tb!Addr3 & " " & tb!addr4 & vbTab & _
        Format$(tb!SampleDate, "dd/mm/yy hh:mm") & vbTab
50  Select Case tb!requestfrom
    Case "X": s = s & "X/M"
60  Case "A": s = s & "A/N"
70  Case "G": s = s & "G+H"
80  Case "D": s = s & "DAT"
90  End Select
100 s = s & vbTab & tb!Comment & vbTab
110 If tb!Hold = "Y" Then
120     s = s & "On Hold"
130 End If

140 sql = "Select distinct isbt128, Number, BarCode, DateExpiry  from product where " & _
          "LabNumber = '" & tb!LabNumber & "'"
150 Set tbHistory = New Recordset
160 RecOpenServerBB 0, tbHistory, sql

170 If Not tbHistory.EOF Then
180     Do While Not tbHistory.EOF
190         If Len(Trim$(tbHistory!ISBT128 & "")) > 0 Then
200             sql = "Select top 1 * from Product where " & _
                      "ISBT128  = '" & tbHistory!ISBT128 & "' " & _
                      "and BarCode = '" & tbHistory!BarCode & "' " & _
                      "order by DateTime desc"
210         Else
220             sql = "Select top 1 * from Product where " & _
                      "Number = '" & tbHistory!Number & "' " & _
                      "and BarCode = '" & tbHistory!BarCode & "' and DateExpiry = '" & Format(tbHistory!DateExpiry, "dd/mmm/yyyy") & "'" & _
                      "order by DateTime desc"
230         End If
240         Set XMH = New Recordset
250         RecOpenServerBB 0, XMH, sql

260         If Not XMH.EOF Then
270             If TopLineFilled Then
280                 s = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
290             End If
300             TopLineFilled = True
310             If Len(Trim$(tbHistory!ISBT128 & "")) > 0 Then
320                 s = s & XMH!ISBT128 & vbTab
330             Else
340                 s = s & XMH!Number & vbTab
350             End If
360             If XMH!Patid = tb!Patnum And XMH!Event = "S" Then s = s & "Transfused"
370             If XMH!Patid <> tb!Patnum And XMH!Event = "S" Then s = s & "Other Patient Transfused"
380             If XMH!Patid = tb!Patnum And XMH!Event = "X" Then s = s & "X Matched"
390             If XMH!Patid <> tb!Patnum And XMH!Event = "X" Then s = s & "Other Patient X Matched"
400             If XMH!Event = "D" Then s = s & "Destroyed"
410             If XMH!Event = "T" Then s = s & "Returned to Supplier"
420             If XMH!Event = "R" Then s = s & "Restocked"
430             If XMH!Event = "P" Then s = s & "Pending"
440             If XMH!Event = "I" Then s = s & "Issued"
445             If XMH!Event = "V" Then s = s & "Electronic Issued"
450             Product = ProductWordingFor(XMH!BarCode)
460             s = s & vbTab & tb!Operator & vbTab & Product
470             g.AddItem s
480         End If
490         tbHistory.MoveNext
500     Loop
510 Else
520     g.AddItem s
530     g.row = g.Rows - 1
540     If InStr(UCase$(g.TextMatrix(g.row, 1)), "POS") Then
550         g.col = 1
560         g.CellBackColor = vbRed
570     End If
580 End If

    Dim BPs As New BatchProducts
    Dim BP As BatchProduct
590 BPs.LoadSampleID tb!LabNumber
600 For Each BP In BPs
610     If TopLineFilled Then
620         s = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
630     End If
640     TopLineFilled = True
650     s = s & BP.BatchNumber & "(" & BP.Identifier & ")" & vbTab
660     If BP.Chart & "" = tb!Patnum And BP.EventCode & "" = "S" Then s = s & "Transfused"
670     If BP.Chart & "" <> tb!Patnum And BP.EventCode & "" = "S" Then s = s & "Other Patient Transfused"
680     If BP.EventCode & "" = "D" Then s = s & "Destroyed"
690     If BP.EventCode = "T" Then s = s & "Returned to Supplier"
700     If BP.EventCode = "R" Then s = s & "Restocked"
710     If BP.EventCode = "P" Then s = s & "Pending"
720     If BP.EventCode = "I" Then s = s & "Issued"
730     If BP.EventCode = "J" Then s = s & "Expired"

740     Product = BP.Product
750     s = s & vbTab & BP.UserName & vbTab & _
            BP.Product & vbTab & _
            Format(BP.RecordDateTime, "dd/mmm/yyyy hh:nn:ss")

760     g.AddItem s
770 Next

    ''''''''''''''''''
    'Check batched products
780 sql = "Select * from BatchDetails where " & _
          "SampleID = '" & tb!LabNumber & "' Order By Date Desc"
790 Set tbHistory = New Recordset
800 RecOpenServerBB 0, tbHistory, sql

810 If Not tbHistory.EOF Then
820     Do While Not tbHistory.EOF
830         If TopLineFilled Then
840             s = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
850         End If
860         TopLineFilled = True
870         s = s & tbHistory!BatchNumber & "(" & tbHistory!Bottles & " Bottle"
880         If Val(tbHistory!Bottles & "") > 1 Then
890             s = s & "s"
900         End If
910         s = s & ")" & vbTab
920         If tbHistory!Chart & "" = tb!Patnum And tbHistory!Event & "" = "S" Then s = s & "Transfused"
930         If tbHistory!Chart & "" <> tb!Patnum And tbHistory!Event & "" = "S" Then s = s & "Other Patient Transfused"
940         If tbHistory!Event & "" = "D" Then s = s & "Destroyed"
950         If tbHistory!Event = "T" Then s = s & "Returned to Supplier"
960         If tbHistory!Event = "R" Then s = s & "Restocked"
970         If tbHistory!Event = "P" Then s = s & "Pending"
980         If tbHistory!Event = "I" Then s = s & "Issued"
990         If tbHistory!Event = "J" Then s = s & "Expired"

1000        Product = tbHistory!Product & ""
1010        s = s & vbTab & tbHistory!UserCode & vbTab & Product & vbTab & Format(tbHistory!Date, "dd/mmm/yyyy hh:mm:ss")

1020        g.AddItem s
1030        tbHistory.MoveNext
1040    Loop
1050 Else
1060    g.row = g.Rows - 1
1070    If InStr(UCase$(g.TextMatrix(g.row, 1)), "POS") Then
1080        g.col = 1
1090        g.CellBackColor = vbRed
1100    End If
1110 End If

1120 g.AddItem ""

1130 Exit Sub

FillEither_Error:

    Dim strES As String
    Dim intEL As Integer

1140 intEL = Erl
1150 strES = Err.Description
1160 LogError "fpathistory", "FillEither", intEL, strES, sql

End Sub
Private Sub cmdCancel_Click()

10  Unload Me

End Sub

Private Sub bprint_Click()

    Dim Y As Integer
    Dim Px As Printer
    Dim OriginalPrinter As String

10  On Error GoTo bprint_Click_Error

20  OriginalPrinter = Printer.DeviceName

30  If Not SetFormPrinter() Then Exit Sub

40  Printer.FontName = "Courier New"
50  Printer.Font.Size = 9
60  Printer.Orientation = vbPRORPortrait

70  Printer.Font.Bold = True
80  Printer.Print
90  Printer.Print "                                             Patient History"
100 Printer.Print
110 Printer.Print "Search results for ";
120 If chkSoundex = 1 Then
130     Printer.Print "(Soundex) ";
140 End If
150 If optName Then
160     Printer.Print "Name ";
170 ElseIf optChart Then
180     Printer.Print "Chart ";
190 Else
200     Printer.Print "D.O.B. ";
210 End If
220 Printer.Print txtName

230 Printer.Print


240 For Y = 1 To g.Rows - 1
        'g.Row = Y
        'if its parent record then write patient info in paragraph format.
250     If g.TextMatrix(Y, 0) <> "" Then
            'patient info goes here
260         Printer.Print "____________________________________________________________________________________________________________"
270         Printer.Font.Italic = True
280         Printer.Print
290         Printer.Font.Bold = False
300         Printer.Print "Lab#: ";
310         Printer.Font.Bold = True
320         Printer.Print g.TextMatrix(Y, 0); " ";    'lab no
330         Printer.Font.Bold = False
340         Printer.Print "Name: ";
350         Printer.Font.Bold = True    'Name
360         Printer.Print g.TextMatrix(Y, 2)
370         Printer.Font.Bold = False
380         Printer.Print "Antibodies: ";
390         Printer.Font.Bold = True    'Anti bodies
400         Printer.Print g.TextMatrix(Y, 1)
410         Printer.Font.Bold = False
420         Printer.Print "D.O.B. ";
430         Printer.Font.Bold = True    'DoB
440         Printer.Print g.TextMatrix(Y, 4); " ";
450         Printer.Font.Bold = False
460         Printer.Print "Group: ";
470         Printer.Font.Bold = True    'Group
480         Printer.Print g.TextMatrix(Y, 5); " ";
490         Printer.Font.Bold = False
500         Printer.Print "Date: ";
510         Printer.Font.Bold = True    'Date
520         Printer.Print Left$(Format(g.TextMatrix(Y, 8), "dd/MM/yyyy"), 10)
530         Printer.Font.Bold = False
540         Printer.Print "Request: ";
550         Printer.Font.Bold = True    'Request
560         Printer.Print g.TextMatrix(Y, 9)
570         Printer.Font.Bold = False
580         Printer.Print "Comments: ";
590         Printer.Font.Bold = True
600         Printer.Print g.TextMatrix(Y, 10)    'Comment

610         Printer.Print

620         Printer.Font.Italic = False
630         Printer.Print FormatString(" ", 14);
640         Printer.Print FormatString("Unit", 16, "|");
650         Printer.Print FormatString("Latest Event", 35, "|");
660         Printer.Print FormatString("Op.", 5, "|");
670         Printer.Print FormatString("Product", 40, "|")

680         Printer.Print FormatString(" ", 14); "----------------------------------------------------------------------------------------------"
690         Printer.Font.Bold = False
700         Printer.Print FormatString(" ", 14);
710         Printer.Print FormatString(g.TextMatrix(Y, 11), 16, "|");    'Unit
720         Printer.Print FormatString(g.TextMatrix(Y, 12), 35, "|");    't
730         Printer.Print FormatString(g.TextMatrix(Y, 13), 5, "|");    'op
740         Printer.Print FormatString(g.TextMatrix(Y, 14), 40, "|")    'Product
750     Else

            'units detail goes here
760         Printer.Print FormatString(" ", 14);
770         Printer.Print FormatString(g.TextMatrix(Y, 11), 16, "|");    'Unit
780         Printer.Print FormatString(g.TextMatrix(Y, 12), 35, "|");    't
790         Printer.Print FormatString(g.TextMatrix(Y, 13), 5, "|");    'op
800         Printer.Print FormatString(g.TextMatrix(Y, 14), 40, "|")    'Product
810     End If

820 Next

830 Printer.EndDoc

840 For Each Px In Printers
850     If Px.DeviceName = OriginalPrinter Then
860         Set Printer = Px
870         Exit For
880     End If
890 Next

900 Exit Sub

bprint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

910 intEL = Erl
920 strES = Err.Description
930 LogError "fpathistory", "bprint_Click", intEL, strES

End Sub

Private Sub cmdSearch_Click()

10  FillG

End Sub

Private Sub cmdXL_Click()
    Dim strHeading As String

10  strHeading = "Search results for "
20  If chkSoundex = 1 Then
30      strHeading = strHeading & "(Soundex) "
40  End If
50  If optName Then
60      strHeading = strHeading & "Name "
70  ElseIf optChart Then
80      strHeading = strHeading & "Chart "
90  Else
100     strHeading = strHeading & "D.O.B. "
110 End If
120 strHeading = strHeading & txtName
130 strHeading = strHeading & vbCr & " " & vbCr

140 ExportFlexGrid g, Me, strHeading

End Sub

Private Sub Form_Activate()

10  txtName.SetFocus

End Sub

Private Sub Form_Load()

10  mNoPrevious = False

20  optName.Value = True    'Name

End Sub

Private Sub g_Click()

    Dim X As Integer
    Dim Y As Integer
    Dim ySave As Integer

10  On Error GoTo g_Click_Error

20  ySave = g.row

30  For Y = 1 To g.Rows - 1
40      g.col = 0
50      g.row = Y
60      If g.CellBackColor = vbYellow Then
70          For X = 0 To g.Cols - 1
80              g.col = X
90              g.CellBackColor = 0
100         Next
110     End If
120 Next

130 g.row = ySave

140 If g.TextMatrix(g.row, 10) = "" Then Exit Sub

150 Do While g.TextMatrix(g.row, 10) <> "" And g.TextMatrix(g.row, 10) <> "Unit"
160     g.row = g.row - 1
170 Loop
180 g.row = g.row + 1
190 Do While g.TextMatrix(g.row, 10) <> ""
200     For X = 0 To g.Cols - 1
210         g.col = X
220         g.CellBackColor = vbYellow
230     Next
240     g.row = g.row + 1
250 Loop

260 Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "fpathistory", "g_Click", intEL, strES

End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

    Dim d1 As String
    Dim d2 As String

10  If Not IsDate(g.TextMatrix(Row1, g.col)) Then
20      Cmp = 0
30      Exit Sub
40  End If

50  If Not IsDate(g.TextMatrix(Row2, g.col)) Then
60      Cmp = 0
70      Exit Sub
80  End If

90  d1 = Format(g.TextMatrix(Row1, g.col), "dd/mmm/yyyy hh:mm:ss")
100 d2 = Format(g.TextMatrix(Row2, g.col), "dd/mmm/yyyy hh:mm:ss")

110 If SortOrder Then
120     Cmp = Sgn(DateDiff("D", d1, d2))
130 Else
140     Cmp = -Sgn(DateDiff("D", d1, d2))
150 End If

End Sub


Private Sub FillG()

    Dim sql As String
    Dim pos As Integer
    Dim SearchFor As String

10  On Error GoTo FillG_Error

20  g.Rows = 2
30  g.AddItem ""
40  g.RemoveItem 1

50  txtName = Trim$(txtName)
60  If txtName = "" Then Exit Sub

70  If optName Then
80      If chkSoundex.Value = 1 Then
90          SearchFor = "SOUNDEX(Name) = SOUNDEX('" & AddTicks(txtName) & "') "
100     Else
110         SearchFor = "Name like '" & AddTicks(txtName) & "%' "
120     End If
130 ElseIf optChart Then
140     SearchFor = "PatNum = '" & AddTicks(txtName) & "' "
150 Else
160     txtName = Convert62Date(txtName, BACKWARD)
170     If Not IsDate(txtName) Then Exit Sub
180     SearchFor = "DoB = '" & Format$(txtName, "dd/MMM/yyyy") & "' "
190 End If

200 g.Rows = 2
210 g.AddItem ""
220 g.RemoveItem 1

230 sql = "SELECT * FROM PatientDetails WHERE " & _
          SearchFor & _
          "ORDER BY SampleDate DESC"

240 Set LocalRecords = New Recordset
250 RecOpenClientBB 0, LocalRecords, sql

260 If LocalRecords.EOF Then
270     Answer = iMsg("No Previous details!" & vbCrLf & "Continue Search?", vbQuestion + vbYesNo)
280     If TimedOut Then Unload Me: Exit Sub
290     If Answer = vbNo Then
300         mNoPrevious = True
310         Unload Me
320     End If
330     Exit Sub
340 End If

350 g.col = 1: g.row = 0: pos = 0

360 strCurrentLabNumber = ""
370 Do While Not LocalRecords.EOF
380     FillEither LocalRecords
390     LocalRecords.MoveNext
400 Loop

410 If g.Rows > 2 Then
420     g.RemoveItem 1
430 End If

440 Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

450 intEL = Erl
460 strES = Err.Description
470 LogError "fpathistory", "FillG", intEL, strES, sql

End Sub


Public Property Get NoPrevious() As Boolean

10  NoPrevious = mNoPrevious

End Property

Private Sub txtName_LostFocus()

10  If optDoB Then
20      txtName = Convert62Date(txtName, BACKWARD)
30  End If

End Sub


