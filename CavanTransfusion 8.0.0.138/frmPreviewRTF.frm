VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmPreviewRTF 
   Caption         =   "NetAcquire - Print Preview"
   ClientHeight    =   7800
   ClientLeft      =   75
   ClientTop       =   615
   ClientWidth     =   13740
   Icon            =   "frmPreviewRTF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   13740
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   11310
      TabIndex        =   8
      Text            =   "cmbDepartment"
      Top             =   450
      Width           =   2265
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   11310
      TabIndex        =   7
      Top             =   90
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   38582
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4905
      Left            =   11310
      TabIndex        =   6
      Top             =   840
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   8652
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      FormatString    =   "<Time         |<Identifier           |<PaperSize |<PageNumber"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   12570
      Picture         =   "frmPreviewRTF.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   645
      Left            =   11400
      Picture         =   "frmPreviewRTF.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   90
      SmallChange     =   100
      TabIndex        =   0
      Top             =   7200
      Width           =   10815
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7155
      LargeChange     =   500
      Left            =   10920
      SmallChange     =   100
      TabIndex        =   2
      Top             =   60
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Height          =   7125
      Left            =   90
      ScaleHeight     =   7065
      ScaleWidth      =   10755
      TabIndex        =   1
      Top             =   60
      Width           =   10815
      Begin RichTextLib.RichTextBox rtb 
         Height          =   7035
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   12409
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmPreviewRTF.frx":159E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   10
      Top             =   7590
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image imgDown 
      Height          =   465
      Left            =   11370
      Picture         =   "frmPreviewRTF.frx":161E
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   480
   End
   Begin VB.Image imgUp 
      Height          =   480
      Left            =   13110
      Picture         =   "frmPreviewRTF.frx":1A60
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   480
   End
   Begin VB.Label lblPages 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Page 99 of 99"
      Height          =   285
      Left            =   11820
      TabIndex        =   9
      Top             =   5940
      Width           =   1245
   End
End
Attribute VB_Name = "frmPreviewRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDept As String
Private pSampleID As String
Private pPaperSize As String

Private pPageCounter As Integer 'First page is page #1

Public MaxLines As Long

Private Sub FillG()

      Dim tb As Recordset
      Dim tbD As Recordset
      Dim sql As String
      Dim Department As String

10    On Error GoTo FillG_Error

20    grd.Rows = 2
30    grd.AddItem ""
40    grd.RemoveItem 1

50    Select Case cmbDepartment
        Case "AHG QC Report": Department = "TA"
60      Case "Centrifuge QC Report": Department = "TC"
70      Case "Grouping Cards QC": Department = "TG"
80      Case "Group & Screen Form": Department = "GH"
90      Case "Cross Match Form": Department = "XM"
100     Case "Cross Match Label": Department = "XL"
110     Case "Ante-Natal Form": Department = "AN"
120     Case "Cord Blood Form": Department = "CD"
130     Case "Batch Issue Form": Department = "BF"
140     Case "Cross Match Label (PDF)": Department = "PD"
150   End Select

160   CheckPrintedTableInDb

170   sql = "Select distinct DateTime from Printed where " & _
            "DateTime between '" & Format$(dt, "dd/mmm/yyyy" & " 00:00:00") & "' " & _
            "and '" & Format$(dt, "dd/mmm/yyyy" & " 23:59:59") & "' " & _
            "and Dept = '" & Department & "'"
180   Set tb = New Recordset
190   RecOpenServerBB 0, tb, sql
200   Do While Not tb.EOF
210     sql = "Select LabNumber, PaperSize from Printed where DateTime = '" & Format$(tb!DateTime, "dd/mmm/yyyy hh:mm:ss") & "' and Dept = '" & Department & "'"
220     Set tbD = New Recordset
230     RecOpenServerBB 0, tbD, sql
240     If Not tbD.EOF Then
250     grd.AddItem Format$(tb!DateTime, "hh:mm:ss") & vbTab & _
                    tbD!LabNumber & vbTab & _
                    tbD!PaperSize & ""
260     End If
270     tb.MoveNext
280   Loop

290   If grd.Rows > 2 Then
300     grd.RemoveItem 1
310   End If

320   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmPreviewRTF", "FillG", intEL, strES, sql


End Sub

Private Function GetCurrentPage() As Integer

      Dim s() As String

10    s = Split(lblPages, " ")

20    GetCurrentPage = Val(s(1))

End Function

Public Sub PrintRTB()

      Dim sql As String
      Dim n As Integer
      Dim dtNow As String

10    On Error GoTo PrintRTB_Error

20    dtNow = Format$(Now, "dd/mmm/yyyy hh:mm:ss")

30    For n = 1 To pPageCounter
40      rtb(n).SelStart = 0
      '        rtb(n).upto vbCr
50      rtb(n).SelFontSize = 6
60      rtb(n).SelBold = False
70      rtb(n).SelUnderline = False
80      rtb(n).SelItalic = False
90      rtb(n).SelColor = vbBlack
100     rtb(n).SelAlignment = rtfRight
110     rtb(n).SelText = "Page " & Format$(n) & " of " & Format$(pPageCounter) & vbCrLf
120     rtb(n).SelAlignment = rtfLeft
  
      '  Printer.Orientation = vbPRORLandscape
130     Printer.Print ""
140     rtb(n).SelPrint Printer.hDC

150     If cmdPrint.Caption = "&Print" Then
          'Dont do this if re-print
160       CheckPrintedTableInDb
170       rtb(n).TextRTF = AddTicks(rtb(n).TextRTF)
180       sql = "INSERT INTO Printed " & _
                "(LabNumber, DateTime, Operator, Dept, PaperSize, RTF, PageNumber) VALUES " & _
                "('" & pSampleID & "', " & _
                " '" & dtNow & "', " & _
                " '" & UserName & "', " & _
                " '" & pDept & "', " & _
                " '" & pPaperSize & "', " & _
                " '" & rtb(n).TextRTF & "', " & _
                " '" & n & "' )"
190       CnxnBB(0).Execute sql
  
200     End If
210   Next

220   Exit Sub

PrintRTB_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmPreviewRTF", "PrintRTB", intEL, strES, sql

End Sub
Public Sub CheckPrintedTableInDb()

      Dim sql As String
      Dim tb As Recordset
      Dim Design(0 To 6, 0 To 1) As String
      Dim n As Integer
      Dim Found As Boolean
      Dim f As Integer

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist "
      'if it has a record then the table does exist.

10    On Error GoTo CheckPrintedTableInDb_Error

20    sql = "SELECT Name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = 'Printed'"
30    Set tb = New Recordset
40    Set tb = CnxnBB(0).Execute(sql)

50    If tb.EOF Then 'There is no table 'Printed' in database
60      sql = "CREATE TABLE Printed " & _
              "( LabNumber nvarchar(10) NULL, " & _
              "  DateTime  datetime NULL, " & _
              "  Operator  nvarchar(50) NULL, " & _
              "  Dept  nvarchar(2) NULL, " & _
              "  PaperSize nvarchar(20) NULL, " & _
              "  RTF ntext NULL, " & _
              "  PageNumber int NULL )"
70      CnxnBB(0).Execute sql
80    Else 'are all the fields there?
90      Design(0, 0) = "LabNumber"
100     Design(0, 1) = "nvarchar(10)"
110     Design(1, 0) = "DateTime"
120     Design(1, 1) = "datetime"
130     Design(2, 0) = "Operator"
140     Design(2, 1) = "nvarchar(50)"
150     Design(3, 0) = "Dept"
160     Design(3, 1) = "nvarchar(2)"
170     Design(4, 0) = "PaperSize"
180     Design(4, 1) = "nvarchar(20)"
190     Design(5, 0) = "RTF"
200     Design(5, 1) = "ntext"
210     Design(6, 0) = "PageNumber"
220     Design(6, 1) = "int"
230     sql = "Select top 1 * from [Printed]"
240     Set tb = New Recordset
250     RecOpenServerBB 0, tb, sql
260     For n = 0 To UBound(Design)
270       Found = False
280       For f = 0 To tb.Fields.Count - 1
290         If UCase$(Design(n, 0)) = UCase$(tb.Fields(f).Name) Then
300           Found = True
310           Exit For
320         End If
330       Next
340       If Not Found Then

350         sql = "ALTER TABLE [Printed] " & _
                  "ADD [" & Design(n, 0) & "] " & _
                  Design(n, 1) & " NULL"
360         CnxnBB(0).Execute sql
370         If Err <> 0 Then
380           iMsg "Cannot Add " & Design(n, 0) & " to Printed", vbCritical
390           If TimedOut Then Unload Me: Exit Sub
400         End If

410       End If
420     Next
430   End If

440   Exit Sub

CheckPrintedTableInDb_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmPreviewRTF", "CheckPrintedTableInDb", intEL, strES, sql

End Sub

Public Sub WriteText(ByVal Txt As String)

  
10      Debug.Print "Pre "; rtb(pPageCounter).GetLineFromChar(10000000#),
20      rtb(pPageCounter).SelText = Txt
30      Debug.Print " Post "; rtb(pPageCounter).GetLineFromChar(10000000#)

    
40    rtb(pPageCounter).SelStart = 10000000# 'Len(rtb(pPageCounter).Text)
50    If rtb(pPageCounter).GetLineFromChar(10000000#) > 44 Then
60      ForceNewPage
70    End If
80    cmdPrint.Caption = "&Print"

End Sub

Public Sub WriteFormattedText(ByVal Txt As String, _
                              Optional ByVal blnBold As Variant, _
                              Optional ByVal lngSize As Variant, _
                              Optional ByVal lngColour As Variant, _
                              Optional ByVal blnUnderline As Variant, _
                              Optional ByVal strFontName As Variant)
    
10    If Txt = "" Then Exit Sub

20    If Not IsMissing(blnBold) Then
30      rtb(pPageCounter).SelBold = blnBold
40    End If

50    If Not IsMissing(lngSize) Then
60      rtb(pPageCounter).SelFontSize = lngSize
70    End If

80    If Not IsMissing(lngColour) Then
90      rtb(pPageCounter).SelColor = lngColour
100   End If

110   If Not IsMissing(blnUnderline) Then
120     rtb(pPageCounter).SelUnderline = blnUnderline
130   End If

140   If Not IsMissing(strFontName) Then
150     rtb(pPageCounter).SelFontName = strFontName
160   End If

170   If Right$(Txt, 1) = ";" Then
180     rtb(pPageCounter).SelText = Left$(Txt, Len(Txt) - 1)
190   Else
200     rtb(pPageCounter).SelText = Txt & vbCrLf
210   End If

220   cmdPrint.Caption = "&Print"
    
230   rtb(pPageCounter).SelStart = 10000000# 'Len(rtb(pPageCounter).Text)
240   If rtb(pPageCounter).GetLineFromChar(10000000#) > MaxLines + 1 Then
250     ForceNewPage
260   End If

End Sub

Public Property Get LineCounter() As Long

10    LineCounter = rtb(pPageCounter).GetLineFromChar(10000000#)

End Property


Public Sub Clear()

10    rtb(pPageCounter) = ""

End Sub

Public Property Let fName(ByVal strNewValue As String)

10    rtb(pPageCounter).SelFontName = strNewValue

End Property

Public Property Let FBold(ByVal blnNewValue As Boolean)

10    rtb(pPageCounter).SelBold = blnNewValue

End Property





Public Sub FDetails(Optional ByVal blnBold As Variant, _
                    Optional ByVal lngSize As Variant, _
                    Optional ByVal lngColour As Variant, _
                    Optional ByVal blnUnderline As Variant, _
                    Optional ByVal strFontName As Variant)
     
10    If Not IsMissing(blnBold) Then
20      rtb(pPageCounter).SelBold = blnBold
30    End If

40    If Not IsMissing(lngSize) Then
50      rtb(pPageCounter).SelFontSize = lngSize
60    End If

70    If Not IsMissing(lngColour) Then
80      rtb(pPageCounter).SelColor = lngColour
90    End If

100   If Not IsMissing(blnUnderline) Then
110     rtb(pPageCounter).SelUnderline = blnUnderline
120   End If

130   If Not IsMissing(strFontName) Then
140     rtb(pPageCounter).SelFontName = strFontName
150   End If

End Sub

Public Property Let FColour(ByVal lngNewValue As Long)

10    rtb(pPageCounter).SelColor = lngNewValue

End Property


Public Property Let FUnderline(ByVal blnNewValue As Boolean)

10    rtb(pPageCounter).SelUnderline = blnNewValue

End Property


Public Property Let FSize(ByVal lngNewValue As Long)

10    rtb(pPageCounter).SelFontSize = lngNewValue

End Property


Private Sub cmbDepartment_Click()

10    FillG

End Sub

Private Sub cmbDepartment_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

10    PrintRTB

End Sub

Public Sub AdjustPaperSize(ByVal strSizeOrient As String)

      Dim xx As Integer
      Dim w As Long
      Dim H As Long

      'A format strSizeOrient of eg "100x70" is Width*Height in mm

      'A4, 210 x 297 mm
      'A5, 148 x 210 mm

      'A twip is 1/20 of a printer’s point
      '(1,440 twips equal one inch,
      'and 567 twips equal one centimeter).

10    pPaperSize = strSizeOrient
20    If pPageCounter = 0 Then pPageCounter = 1

30    Select Case UCase$(strSizeOrient)
  
        Case "A4PORT":
40        rtb(pPageCounter).Width = 567 * 21
50        rtb(pPageCounter).Height = 567 * 29.7
  
60      Case "A4LAND":
70        rtb(pPageCounter).Width = 567 * 29.7
80        rtb(pPageCounter).Height = 567 * 21
  
90      Case "A5PORT":
100       rtb(pPageCounter).Width = 567 * 14.8
110       rtb(pPageCounter).Height = 567 * 21
  
120     Case "A5LAND":
130       rtb(pPageCounter).Width = 567 * 21
140       rtb(pPageCounter).Height = 567 * 14.8
  
150     Case Else:
160       xx = InStr(UCase$(strSizeOrient), "X")
170       If xx > 1 Then
180         w = Val(Left$(strSizeOrient, xx - 1))
190         H = Val(Mid$(strSizeOrient, xx + 1))
200         rtb(pPageCounter).Width = 56.7 * w
210         rtb(pPageCounter).Height = 56.7 * H
220       Else
            'Set to A4PORT
230         rtb(pPageCounter).Width = 567 * 21
240         rtb(pPageCounter).Height = 567 * 29.7
250       End If
    
260   End Select
   
270   rtb(pPageCounter).Visible = True

280   VScroll1.max = rtb(1).Height
290   HScroll1.max = rtb(1).Width

300   MaxLines = rtb(1).Height / (TextHeight("H") * 1.3)

End Sub

Private Sub dt_CloseUp()

10    FillG

End Sub


Private Sub Form_Activate()

      Dim n As Integer

10    lblPages = "Page 1 of " & pPageCounter

20    For n = 2 To pPageCounter
30      rtb(n).Visible = False
40    Next

End Sub

Private Sub Form_Load()

      'See cmdCelect_Click for list of Department Codes
10    With cmbDepartment
20      .Clear
30      .AddItem "AHG QC Report"
40      .AddItem "Centrifuge QC Report"
50      .AddItem "Grouping Cards QC"
60      .AddItem "Group & Screen Form"
70      .AddItem "Cross Match Form"
80      .AddItem "Cross Match Label"
90      .AddItem "Cross Match Label (PDF)"
100     .AddItem "Ante-Natal Form"
110     .AddItem "Cord Blood Form"
120     .AddItem "Batch Issue Form"
130   End With

      'Set dtPicker to todays date
140   dt = Format$(Now, "dd/mm/yyyy")

      'grd.col(2) = PaperSize - no need to display this
150   grd.ColWidth(2) = 0
      'grd.col(3) = PageNumber - no need to display this
160   grd.ColWidth(3) = 0

170   pPageCounter = 1

180   lblPages = "Page 1 of 1"

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    pPageCounter = 1

End Sub

Private Sub grd_Click()

      Dim tbDistinct As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Y As Integer
      Dim ySave As Integer

10    On Error GoTo grd_Click_Error

20    If grd.MouseRow = 0 Then Exit Sub
30    If grd.TextMatrix(1, 0) = "" Then Exit Sub

40    ySave = grd.Row
50    For Y = 1 To grd.Rows - 1
60      grd.Row = Y
70      If grd.CellBackColor = vbRed Then
80        For n = 0 To grd.Cols - 1
90          grd.Col = n
100         grd.CellBackColor = &H80000018
110       Next
120       Exit For
130     End If
140   Next
150   grd.Row = ySave
160   For n = 0 To grd.Cols - 1
170     grd.Col = n
180     grd.CellBackColor = vbRed
190   Next

200   CheckPrintedTableInDb

210   For n = pPageCounter To 2 Step -1
220     Unload rtb(n)
230   Next
240   pPageCounter = 1

250   sql = "Select distinct DateTime, LabNumber from Printed where " & _
            "DateTime = '" & Format$(dt, "dd/mmm/yyyy") & " " & _
            Format$(grd.TextMatrix(grd.Row, 0), "hh:mm:ss") & "' " & _
            "Order by DateTime"
260   Set tbDistinct = New Recordset
270   RecOpenServerBB 0, tbDistinct, sql
280   Do While Not tbDistinct.EOF
290     sql = "Select * from Printed where " & _
              "DateTime = '" & Format$(tbDistinct!DateTime, "dd/mmm/yyyy hh:mm:ss") & "' " & _
              ""
300     If Not IsNull(tbDistinct!LabNumber) Then
310       sql = sql & "and LabNumber = '" & tbDistinct!LabNumber & "'"
320     End If
330     Set tb = New Recordset
340     RecOpenServerBB 0, tb, sql
350     Do While Not tb.EOF
360       If tb!pagenumber <> 1 Then
370         ForceNewPage
380       End If
390       AdjustPaperSize tb!PaperSize & ""
400       rtb(pPageCounter).TextRTF = tb!rtf & ""
410       tb.MoveNext
420     Loop
430     tbDistinct.MoveNext
440   Loop

450   lblPages = "Page 1 of " & pPageCounter

460   cmdPrint.Caption = "Re-&Print"

470   Exit Sub

grd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmPreviewRTF", "grd_Click", intEL, strES, sql


End Sub

Private Sub HScroll1_Change()

10    rtb(pPageCounter).Left = -HScroll1.Value

End Sub


Private Sub HScroll1_Scroll()

10    rtb(pPageCounter).Left = -HScroll1.Value

End Sub


Private Sub imgUp_Click()

      Dim n As Long
      Dim p As Integer

10    p = GetCurrentPage()

20    If p = pPageCounter Then Exit Sub

30    lblPages = "Page " & Format$(p + 1) & " of " & pPageCounter

40    For n = 1 To pPageCounter
50      rtb(n).Visible = False
60    Next
  
70    rtb(p + 1).Top = 0
80    VScroll1.Value = 0
90    rtb(p + 1).Visible = True

End Sub

Private Sub imgDown_Click()

      Dim n As Long
      Dim p As Integer

10    p = GetCurrentPage()

20    If p = 1 Then Exit Sub

30    lblPages = "Page " & Format$(p - 1) & " of " & pPageCounter

40    For n = 1 To pPageCounter
50      rtb(n).Visible = False
60    Next
  
70    rtb(p - 1).Top = 0
80    VScroll1.Value = 0
90    rtb(p - 1).Visible = True

End Sub

Private Sub VScroll1_Change()

      Dim p As Integer

10    p = GetCurrentPage()
  
20    rtb(p).Top = -VScroll1.Value

      'For n = 0 To pPageCounter
      '  rtb(n).Visible = False
      '  rtb(p).Visible = True
      'Next

End Sub


Private Sub VScroll1_Scroll()

      Dim p As Integer

10    p = GetCurrentPage()
  
20    rtb(p).Top = -VScroll1.Value
      'Dim n As Long
      '
      'For n = 0 To pPageCounter
      '  rtb(n).Top = -VScroll1.Value + (rtb(n).Height * n)
      '  rtb(n).Visible = True
      'Next

End Sub




Public Property Let Dept(ByVal strNewValue As String)

10    pDept = Left$(strNewValue, 2)

End Property

Public Property Let SampleID(ByVal strNewValue As String)

10    pSampleID = strNewValue

End Property

Public Function ForceNewPage() As Integer

10    pPageCounter = pPageCounter + 1

20    Load rtb(pPageCounter)

30    ForceNewPage = pPageCounter
40    AdjustPaperSize pPaperSize

50    Clear

End Function
