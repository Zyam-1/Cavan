VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Print Handler. Version 1.0.0"
   ClientHeight    =   5640
   ClientLeft      =   5190
   ClientTop       =   7905
   ClientWidth     =   6330
   ForeColor       =   &H00400000&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6330
   Begin VB.CommandButton cmdRefreshPrinters 
      Caption         =   "Refresh Printer List"
      Height          =   495
      Left            =   2010
      TabIndex        =   17
      Top             =   4290
      Width           =   2265
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   10635
      Left            =   6330
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   18759
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"fMain.frx":030A
   End
   Begin VB.Frame Frame2 
      Caption         =   "ZetaFax"
      Height          =   1665
      Left            =   90
      TabIndex        =   11
      Top             =   2340
      Width           =   5985
      Begin VB.Label lblDocument 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   1230
         Width           =   5595
      End
      Begin VB.Label lblZSubmit 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   210
         TabIndex        =   14
         Top             =   570
         Width           =   5595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Document Folder"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ZSubmit Folder"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1065
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5265
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "16/06/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "21:42"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8229
            Text            =   "Custom Software"
            TextSave        =   "Custom Software"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Options"
      Height          =   1125
      Left            =   1380
      TabIndex        =   5
      Top             =   960
      Width           =   3645
      Begin VB.TextBox txtMoreThan 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1050
         TabIndex        =   9
         Text            =   "18"
         Top             =   270
         Width           =   525
      End
      Begin VB.OptionButton optSideBySide 
         Caption         =   "Print Side by Side"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optSecondPage 
         Caption         =   "Print on Second Page"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "If more than xxxxxx  Biochemistry Results then "
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   3315
      End
   End
   Begin VB.OptionButton oNREnabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2070
      TabIndex        =   4
      Top             =   690
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.OptionButton oNRDisabled 
      Caption         =   "Disabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3150
      TabIndex        =   3
      Top             =   690
      Width           =   1065
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   3585
      Left            =   120
      ScaleHeight     =   3525
      ScaleWidth      =   5745
      TabIndex        =   1
      Top             =   7380
      Width           =   5805
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   225
      Left            =   510
      TabIndex        =   0
      Top             =   30
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   4950
      Top             =   0
   End
   Begin VB.Label lblHaem 
      Alignment       =   2  'Center
      Caption         =   "Haematology Age/Sex Related Normal Ranges are"
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Top             =   270
      Width           =   2175
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mPrinters 
         Caption         =   "&Printers"
      End
      Begin VB.Menu mHaemNormal 
         Caption         =   "&Haem Normal Ranges"
      End
      Begin VB.Menu mnuFAX 
         Caption         =   "FAX Log On Details"
         Begin VB.Menu mnuAppName 
            Caption         =   "&App Name"
         End
         Begin VB.Menu mnuAppPath 
            Caption         =   "App &Path"
         End
      End
      Begin VB.Menu mnuPrintSplit 
         Caption         =   "Printer &Splits"
      End
      Begin VB.Menu mNull 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean

Private EnableTimer As Boolean

Private Function GetPrinterFromWard(ByVal Ward As String) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo GetPrinterFromWard_Error
20        sql = "Select PrinterAddress from Wards where " & _
                "[Text] = '" & AddTicks(Ward) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            GetPrinterFromWard = Trim$(tb!PrinterAddress & "")
70        Else
80            GetPrinterFromWard = ""
90        End If
100       Exit Function

GetPrinterFromWard_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMain", "GetPrinterFromWard", intEL, strES, sql


End Function

Private Sub ProcessPrintQueue()

          Dim tb As Recordset
          Dim tbCopyTo As Recordset
          Dim sql As String
          Dim SampleID As String
          Dim Initiator As String
          Dim ForcedPrintDone As Boolean
          Dim PrintThis As Boolean
          Dim xFound As Boolean
          Dim Px As Printer
          Dim Printed As Boolean
          Dim Faxed As Boolean
          Dim VV As Integer
          Dim strRePrintCounter As String

10        On Error GoTo ProcessPrintQueue_Error
20        sql = "SELECT * FROM PrintPending " & _
                "ORDER BY DateTimeOfRecord ASC"
30        ForcedPrintDone = False
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
60        Do While Not tb.EOF
70            RP.SampleID = tb!SampleID & ""
              '_____________________________________________
              'Clear RTB ready for next sample report
72            rtb = ""
74            rtb.SelText = ""
              '_____________________________________________
80            RP.Department = tb!Department & ""
90            RP.UsePrinter = Trim$(tb!UsePrinter & "")
100           strRePrintCounter = Trim$(tb!ReprintReportCounter & "")
110           If Len(strRePrintCounter) = 0 Then    'Create report, print & store
120               If IsNull(tb!ThisIsCopy) Then
130                   gPrintCopyReport = 0
140               Else
150                   gPrintCopyReport = tb!ThisIsCopy
160               End If
170               Printed = False
180               Faxed = False
190               VV = gDONTCARE
200               If Not IsNull(tb!PrintOnCondition) Then
210                   Select Case tb!PrintOnCondition
                      Case 1: VV = gVALID
220                   Case 2: VV = gDONTCARE
230                   Case 3: VV = gDONTCARE
240                   End Select
250               End If
260               RP.ThisIsCopy = False
270               RP.SendCopyTo = ""
280               RP.Initiator = Trim$(tb!Initiator & "")
290               RP.Ward = tb!Ward & ""
300               RP.Clinician = tb!Clinician & ""
310               RP.GP = tb!GP & ""
320               RP.FaxNumber = tb!FaxNumber & ""
                  
330               RP.PrintAction = tb!PrintAction & ""
                  
                  'Disable Ward print if in micro
340               If RP.Department <> "M" Then
350                   If RP.UsePrinter = "" Then
360                       RP.UsePrinter = GetPrinterFromWard(RP.Ward)
370                   End If
380               End If
390               If Trim$(RP.FaxNumber & "") = "" Then
400                   pForcePrintTo = RP.UsePrinter
410                   If pForcePrintTo <> "" Then
420                       OriginalPrinter = Printer.DeviceName
430                       xFound = False
440                       For Each Px In Printers
450                           If UCase$(Px.DeviceName) = UCase$(pForcePrintTo) Then
460                               Set Printer = Px
470                               xFound = True
480                               Exit For
490                           End If
500                       Next
510                       If Not xFound Then
520                           LogError "frmMain", "ProcessPrintQueue", 0, "Cant use forced printer (" & pForcePrintTo & ")"
530                           pForcePrintTo = ""
540                       End If
550                   End If

560                   PrintThis = False
570                   If Trim$(tb!UsePrinter & "") = "" Then
580                       PrintThis = True
590                   Else
600                       If Not ForcedPrintDone Then
610                           PrintThis = True
620                           ForcedPrintDone = True
630                       End If
640                   End If

650                   If PrintThis Then
660                       PrintRecord VV
670                       Printed = True
680                   End If
690               Else
                      'Zyam commented this 15-06-24
700                   'FaxRecord
710                   'Faxed = True
                      'Zyam 15-06-24
720               End If
730               SampleID = Format$(tb!SampleID)
740               If RP.Department = "M" Then
750                   SampleID = Format$(Val(tb!SampleID) + sysOptMicroOffset(0))
760               End If
770               sql = "SELECT * FROM SendCopyTo WHERE " & _
                        "SampleID = '" & SampleID & "' "
780               Set tbCopyTo = New Recordset
790               RecOpenServer 0, tbCopyTo, sql
800               Do While Not tbCopyTo.EOF
810                   RP.FaxNumber = tbCopyTo!Destination
820                   RP.ThisIsCopy = True
830                   If Trim$(tbCopyTo!Ward & "") <> "" Then
840                       RP.SendCopyTo = tbCopyTo!Ward
850                   ElseIf Trim$(tbCopyTo!Clinician & "") <> "" Then
860                       RP.SendCopyTo = tbCopyTo!Clinician
870                   ElseIf Trim$(tbCopyTo!GP & "") <> "" Then
880                       RP.SendCopyTo = tbCopyTo!GP
890                   End If
900                   If UCase$(tbCopyTo!Device & "") = "FAX" Then
910                       'FaxRecord
920                       Faxed = True
930                   Else
940                       PrintRecord VV
950                       Printed = True
960                   End If
970                   tbCopyTo.MoveNext
980               Loop
990               If Faxed Or Printed Then
1000                  sql = "Delete from PrintPending where " & _
                            "SampleID = '" & RP.SampleID & "' " & _
                            "and Department = '" & RP.Department & "'"
1010                  Cnxn(0).Execute sql
1020              End If
1030          Else    'Reprint previously printed reports
                  'get Printer to use
1040              If RP.UsePrinter <> "" Then    'Force Print
1050                  OriginalPrinter = Printer.DeviceName
1060                  xFound = False
1070                  For Each Px In Printers
1080                      If UCase$(Px.DeviceName) = UCase$(RP.UsePrinter) Then
1090                          Set Printer = Px
1100                          xFound = True
1110                          Exit For
1120                      End If
1130                  Next
1140                  If Not xFound Then
1150                      LogError "frmMain", "ProcessPrintQueue", 0, "Cant use forced printer (" & pForcePrintTo & ")"
1160                      pForcePrintTo = ""
1170                  End If
1180              Else    'get default dept printer
1190                  Select Case UCase(RP.Department)
                      Case "B": SetPrinter ("CHBIO")
1200                  Case "H": SetPrinter ("CHHAEM")
1210                  Case "D": SetPrinter ("CHCOAG")    'Coag
1220                  Case "M": SetPrinter ("MICRO")    'Micro
1230                  Case "I":    'imm
1240                  Case "E":    'Blood gas
1250                  Case "X":    'externals
1260                  Case Else: SetPrinter ("CHBIO")
1270                  End Select
1280              End If
                  'get Report from Reports table into RTF control
1290              FillReport (strRePrintCounter)
                  'print report
1300              With frmMain.rtb
1310                  .SelStart = 0
1320                  .SelPrint Printer.hDC
1330              End With
                  'Reset to original printer
1340              ReSetPrinter
                  'Delete job
1350              sql = "Delete from PrintPending where " & _
                        "SampleID = '" & RP.SampleID & "' " & _
                        "and ReprintReportCounter = '" & strRePrintCounter & "'"
1360              Cnxn(0).Execute sql
1370          End If
1380          tb.MoveNext    'Next print job
1390      Loop
            '_____________________________________________
            'Clear RTB at the end of any printing run and be ready for next sample report
1392            rtb = ""
1396            rtb.SelText = ""
            '_____________________________________________
1400      Exit Sub

ProcessPrintQueue_Error:

          Dim strES As String
          Dim intEL As Integer

1410      intEL = Erl
1420      strES = Err.Description
1430      LogError "frmMain", "ProcessPrintQueue", intEL, strES, sql

End Sub


Public Sub DrawPicture(ByVal Chart As String)

      Dim tb As Recordset
      Dim sn As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Counter As Integer
      Dim LatestDate As String
      Dim EarliestDate As String
      Dim NumberOfDays As Long
      Dim X As Integer
      Dim Y As Integer
      Dim PixelsPerDay As Single
      Dim PixelsPerPointY As Integer
      Dim FirstDayFilled As Boolean
      Dim CR As CoagResult
      Dim CRs As CoagResults

10    On Error GoTo DrawPicture_Error
20    pb.Cls
30    pb.Picture = LoadPicture("")
40    For n = 1 To 365
50        gData(n, 1) = 0
60        gData(n, 2) = 0
70        gData(n, 3) = 0
80    Next
90    pb.Font.Bold = True
100   For Y = 0 To 6
110       pb.CurrentY = (pb.Height - ((pb.Height - 1.5 * pb.TextHeight("W")) / 6) * Y) - 1.5 * TextHeight("W")
120       pb.CurrentX = 0
130       pb.ForeColor = vbBlue
140       pb.Print Format(Y)
150       pb.CurrentY = (pb.Height - ((pb.Height - 1.5 * pb.TextHeight("W")) / 6) * Y) - 1.5 * TextHeight("W")
160       pb.CurrentX = pb.Width - pb.TextWidth("WW")
170       pb.ForeColor = vbGreen
180       pb.Print Format(Y * 2)
190   Next
200   sql = "select sampleid, rundate from demographics where " & _
            "chart = '" & Chart & "' " & _
            "order by rundate desc"
210   Set sn = New Recordset
220   RecOpenClient 0, sn, sql
230   If sn.EOF Then Exit Sub
240   FirstDayFilled = False
250   Counter = 0
260   Do While Not sn.EOF
270       sql = "Select * from HaemResults where " & _
                "SampleID = '" & sn!SampleID & "'"
280       Set tb = New Recordset
290       RecOpenClient 0, tb, sql
300       Set CRs = New CoagResults
310       Set CRs = CRs.Load(sn!SampleID & "", gDONTCARE, "Results")
320       If Not CRs Is Nothing Then
330           For Each CR In CRs
340               If CR.Code = "044" Then
350                   If Not FirstDayFilled Then
360                       FirstDayFilled = True
370                       gData(365, 1) = Format(sn!Rundate, "dd/mmm/yyyy")
380                       gData(365, 2) = Val(CR.Result)
390                       If Not tb.EOF Then
400                           gData(365, 3) = Val(tb!Warfarin & "")
410                           CurrentDose = tb!Warfarin & ""
420                       Else
430                           CurrentDose = ""
440                       End If
450                       RP.SampleID = sn!SampleID & ""
460                       LatestDate = Format(sn!Rundate, "dd/mmm/yyyy")
470                       LatestINR = CR.Result
480                       pLatest = Format(LatestDate, "dd/mm/yyyy")
490                   Else
500                       NumberOfDays = Abs(DateDiff("D", LatestDate, Format(sn!Rundate, "dd/mmm/yyyy")))
510                       If NumberOfDays < 365 Then
520                           gData(365 - NumberOfDays, 1) = Format(sn!Rundate, "dd/mmm/yyyy")
530                           gData(365 - NumberOfDays, 2) = CR.Result
540                           If Not tb.EOF Then
550                               gData(365 - NumberOfDays, 3) = Val(tb!Warfarin & "")
560                           End If
570                           EarliestDate = Format(sn!Rundate, "dd/mmm/yyyy")
580                           pEarliest = Format(sn!Rundate, "dd/mm/yyyy")
590                       Else
600                           Exit Do
610                       End If
620                   End If
630                   Counter = Counter + 1
640                   If Counter = 15 Then
650                       Exit Do
660                   End If
670               End If
680           Next
690       End If
700       sn.MoveNext
710   Loop
720   If EarliestDate = "" Or LatestDate = "" Then Exit Sub
730   NumberOfDays = Abs(DateDiff("d", EarliestDate, LatestDate))
740   PixelsPerDay = (pb.Width - 1060) / NumberOfDays
750   PixelsPerPointY = pb.Height / 6
760   X = pb.Width - 580
770   Y = pb.Height - (gData(365, 2) * PixelsPerPointY)
780   pb.ForeColor = vbBlue
790   pb.Circle (X, Y), 30
800   pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
810   pb.PSet (X, Y)
820   For n = 364 To 1 Step -1
830       If gData(n, 1) <> 0 Then
840           NumberOfDays = Abs(DateDiff("d", EarliestDate, gData(n, 1)))
850           X = 580 + (NumberOfDays * PixelsPerDay)
860           Y = pb.Height - (gData(n, 2) * PixelsPerPointY)
870           pb.Line -(X, Y)
880           pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
890           pb.Circle (X, Y), 30
900           pb.PSet (X, Y)
910       End If
920   Next
      'Draw Warfarin
930   NumberOfDays = Abs(DateDiff("d", EarliestDate, LatestDate))
940   PixelsPerPointY = pb.Height / 12
950   X = pb.Width - 580
960   Y = pb.Height - (gData(365, 3) * PixelsPerPointY)
970   pb.ForeColor = vbGreen
980   pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbGreen, BF
990   pb.PSet (X, Y)
1000  For n = 364 To 1 Step -1
1010      If gData(n, 1) <> 0 Then
1020          NumberOfDays = Abs(DateDiff("d", EarliestDate, gData(n, 1)))
1030          X = 480 + (NumberOfDays * PixelsPerDay)
1040          Y = pb.Height - (gData(n, 3) * PixelsPerPointY)
1050          pb.Line -(X, Y)
1060          pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbGreen, BF
1070          pb.PSet (X, Y)
1080      End If
1090  Next
1100  sql = "select * from INRHistory where " & _
            "chart = '" & Chart & "'"
1110  Set tb = New Recordset
1120  RecOpenClient 0, tb, sql
1130  pCondition = ""
1140  pLowerTarget = ""
1150  pUpperTarget = ""
1160  If Not tb.EOF Then
1170      PixelsPerPointY = pb.Height / 6
1180      pb.ForeColor = vbRed
1190      Y = pb.Height - (Val(tb!UpperTarget & "") * PixelsPerPointY)
1200      pb.Line (480, Y)-(pb.Width - 580, Y), vbRed
1210      Y = pb.Height - (Val(tb!LowerTarget & "") * PixelsPerPointY)
1220      pb.Line (480, Y)-(pb.Width - 580, Y), vbRed
1230      pLowerTarget = tb!LowerTarget & ""
1240      pUpperTarget = tb!UpperTarget & ""
1250      pCondition = tb!Condition & ""
1260  End If
1270  pb.ForeColor = vbBlack
1280  pb.Font.Bold = True
1290  pb.CurrentX = pb.TextWidth("I") + 380
1300  pb.CurrentY = pb.Height - 1.5 * pb.TextHeight("W")
1310  pb.Print pEarliest
1320  pb.CurrentX = pb.Width - pb.TextWidth("ww/ww/www") - 480
1330  pb.CurrentY = pb.Height - 1.5 * pb.TextHeight("W")
1340  pb.Print pLatest

1350  Exit Sub

DrawPicture_Error:

      Dim strES As String
      Dim intEL As Integer

1360  intEL = Erl
1370  strES = Err.Description
1380  LogError "frmMain", "DrawPicture", intEL, strES, sql

End Sub

Private Sub RefreshInstalledPrinters()

          Dim Px As Printer
          Dim sql As String

10        On Error GoTo RefreshInstalledPrinters_Error
20        sql = "Delete from InstalledPrinters"
30        Cnxn(0).Execute sql
40        For Each Px In Printers
50            sql = "Insert into InstalledPrinters " & _
                    "(PrinterName) VALUES " & _
                    "('" & Px.DeviceName & "')"
60            Cnxn(0).Execute sql
70        Next
80        Exit Sub

RefreshInstalledPrinters_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmMain", "RefreshInstalledPrinters", intEL, strES, sql


End Sub

Private Sub cmdRefreshPrinters_Click()

10    RefreshInstalledPrinters

End Sub

Private Sub Form_Activate()

          Dim Path As String
          Dim strVersion As String

10        If Not IsIDE Then
20            Path = CheckNewEXE("PrintHandler")
30            If Path <> "" Then
40                Shell App.Path & "\CustomStart.exe PrintHandler"
50                End
60                Exit Sub
70            End If
80        End If
90        strVersion = App.Major & "." & App.Minor & "." & App.Revision
100       Me.Caption = "NetAcquire Print Handler. Version " & strVersion
110       If EnableTimer Then
120           Timer1.Enabled = True
130       End If

End Sub



Private Sub Form_Load()

      Dim strUseSecondPage As String

10    On Error GoTo Form_Load_Error

20    If App.PrevInstance Then End
30    CheckIDE
40    GetINI    ' ConnectToDatabase
50    EnsureColumnExists "CoagTestDefinitions", "PrintRefRange", "tinyint NOT NULL DEFAULT 1"
60    EnsureColumnExists "PrintPending", "ThisIsCopy", "tinyint"
70    EnsureColumnExists "BioTestDefinitions", "PrintSplit", "int DEFAULT 0"
80    EnsureColumnExists "ImmTestDefinitions", "PrintSplit", "int DEFAULT 0"
90    EnsureColumnExists "GPs", "PrintReport", "bit NOT NULL DEFAULT 1"
100   RefreshInstalledPrinters
110   LoadOptions
120   Loading = True
130   txtMoreThan = GetSetting("NetAcquire", "PrintOptions", "IfMoreThan", "18")
140   strUseSecondPage = GetSetting("NetAcquire", "PrintOptions", "UseSecondPage", "True")
150   If strUseSecondPage = "True" Then
160       optSecondPage = True
170   Else
180       optSideBySide = True
190   End If
200   sAppName = GetSetting("NetAcquire", "CavMonFAX", "AppName", "Workspace - Lotus Notes")
210   sAppPath = GetSetting("NetAcquire", "CavMonFAX", "AppPath", "c:\lotus\notes\nlnotes.exe")
220   lblZSubmit = GetOptionSetting("ZSubmitFolder", "")
230   lblDocument = GetOptionSetting("ZSubmitDocument", "")
240   Loading = False
250   EnableTimer = True

260   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmMain", "Form_Load", intEL, strES

End Sub

Private Sub lblDocument_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim Path As String

10        On Error GoTo lblDocument_MouseUp_Error
20        EnableTimer = False
30        Timer1.Enabled = False
40        If UCase$(iBOX("Password", , , True)) = "TEMO" Then
50            Path = Trim$(iBOX("Document Folder PathName?", , lblDocument))
60            If Len(Path) > 1 Then
70                If Right$(Path, 1) = "\" Then
80                    Path = Left$(Path, Len(Path) - 1)
90                End If
100           End If
110           lblDocument = Path
120           SaveOptionSetting "ZSubmitDocument", lblDocument
130       End If

140       EnableTimer = True
150       Timer1.Enabled = True
160       Exit Sub

lblDocument_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmMain", "lblDocument_MouseUp", intEL, strES

End Sub


Private Sub lblZSubmit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim Path As String

10    On Error GoTo lblZSubmit_MouseUp_Error
20    EnableTimer = False
30    Timer1.Enabled = False
40    If UCase$(iBOX("Password", , , True)) = "TEMO" Then
50        Path = Trim$(iBOX("ZSubmit Folder PathName?", , lblZSubmit))
60        If Len(Path) > 1 Then
70            If Right$(Path, 1) = "\" Then
80                Path = Left$(Path, Len(Path) - 1)
90            End If
100       End If
110       lblZSubmit = Path
120       SaveOptionSetting "ZSubmitFolder", lblZSubmit
130   End If
140   EnableTimer = True
150   Timer1.Enabled = True
160   Exit Sub

lblZSubmit_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmMain", "lblZSubmit_MouseUp", intEL, strES

End Sub


Private Sub mExit_Click()

10        Unload Me

End Sub

Private Sub mHaemNormal_Click()

10        fHaemNoSexNormal.Show 1

End Sub

Private Sub mnuAppName_Click()

10        sAppName = GetSetting("NetAcquire", "CavMonFAX", "AppName", "Workspace - Lotus Notes")
20        sAppName = iBOX("Enter App Name for FAX/MAPI", , sAppName)
30        SaveSetting "NetAcquire", "CavMonFAX", "AppName", sAppName

End Sub

Private Sub mnuAppPath_Click()

10        sAppPath = GetSetting("NetAcquire", "CavMonFAX", "AppPath", "c:\lotus\notes\nlnotes.exe")
20        sAppPath = iBOX("Enter App Path for FAX/MAPI", , sAppPath)
30        SaveSetting "NetAcquire", "CavMonFAX", "AppPath", sAppPath

End Sub


Private Sub mnuPrintSplit_Click()

10        frmPrintSplit.Show 1

End Sub

Private Sub mPrinters_Click()

10        fPrinters.Show 1

End Sub

Private Sub optSecondPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        SaveSetting "NetAcquire", "PrintOptions", "UseSecondPage", "True"

End Sub


Private Sub optSideBySide_Click()

10        SaveSetting "NetAcquire", "PrintOptions", "UseSecondPage", "False"

End Sub


Private Sub Timer1_Timer()

      Static pCounter As Integer
      Static nCounter As Integer
      Dim Path As String

10    On Error GoTo Timer1_Timer_Error

20    pbar = pbar + 10
30    If pbar = pbar.Max Then
40        pCounter = pCounter + 1
50        If pCounter >= 3 Then    '30 seconds
60            pCounter = 0
              '60                RefreshInstalledPrinters
70        End If
80        nCounter = nCounter + 1
90        If nCounter >= 180 Then    '30 minutes
100           nCounter = 0
110           If Not IsIDE Then
120               Path = CheckNewEXE("PrintHandler")
130               If Path <> "" Then
140                   Shell App.Path & "\CustomStart.exe PrintHandler"
150                   End
160                   Exit Sub
170               End If
180           End If
190       End If
200       ProcessPrintQueue
          'Zyam 2-5-24
          PrintRemMicroRep
          
          'Zyam 2-5-24
210       pbar = 0

220   End If

230   Exit Sub

Timer1_Timer_Error:

       Dim strES As String
       Dim intEL As Integer

240    intEL = Erl
250    strES = Err.Description
260    LogError "frmMain", "Timer1_Timer", intEL, strES

End Sub
'Zyam 15-06-24
Private Sub PrintRemMicroRep()

          Dim sqlPVL As String
          Dim tbPVL As Recordset

          Dim sqlPrintPend As String
          Dim tbPrintPend As Recordset

          Dim sqlDemo As String
          Dim tbDemo As Recordset

10        On Error GoTo PrintRemMicroRep_Error

20        sqlPVL = "SELECT PrintedBy, SampleID FROM PrintValidLog WHERE Printed = 0 AND Valid = 1 AND (PrintedBy <> '' Or PrintedBy IS NOT NULL) AND ValidatedDateTime BETWEEN DATEADD(m, -3, GETDATE()) AND GETDATE()"
          
30        Set tbPVL = New Recordset
40        RecOpenClient 0, tbPVL, sqlPVL

50        If Not tbPVL Is Nothing Then
60            Do While Not tbPVL.EOF
70                sqlPrintPend = "SELECT * FROM PrintPending WHERE SampleID = '" & Trim(tbPVL!SampleID) & "'"
80                Set tbPrintPend = New Recordset
90                RecOpenClient 0, tbPrintPend, sqlPrintPend
100               If Not tbPrintPend Is Nothing Then
110                   If tbPrintPend.EOF Then
120                       tbPrintPend.AddNew
130                       sqlDemo = "SELECT Ward, Clinician, GP, ForMicro, ForBio, ForCoag, ForHaem FROM Demographics WHERE SampleID = '" & Trim(tbPVL!SampleID) & "'"
140                       Set tbDemo = New Recordset
150                       RecOpenClient 0, tbDemo, sqlDemo
160                       If Not tbDemo Is Nothing Then
170                           If Not tbDemo.EOF Then
180                               tbPrintPend!SampleID = Trim(tbPVL!SampleID)
190                               If Not IsNull(tbDemo!ForMicro) Then
200                                   tbPrintPend!Department = "M"
210                               ElseIf Not IsNull(tbDemo!ForBio) Then
220                                   tbPrintPend!Department = "B"
230                               ElseIf Not IsNull(tbDemo!ForCoag) Then
240                                   tbPrintPend!Department = "C"
250                               ElseIf Not IsNull(tbDemo!ForHaem) Then
260                                   tbPrintPend!Department = "H"
270                               End If
280                               tbPrintPend!Initiator = tbPVL!PrintedBy
290                               tbPrintPend!Clinician = tbDemo!Clinician
300                               tbPrintPend!Ward = tbDemo!Ward
310                               tbPrintPend!GP = tbDemo!GP
320                               tbPrintPend.Update
330                           End If
340                       End If
                  
350                   End If
360               End If
370           tbPVL.MoveNext
380           Loop

390       End If
400       Exit Sub
PrintRemMicroRep_Error:

410   LogError "frmMain", "PrintRemMicroRep", Erl, Err.Description

End Sub
'Zyam 15-06-24

Private Sub txtMoreThan_Change()

10        If Not Loading Then
20            If Val(txtMoreThan) > 0 And Val(txtMoreThan) < 30 Then
30                SaveSetting "NetAcquire", "PrintOptions", "IfMoreThan", txtMoreThan
40            Else
50                txtMoreThan = GetSetting("NetAcquire", "PrintOptions", "IfMoreThan", "18")
60            End If
70        End If

End Sub


Private Sub FillReport(ByVal strCounter As String)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillReport_Error
20        rtb = ""
30        rtb.SelText = ""
40        sql = "SELECT Report FROM Reports WHERE " & _
                "Counter = '" & strCounter & "' "
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80          If Trim(tb!Report & "") <> "" Then
90            rtb.SelText = Trim(tb!Report)
100           rtb.TextRTF = Replace(rtb.TextRTF, String$(200, "-"), _
              "- THIS IS A COPY REPORT - NOT FOR FILING - THIS IS A COPY REPORT -- THIS IS A COPY REPORT - NOT FOR FILING - THIS IS A COPY REPORT -")
110         End If
120       End If
130       Exit Sub

FillReport_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmMain", "FillReport", intEL, strES, sql

End Sub

