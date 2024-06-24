VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWardList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Ward List"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGPPrinting 
      Caption         =   "DR. Name"
      Height          =   1485
      Left            =   9495
      TabIndex        =   24
      Top             =   7065
      Visible         =   0   'False
      Width           =   3585
      Begin VB.CommandButton cmdCancelGpPrinting 
         Appearance      =   0  'Flat
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2565
         TabIndex        =   27
         Top             =   720
         Width           =   930
      End
      Begin VB.CommandButton cmdUpdateGpPrinting 
         Appearance      =   0  'Flat
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2565
         TabIndex        =   26
         Top             =   270
         Width           =   930
      End
      Begin VB.ListBox LstGpPrinting 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         ItemData        =   "frmWardList.frx":0000
         Left            =   135
         List            =   "frmWardList.frx":0010
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   270
         Width           =   2370
      End
   End
   Begin VB.ComboBox cmbListItems 
      Height          =   315
      ItemData        =   "frmWardList.frx":004A
      Left            =   8520
      List            =   "frmWardList.frx":004C
      TabIndex        =   22
      Top             =   720
      Width           =   2115
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      Height          =   705
      Left            =   12105
      Picture         =   "frmWardList.frx":004E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   45
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   13215
      Picture         =   "frmWardList.frx":0918
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   60
      Width           =   975
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8910
      Top             =   7740
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8460
      Top             =   7740
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   7290
      Visible         =   0   'False
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   705
      Left            =   13200
      Picture         =   "frmWardList.frx":0C22
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5145
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   525
      Left            =   13185
      Picture         =   "frmWardList.frx":1064
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4050
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   13185
      Picture         =   "frmWardList.frx":14A6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3465
      Width           =   465
   End
   Begin VB.ListBox lHospital 
      Height          =   1230
      Left            =   6270
      TabIndex        =   9
      Top             =   150
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add New Ward"
      Height          =   1365
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   5925
      Begin VB.TextBox txtPrinter 
         Height          =   285
         Left            =   810
         TabIndex        =   16
         Top             =   960
         Width           =   4755
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   11
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   420
         Width           =   705
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   810
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3645
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   810
         MaxLength       =   12
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   330
         TabIndex        =   17
         Top             =   990
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   1860
         TabIndex        =   10
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   270
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdWardList 
      Height          =   5445
      Left            =   150
      TabIndex        =   2
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   9
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
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmWardList.frx":18E8
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "&Exit"
      Height          =   705
      Left            =   13230
      Picture         =   "frmWardList.frx":19CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6270
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   945
      Left            =   13275
      MaskColor       =   &H8000000F&
      Picture         =   "frmWardList.frx":2038
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1575
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Items to be displayed in list"
      Height          =   195
      Left            =   8520
      TabIndex        =   23
      Top             =   480
      Width           =   1875
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Click on Code to Edit/Remove record."
      Height          =   195
      Left            =   150
      TabIndex        =   21
      Top             =   7050
      Width           =   3255
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13170
      TabIndex        =   19
      Top             =   810
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "frmWardList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer






Private Sub cmbListItems_Click()
50450 cmdSave.Visible = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
50460 KeyAscii = 0
End Sub




Private Sub cmdDelete_Click()
      'check if record doesn't exist in demographics
      Dim tb As Recordset
      Dim sql As String

50470 On Error GoTo cmdDelete_Click_Error

50480 If grdWardList.row = 0 Or grdWardList.Rows <= 2 Then Exit Sub

50490 sql = "SELECT Count(*) as RC FROM Demographics WHERE Ward = '" & grdWardList.TextMatrix(grdWardList.row, 2) & "'"
50500 Set tb = New Recordset
50510 RecOpenClient 0, tb, sql
50520 If tb!rc > 0 Then
50530     iMsg "Reference to " & grdWardList.TextMatrix(grdWardList.row, 2) & " is in use so cannot be deleted"
50540     Exit Sub
50550 Else
50560     If iMsg("Are you sure you want to delete " & grdWardList.TextMatrix(grdWardList.row, 2) & "?", vbYesNo) = vbYes Then
50570         Cnxn(0).Execute "DELETE FROM Wards WHERE " & _
                              "Code = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "' AND Text = '" & grdWardList.TextMatrix(grdWardList.row, 2) & "'"
50580         FillG
50590     End If

50600 End If

50610 Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50620 intEL = Erl
50630 strES = Err.Description
50640 LogError "fWardList", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

50650 With grdWardList
50660     If .row = .Rows - 1 Then Exit Sub
50670     n = .row

50680     FireCounter = FireCounter + 1
50690     If FireCounter > 5 Then
50700         tmrDown.Interval = 100
50710     End If

50720     VisibleRows = .height \ .RowHeight(1) - 1

50730     .Visible = False

50740     s = ""
50750     For X = 0 To .Cols - 1
50760         s = s & .TextMatrix(n, X) & vbTab
50770     Next
50780     s = Left$(s, Len(s) - 1)

50790     .RemoveItem n
50800     If n < .Rows Then
50810         .AddItem s, n + 1
50820         .row = n + 1
50830     Else
50840         .AddItem s
50850         .row = .Rows - 1
50860     End If

50870     For X = 0 To .Cols - 1
50880         .Col = X
50890         .CellBackColor = vbYellow
50900     Next

50910     If Not .RowIsVisible(.row) Or .row = .Rows - 1 Then
50920         If .row - VisibleRows + 1 > 0 Then
50930             .TopRow = .row - VisibleRows + 1
50940         End If
50950     End If

50960     .Visible = True
50970 End With

50980 cmdSave.Visible = True

End Sub

Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

50990 With grdWardList
51000     If .row = 1 Then Exit Sub

51010     FireCounter = FireCounter + 1
51020     If FireCounter > 5 Then
51030         tmrUp.Interval = 100
51040     End If

51050     n = .row

51060     .Visible = False

51070     s = ""
51080     For X = 0 To .Cols - 1
51090         s = s & .TextMatrix(n, X) & vbTab
51100     Next
51110     s = Left$(s, Len(s) - 1)

51120     .RemoveItem n
51130     .AddItem s, n - 1

51140     .row = n - 1
51150     For X = 0 To .Cols - 1
51160         .Col = X
51170         .CellBackColor = vbYellow
51180     Next

51190     If Not .RowIsVisible(.row) Then
51200         .TopRow = .row
51210     End If

51220     .Visible = True

51230     cmdSave.Visible = True
51240 End With

End Sub



Private Sub cmdExit_Click()

51250 Unload Me

End Sub


Private Sub cmdCancelGpPrinting_Click()
51260 On Error GoTo cmdCancelGpPrinting_Click_Error

51270 fraGPPrinting.Visible = False

51280 Exit Sub

cmdCancelGpPrinting_Click_Error:
      Dim strES As String
      Dim intEL As Integer

51290 intEL = Erl
51300 strES = Err.Description
51310 LogError "frmWardList", "cmdCancelGpPrinting_Click", intEL, strES

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

51320 FireDown

51330 tmrDown.Interval = 250
51340 FireCounter = 0

51350 tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

51360 tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

51370 FireUp

51380 tmrUp.Interval = 250
51390 FireCounter = 0

51400 tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

51410 tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()


      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

51420 Screen.MousePointer = 11

51430 OriginalPrinter = Printer.DeviceName

51440 If Not SetFormPrinter() Then Exit Sub

51450 Printer.FontName = "Courier New"
51460 Printer.Orientation = vbPRORPortrait


      '****Report heading
51470 Printer.FontSize = 10
51480 Printer.Font.Bold = True
51490 Printer.Print
51500 Printer.Print FormatString("List Of Wards (" & lHospital & ")", 99, , AlignCenter)

      '****Report body heading

51510 Printer.Font.size = 9
51520 For i = 1 To 108
51530     Printer.Print "-";
51540 Next i
51550 Printer.Print


51560 Printer.Print FormatString("", 0, "|");
51570 Printer.Print FormatString("In Use", 6, "|", AlignCenter);
51580 Printer.Print FormatString("Code", 10, "|", AlignCenter);
51590 Printer.Print FormatString("Description", 88, "|", AlignCenter)
      '****Report body

51600 Printer.Font.Bold = False

51610 For i = 1 To 108
51620     Printer.Print "-";
51630 Next i
51640 Printer.Print
51650 For Y = 1 To grdWardList.Rows - 1
51660     Printer.Print FormatString("", 0, "|");
51670     Printer.Print FormatString(grdWardList.TextMatrix(Y, 0), 6, "|", AlignLeft);
51680     Printer.Print FormatString(grdWardList.TextMatrix(Y, 1), 10, "|", AlignLeft);
51690     Printer.Print FormatString(grdWardList.TextMatrix(Y, 2), 88, "|", AlignLeft)


51700 Next



51710 Printer.EndDoc

51720 Screen.MousePointer = 0

51730 For Each Px In Printers
51740     If Px.DeviceName = OriginalPrinter Then
51750         Set Printer = Px
51760         Exit For
51770     End If
51780 Next
End Sub

Private Sub cmdAdd_Click()

51790 txtCode = UCase$(Trim$(txtCode))
51800 If txtCode = "" Then
51810     iMsg "Enter Code.", vbCritical
51820     Exit Sub
51830 End If

51840 txtText = Trim$(txtText)
51850 If txtText = "" Then
51860     iMsg "Enter Ward.", vbCritical
51870     Exit Sub
51880 End If

51890 grdWardList.AddItem "Yes" & vbTab & _
                          txtCode & vbTab & _
                          txtText & vbTab & _
                          txtFAX & vbTab & _
                          txtPrinter

51900 txtCode = ""
51910 txtText = ""
51920 txtFAX = ""
51930 txtPrinter = ""

51940 cmdSave.Visible = True

End Sub

Private Sub cmdSave_Click()

      Dim Hosp As String
      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String

51950 On Error GoTo cmdSave_Click_Error

51960 sql = "Select * from Lists where " & _
            "ListType = 'HO' " & _
            "and Text = '" & lHospital & "' and InUse = 1"
51970 Set tb = New Recordset
51980 RecOpenServer 0, tb, sql
51990 If Not tb.EOF Then
52000     Hosp = tb!Code & ""
52010 End If

52020 pb.max = grdWardList.Rows - 1
52030 pb.Visible = True
52040 cmdSave.Caption = "Saving..."

52050 For Y = 1 To grdWardList.Rows - 1
52060     pb = Y
52070     sql = "Select * from Wards where " & _
                "Code = '" & grdWardList.TextMatrix(Y, 1) & "' " & _
                "and HospitalCode = '" & Hosp & "'"
52080     Set tb = New Recordset
52090     RecOpenServer 0, tb, sql
52100     If tb.EOF Then tb.AddNew
52110     With tb
52120         !Code = grdWardList.TextMatrix(Y, 1)
52130         !HospitalCode = Hosp
52140         !InUse = grdWardList.TextMatrix(Y, 0) = "Yes"
52150         !Text = grdWardList.TextMatrix(Y, 2)
52160         !FAX = grdWardList.TextMatrix(Y, 3)
52170         !PrinterAddress = grdWardList.TextMatrix(Y, 4)
52180         !ListOrder = Y
52190         !Location = grdWardList.TextMatrix(Y, 5)
52200         .Update

52210         sql = "IF EXISTS (SELECT * FROM IncludeEGFR " & _
                    "           WHERE SourceType = 'Ward' " & _
                    "           AND Hospital = '" & lHospital & "' " & _
                    "           AND SourceName = '" & AddTicks(grdWardList.TextMatrix(Y, 2)) & "' ) " & _
                    "    UPDATE IncludeEGFR " & _
                    "    SET Include = '" & IIf(grdWardList.TextMatrix(Y, 6) = "Yes", 1, 0) & "' " & _
                    "    WHERE SourceType = 'Ward' " & _
                    "    AND SourceName = '" & AddTicks(grdWardList.TextMatrix(Y, 2)) & "' " & _
                    "ELSE " & _
                    "    INSERT INTO IncludeEGFR (SourceType, Hospital, SourceName, Include) " & _
                    "    VALUES ('Ward', " & _
                    "            '" & lHospital & "', " & _
                    "            '" & AddTicks(grdWardList.TextMatrix(Y, 2)) & "', " & _
                    "            '" & IIf(grdWardList.TextMatrix(Y, 6) = "Yes", 1, 0) & "')"
52220         Cnxn(0).Execute sql

52230         sql = "IF EXISTS (SELECT * FROM IncludeAutoValUrine " & _
                    "           WHERE SourceType = 'Ward' " & _
                    "           AND Hospital = '" & lHospital & "' " & _
                    "           AND SourceName = '" & AddTicks(grdWardList.TextMatrix(Y, 2)) & "' ) " & _
                    "    UPDATE IncludeAutoValUrine " & _
                    "    SET Include = '" & IIf(grdWardList.TextMatrix(Y, 7) = "Yes", 1, 0) & "' " & _
                    "    WHERE SourceType = 'Ward' " & _
                    "    AND SourceName = '" & AddTicks(grdWardList.TextMatrix(Y, 2)) & "' " & _
                    "ELSE " & _
                    "    INSERT INTO IncludeAutoValUrine (SourceType, Hospital, SourceName, Include) " & _
                    "    VALUES ('Ward', " & _
                    "            '" & lHospital & "', " & _
                    "            '" & AddTicks(grdWardList.TextMatrix(Y, 2)) & "', " & _
                    "            '" & IIf(grdWardList.TextMatrix(Y, 7) = "Yes", 1, 0) & "')"
52240         Cnxn(0).Execute sql

52250     End With
52260 Next

52270 Call SaveOptionSetting("WardListLength", cmbListItems)

52280 pb.Visible = False
52290 cmdSave.Visible = False
52300 cmdSave.Caption = "Save"

52310 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

52320 intEL = Erl
52330 strES = Err.Description
52340 LogError "fWardList", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub cmdUpdateGpPrinting_Click()
      Dim Y As Long
      Dim sql As String
      Dim tb As Recordset
      Dim Dept As String

52350 On Error GoTo cmdUpdateGpPrinting_Click_Error

52360 For Y = 0 To LstGpPrinting.ListCount - 1
52370     Dept = LstGpPrinting.List(Y)

52380     If LstGpPrinting.Selected(Y) = True Then

52390         sql = "IF EXISTS (SELECT * FROM DisablePrinting WHERE " & _
                    "           Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "           AND GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "' )" & _
                    "    UPDATE DisablePrinting " & _
                    "    SET Department = '" & LstGpPrinting.List(Y) & "', " & _
                    "    GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "', " & _
                    "    GPName = '" & grdWardList.TextMatrix(grdWardList.row, 2) & "', " & _
                    "    Type = 'WARD', " & _
                    "    Disable = '1' " & _
                    "    WHERE Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "    AND GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "' " & _
                    "ELSE " & _
                    "    INSERT INTO DisablePrinting " & _
                    "    (GPCode, GPName, Type, Disable, Department) "
52400         sql = sql & _
                    "    VALUES ( " & _
                    "    '" & grdWardList.TextMatrix(grdWardList.row, 1) & "', " & _
                    "    '" & grdWardList.TextMatrix(grdWardList.row, 2) & "', " & _
                    "    'WARD', " & _
                    "    '1', " & _
                    "    '" & LstGpPrinting.List(Y) & "'" & _
                    "     )"
52410         Cnxn(0).Execute sql
52420     Else
52430         sql = "DELETE FROM DisablePrinting WHERE " & _
                    "Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "AND GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "'"
52440         Cnxn(0).Execute sql
52450     End If
52460 Next

52470 fraGPPrinting.Visible = False
      'MsgBox "Updated!", vbOKOnly, "Disable Printing"

52480 Exit Sub

cmdUpdateGpPrinting_Click_Error:
      Dim strES As String
      Dim intEL As Integer

52490 intEL = Erl
52500 strES = Err.Description
52510 LogError "frmWardList", "cmdUpdateGpPrinting_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

52520 ExportFlexGrid grdWardList, Me

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

52530 On Error GoTo Form_Load_Error

52540 lHospital.Clear

52550 sql = "Select * from Lists where " & _
            "ListType = 'HO' and InUse = 1 " & _
            "order by ListOrder"
52560 Set tb = New Recordset
52570 RecOpenServer 0, tb, sql
52580 Do While Not tb.EOF
52590     lHospital.AddItem tb!Text & ""
52600     tb.MoveNext
52610 Loop
52620 If lHospital.ListCount > 0 Then
52630     lHospital.ListIndex = 0
52640 End If

52650 FillG

      Dim i As Integer
52660 cmbListItems.Clear
52670 For i = 8 To 32 Step 8
52680     cmbListItems.AddItem i
52690 Next i
52700 cmbListItems.Text = GetOptionSetting("WardListLength", 8)

52710 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

52720 intEL = Erl
52730 strES = Err.Description
52740 LogError "fWardList", "Form_Load", intEL, strES, sql


End Sub

Private Sub FillG()

      Dim s As String
      Dim sql As String
      Dim tb As Recordset

52750 On Error GoTo FillG_Error

52760 grdWardList.Rows = 2
52770 grdWardList.AddItem ""
52780 grdWardList.RemoveItem 1

52790 sql = "SELECT W.*, COALESCE(E.Include, 0) EGFR, Coalesce(U.Include, 0) UAV " & _
            "FROM Wards W JOIN Lists L ON W.HospitalCode = L.Code " & _
            "LEFT JOIN IncludeEGFR E ON W.Text = E.SourceName " & _
            "LEFT JOIN IncludeAutoValUrine U ON W.Text = U.SourceName " & _
            "WHERE L.Text = '" & lHospital & "' " & _
            "AND L.ListType = 'HO'"
52800 Set tb = New Recordset
52810 RecOpenServer 0, tb, sql

52820 Do While Not tb.EOF
52830     With tb
52840         s = IIf(!InUse, "Yes", "No") & vbTab & _
                  !Code & vbTab & _
                  !Text & vbTab & _
                  !FAX & vbTab & _
                  !PrinterAddress & vbTab & _
                  !Location & vbTab & _
                  IIf(!EGFR, "Yes", "No") & vbTab & _
                  IIf(!UAV, "Yes", "No") & vbTab & _
                  "Click To View"

52850         grdWardList.AddItem s
52860     End With
52870     tb.MoveNext
52880 Loop

52890 If grdWardList.Rows > 2 Then
52900     grdWardList.RemoveItem 1
52910 End If

52920 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

52930 intEL = Erl
52940 strES = Err.Description
52950 LogError "fWardList", "FillG", intEL, strES, sql

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

52960 If cmdSave.Visible Then
52970     If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
52980         Cancel = True
52990     End If
53000 End If

End Sub

Private Sub grdWardList_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim tb As Recordset
      Dim sql As String


53010 On Error GoTo grdWardList_Click_Error

53020 ySave = grdWardList.row

53030 If grdWardList.MouseRow = 0 Then
53040     If SortOrder Then
53050         grdWardList.Sort = flexSortGenericAscending
53060     Else
53070         grdWardList.Sort = flexSortGenericDescending
53080     End If
53090     SortOrder = Not SortOrder
53100     cmdMoveUp.Enabled = False
53110     cmdMoveDown.Enabled = False
53120     cmdSave.Visible = True
53130     Exit Sub
53140 End If


53150 If grdWardList.Col = 0 Or grdWardList.Col = 6 Or grdWardList.Col = 7 Then
53160     grdWardList = IIf(grdWardList = "No", "Yes", "No")
53170     cmdSave.Visible = True
53180     Exit Sub
53190 End If

53200 If grdWardList.Col = 1 Then
53210     sql = "SELECT Count(*) as RC FROM Demographics WHERE Ward = '" & grdWardList.TextMatrix(grdWardList.row, 2) & "'"
53220     Set tb = New Recordset
53230     RecOpenClient 0, tb, sql
53240     If tb!rc > 0 Then
53250         iMsg "Reference to " & grdWardList.TextMatrix(grdWardList.row, 2) & " is in use so cannot be deleted"
53260         Exit Sub
53270     End If

53280     grdWardList.Enabled = False
53290     If iMsg("Edit this line?", vbQuestion + vbYesNo) = vbYes Then
53300         txtCode = grdWardList.TextMatrix(grdWardList.row, 1)
53310         txtText = grdWardList.TextMatrix(grdWardList.row, 2)
53320         txtFAX = grdWardList.TextMatrix(grdWardList.row, 3)
53330         txtPrinter = grdWardList.TextMatrix(grdWardList.row, 4)
53340         grdWardList.RemoveItem grdWardList.row
53350         cmdSave.Visible = True
53360     End If
53370     grdWardList.Enabled = True
53380     Exit Sub
53390 End If

53400 If grdWardList.Col = 5 Then
53410     Select Case grdWardList.TextMatrix(grdWardList.row, 5)
          Case "": grdWardList.TextMatrix(grdWardList.row, 5) = "In-House"
53420     Case "In-House": grdWardList.TextMatrix(grdWardList.row, 5) = "External"
53430     Case "External": grdWardList.TextMatrix(grdWardList.row, 5) = ""
53440     Case "Else": grdWardList.TextMatrix(grdWardList.row, 5) = ""
53450     End Select
53460     cmdSave.Visible = True
53470     Exit Sub
53480 End If
      '----------farhan---------
53490 If grdWardList.Col = 8 Then
53500     fraGPPrinting.Visible = True
53510     fraGPPrinting.Caption = grdWardList.TextMatrix(grdWardList.row, 2)
53520     FillLstGpPrinting
53530 End If
      '==============================
53540 grdWardList.Visible = False
53550 grdWardList.Col = 0
53560 For Y = 1 To grdWardList.Rows - 1
53570     grdWardList.row = Y
53580     If grdWardList.CellBackColor = vbYellow Then
53590         For X = 0 To grdWardList.Cols - 1
53600             grdWardList.Col = X
53610             grdWardList.CellBackColor = 0
53620         Next
53630         Exit For
53640     End If
53650 Next
53660 grdWardList.row = ySave
53670 grdWardList.Visible = True

53680 For X = 0 To grdWardList.Cols - 1
53690     grdWardList.Col = X
53700     grdWardList.CellBackColor = vbYellow
53710 Next

53720 cmdMoveUp.Enabled = True
53730 cmdMoveDown.Enabled = True

53740 Exit Sub

grdWardList_Click_Error:

      Dim strES As String
      Dim intEL As Integer

53750 intEL = Erl
53760 strES = Err.Description
53770 LogError "fWardList", "grdWardList_Click", intEL, strES, sql


End Sub
Private Sub FillLstGpPrinting()

      Dim tb As Recordset
      Dim sql As String
      'Dim dept As String
      Dim Y As Integer

53780 On Error GoTo FillLstGpPrinting_Error

53790 sql = "SELECT * FROM DisablePrinting WHERE " & _
            " GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "'"
53800 Set tb = New Recordset
53810 RecOpenServer 0, tb, sql

53820 For Y = 0 To LstGpPrinting.ListCount - 1
          'dept = LstGpPrinting.List(Y)
53830     LstGpPrinting.Selected(Y) = False
53840 Next Y

53850 Do While Not tb.EOF
53860     For Y = 0 To LstGpPrinting.ListCount - 1
53870         If LstGpPrinting.List(Y) = tb!Department And tb!Disable = True Then
53880             LstGpPrinting.Selected(Y) = True
53890         End If
53900     Next
53910     tb.MoveNext
53920 Loop


53930 Exit Sub

FillLstGpPrinting_Error:
      Dim strES As String
      Dim intEL As Integer

53940 intEL = Erl
53950 strES = Err.Description
53960 LogError "frmWardList", "FillLstGpPrinting", intEL, strES, sql

End Sub

Private Sub grdWardList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

53970 If grdWardList.MouseRow = 0 Then
53980     grdWardList.ToolTipText = ""
53990 ElseIf grdWardList.MouseCol = 0 Then
54000     grdWardList.ToolTipText = "Click to Toggle"
54010 ElseIf grdWardList.MouseCol = 1 Then
54020     grdWardList.ToolTipText = "Click to Edit"
54030 ElseIf grdWardList.MouseCol = 5 Then
54040     grdWardList.ToolTipText = "Click to Set"
54050 Else
54060     grdWardList.ToolTipText = "Click to Move"
54070 End If

End Sub


Private Sub lHospital_Click()

54080 FillG

End Sub

Private Sub tmrDown_Timer()

54090 FireDown

End Sub


Private Sub tmrUp_Timer()

54100 FireUp

End Sub


