VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestNormalRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Normal Ranges"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   7080
      TabIndex        =   13
      Text            =   "cmbCategory"
      Top             =   390
      Width           =   1485
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Age Ranges"
      Height          =   1100
      Left            =   5040
      Picture         =   "frmTestNormalRange.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5910
      Width           =   1200
   End
   Begin VB.ComboBox cmbInput 
      BackColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "frmTestNormalRange.frx":0ECA
      Left            =   9420
      List            =   "frmTestNormalRange.frx":0ECC
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   210
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1100
      Left            =   6390
      Picture         =   "frmTestNormalRange.frx":0ECE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5910
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   7710
      Picture         =   "frmTestNormalRange.frx":2850
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5910
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   9030
      Picture         =   "frmTestNormalRange.frx":846E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5910
      Width           =   1200
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   4530
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "cmbSampleType"
      Top             =   390
      Width           =   2295
   End
   Begin VB.ListBox lstNames 
      Height          =   6600
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   330
      TabIndex        =   9
      Top             =   7110
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdRange 
      Height          =   4995
      Left            =   1770
      TabIndex        =   11
      Top             =   750
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   8811
      _Version        =   393216
      Rows            =   3
      Cols            =   13
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmTestNormalRange.frx":9338
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Index           =   1
      Left            =   7470
      TabIndex        =   14
      Top             =   180
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SampleType"
      Height          =   195
      Left            =   5100
      TabIndex        =   7
      Top             =   180
      Width           =   885
   End
   Begin VB.Label lblLongName 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1770
      TabIndex        =   6
      Top             =   390
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Long Name"
      Height          =   195
      Index           =   0
      Left            =   2490
      TabIndex        =   5
      Top             =   180
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Short Name"
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmTestNormalRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arrDay
  Text As String
  Value As Long
End Type
Private arrDays(0 To 36) As arrDay

Private mDept As String
Private pSampleType1 As String

Private Sub EnableControls(ByVal Enable As Boolean)

13990 lstNames.Enabled = Enable
14000 cmbSampleType.Enabled = Enable
14010 grdRange.Enabled = Enable
14020 cmdExport.Enabled = Enable

End Sub

Private Sub EnableSelection(ByVal Enable As Boolean)

14030 lstNames.Enabled = Enable
14040 cmbSampleType.Enabled = Enable

End Sub
Private Sub FillGrid()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

14050 On Error GoTo FillGrid_Error

14060 With grdRange
14070   .Rows = 3
14080   .AddItem ""
14090   .RemoveItem 2
14100 End With

14110 sql = "SELECT LongName, AgeFromDays, AgeToDays, AgeFromText, AgeToText, " & _
            "COALESCE(MaleLow, 0) AS MaleLow, " & _
            "COALESCE(MaleHigh, 9999) AS MaleHigh, " & _
            "COALESCE(FemaleLow, 0) AS FemaleLow, " & _
            "COALESCE(FemaleHigh, 9999) AS FemaleHigh, " & _
            "COALESCE(FlagMaleLow, 0) AS FlagMaleLow, " & _
            "COALESCE(FlagMaleHigh, 9999) AS FlagMaleHigh, " & _
            "COALESCE(FlagFemaleLow, 0) AS FlagFemaleLow, " & _
            "COALESCE(FlagFemaleHigh, 9999) AS FlagFemaleHigh, " & _
            "CASE COALESCE(PrintRefRange, 1) " & _
            "  WHEN 1 THEN 'Yes' ELSE 'No' END PrnRefRange " & _
            "FROM " & mDept & "TestDefinitions WHERE " & _
            "ShortName = '" & lstNames & "' " & _
            "AND SampleType = '" & pSampleType1 & "' " & _
            "AND Category = '" & cmbCategory & "' " & _
            "AND InUse = 1 ORDER BY AgeFromDays"
14120 Set tb = New Recordset
14130 RecOpenServer 0, tb, sql
14140 Do While Not tb.EOF
14150   lblLongName = tb!LongName & ""
14160   s = tb!AgeFromText & vbTab & _
            tb!AgeToText & vbTab & _
            tb!MaleLow & vbTab & _
            tb!MaleHigh & vbTab & _
            tb!FemaleLow & vbTab & _
            tb!FemaleHigh & vbTab & _
            tb!AgeFromDays & vbTab & _
            tb!AgeToDays & vbTab & _
            tb!FlagMaleLow & vbTab & _
            tb!FlagMaleHigh & vbTab & _
            tb!FlagFemaleLow & vbTab & _
            tb!FlagFemaleHigh & vbTab & tb!PrnRefRange & ""
14170   grdRange.AddItem s
14180   tb.MoveNext
14190 Loop

14200 With grdRange
14210   If .Rows > 3 Then
14220     .RemoveItem 2
14230     .AddItem ""
14240   End If
14250 End With

14260 Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

14270 intEL = Erl
14280 strES = Err.Description
14290 LogError "frmTestNormalRange", "FillGrid", intEL, strES, sql

End Sub

Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String

14300 On Error GoTo FillList_Error

14310 lstNames.Clear

14320 sql = "SELECT DISTINCT ShortName, PrintPriority FROM " & mDept & "TestDefinitions WHERE " & _
            "SampleType = '" & pSampleType1 & "' " & _
            "AND COALESCE(InUse, 1) = 1 " & _
            "ORDER BY PrintPriority, ShortName"
14330 Set tb = New Recordset
14340 RecOpenServer 0, tb, sql
14350 Do While Not tb.EOF
14360   lstNames.AddItem tb!ShortName & ""
14370   tb.MoveNext
14380 Loop

14390 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

14400 intEL = Erl
14410 strES = Err.Description
14420 LogError "frmTestNormalRange", "FillList", intEL, strES, sql

End Sub

Private Sub FillSampleTypes()

      Dim sql As String
      Dim tb As Recordset

14430 On Error GoTo FillSampleTypes_Error

14440 sql = "SELECT Text FROM Lists " & _
            "WHERE ListType = 'ST' " & _
            "ORDER BY ListOrder"
14450 Set tb = New Recordset
14460 RecOpenServer 0, tb, sql

14470 cmbSampleType.Clear
14480 Do While Not tb.EOF
14490   cmbSampleType.AddItem tb!Text & ""
14500   tb.MoveNext
14510 Loop

14520 If pSampleType1 <> "" Then
14530   cmbSampleType = ListTextFor("ST", pSampleType1)
14540 Else
14550   pSampleType1 = "S"
14560   cmbSampleType = "Serum"
14570 End If

14580 Exit Sub

FillSampleTypes_Error:

      Dim strES As String
      Dim intEL As Integer

14590 intEL = Erl
14600 strES = Err.Description
14610 LogError "frmTestNormalRange", "FillSampleTypes", intEL, strES, sql


End Sub

Private Sub cmbInput_Click()

      Dim Y As Integer
      Dim yy As Integer

14620 cmbInput.Visible = False
14630 With grdRange
14640   If .Col = 0 Then 'Adding a new Age Range
14650     .TextMatrix(.row, 0) = cmbInput
14660     .TextMatrix(.row, 6) = cmbInput.ItemData(cmbInput.ListIndex)
14670     .TextMatrix(.row, 1) = "120 Years"
14680     .TextMatrix(.row, 7) = "43830"
14690     .TextMatrix(.row - 1, 1) = cmbInput
14700     .TextMatrix(.row - 1, 7) = cmbInput.ItemData(cmbInput.ListIndex)
14710     .TextMatrix(.row, 2) = "0"
14720     .TextMatrix(.row, 3) = "9999"
14730     .TextMatrix(.row, 4) = "0"
14740     .TextMatrix(.row, 5) = "9999"
14750     .TextMatrix(.row, 8) = "0"
14760     .TextMatrix(.row, 9) = "9999"
14770     .TextMatrix(.row, 10) = "0"
14780     .TextMatrix(.row, 11) = "9999"
          
14790     .AddItem ""
          
14800   Else
          
14810     .TextMatrix(.row, 1) = cmbInput
14820     .TextMatrix(.row, 7) = cmbInput.ItemData(cmbInput.ListIndex)
14830     .TextMatrix(.row + 1, 0) = cmbInput
14840     .TextMatrix(.row + 1, 6) = cmbInput.ItemData(cmbInput.ListIndex)

14850   End If
        
14860   For Y = 2 To .Rows - 1
14870     If .TextMatrix(Y, 1) = "120 Years" Then
14880       For yy = .Rows - 1 To Y + 1 Step -1
14890         .RemoveItem yy
14900       Next
14910       .AddItem ""
14920       Exit For
14930     End If
          
14940     .TextMatrix(Y + 1, 0) = .TextMatrix(Y, 1)
14950     .TextMatrix(Y + 1, 6) = GetValueFromText(.TextMatrix(Y + 1, 0))
14960     If .TextMatrix(Y + 1, 1) = "120 Years" Then
14970       Exit For
14980     End If
14990     If Val(.TextMatrix(Y + 1, 7)) <= Val(.TextMatrix(Y + 1, 6)) Then
15000       .TextMatrix(Y + 1, 1) = GetNextTextFromText(.TextMatrix(Y + 1, 0))
15010       .TextMatrix(Y + 1, 7) = GetValueFromText(.TextMatrix(Y + 1, 1))
15020       If .TextMatrix(Y + 1, 1) <> "120 Years" Then
15030         .TextMatrix(Y + 2, 0) = .TextMatrix(Y + 1, 1)
15040         .TextMatrix(Y + 2, 6) = .TextMatrix(Y + 1, 7)
15050       End If
15060     End If
15070   Next
        
15080   grdRange.Enabled = True
15090   .SetFocus

15100 End With

15110 EnableSelection False
15120 cmdSave.Visible = True

End Sub

Private Sub cmbInput_KeyPress(KeyAscii As Integer)

15130 If KeyAscii = 13 Then
15140   KeyAscii = 0
15150   cmbInput_Click
15160 End If
15170 KeyAscii = 0

End Sub


Private Sub cmbSampleType_Click()

15180 pSampleType1 = ListCodeFor("ST", cmbSampleType)

15190 FillList

15200 grdRange.Rows = 3
15210 grdRange.AddItem ""
15220 grdRange.RemoveItem 2

End Sub


Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

15230 KeyAscii = 0

End Sub


Private Sub cmdExit_Click()

15240 If cmdSave.Visible Then
15250   If iMsg("Cancel without Saving?", vbYesNo + vbQuestion) = vbYes Then
15260     Unload Me
15270   End If
15280 Else
15290   Unload Me
15300 End If

End Sub

Private Sub cmdExport_Click()

15310 ExportFlexGrid grdRange, Me

End Sub

Private Sub cmdRemove_Click()

      Dim tb As Recordset
      Dim sql As String

15320 On Error GoTo cmdRemove_Click_Error

15330 sql = "SELECT * FROM " & mDept & "TestDefinitions WHERE " & _
            "SampleType = '" & pSampleType1 & "' " & _
            "AND ShortName = '" & lstNames & "'"
15340 Set tb = New Recordset
15350 RecOpenClient 0, tb, sql
15360 If Not tb.EOF Then
15370   tb.MoveNext
15380   Do While Not tb.EOF
15390     tb.Delete
15400     tb.MoveNext
15410   Loop
15420   tb.MoveFirst
15430   tb!AgeFromDays = 0
15440   tb!AgeToDays = 43830
15450   tb!AgeFromText = "0 Days"
15460   tb!AgeToText = "120 Years"
15470   tb.Update
15480 End If

15490 FillGrid

15500 Exit Sub

cmdRemove_Click_Error:

      Dim strES As String
      Dim intEL As Integer

15510 intEL = Erl
15520 strES = Err.Description
15530 LogError "frmTestNormalRange", "cmdRemove_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim Y As Integer
      Dim InsertOrUpdate As String
      Dim Found As Boolean

15540 On Error GoTo cmdSave_Click_Error

15550 For Y = 2 To grdRange.Rows - 2
        
15560   sql = "SELECT * FROM " & mDept & "TestDefinitions WHERE " & _
              "ShortName = '" & lstNames & "' " & _
              "AND SampleType = '" & pSampleType1 & "' " & _
              "AND AgeFromDays = '" & grdRange.TextMatrix(Y, 6) & "' " & _
              "AND AgeToDays = '" & grdRange.TextMatrix(Y, 7) & "' " & _
              "and Category = '" & cmbCategory & "' " & _
              "AND InUse = 1"
15570   Set tb = New Recordset
15580   RecOpenClient 0, tb, sql
15590   InsertOrUpdate = "UPDATE"
15600   If tb.EOF Then 'new age range
15610     InsertOrUpdate = "INSERT"
15620     sql = "SELECT TOP 1 * FROM " & mDept & "TestDefinitions WHERE " & _
                "ShortName = '" & lstNames & "' " & _
                "AND SampleType = '" & pSampleType1 & "' " & _
                "AND InUse = 1"
15630     Set tb = New Recordset
15640     RecOpenClient 0, tb, sql
15650     If tb.EOF Then
15660       Exit Sub
15670     End If
15680   End If

15690   If InsertOrUpdate = "UPDATE" Then
          
15700     tb!MaleLow = Val(grdRange.TextMatrix(Y, 2))
15710     tb!MaleHigh = Val(grdRange.TextMatrix(Y, 3))
15720     tb!FemaleLow = Val(grdRange.TextMatrix(Y, 4))
15730     tb!FemaleHigh = Val(grdRange.TextMatrix(Y, 5))
15740     tb!FlagMaleLow = Val(grdRange.TextMatrix(Y, 8))
15750     tb!FlagMaleHigh = Val(grdRange.TextMatrix(Y, 9))
15760     tb!FlagFemaleLow = Val(grdRange.TextMatrix(Y, 10))
15770     tb!FlagFemaleHigh = Val(grdRange.TextMatrix(Y, 11))
15780     tb!PrintRefRange = IIf(grdRange.TextMatrix(Y, 12) = "Yes", 1, 0)
15790     tb.Update
        
15800   Else
        
15810     sql = "INSERT INTO " & mDept & "TestDefinitions " & _
                "( Code, LongName, ShortName, " & _
                "  AgeFromDays, AgeToDays, AgeFromText, AgeToText, " & _
                "  KnownToAnalyser, " & _
                "  DoDelta, DeltaLimit, " & _
                "  PrintPriority, DP, " & _
                "  BarCode, Units, Printable, " & _
                "  PlausibleLow, PlausibleHigh, " & _
                "  MaleLow, MaleHigh, " & _
                "  FemaleLow, FemaleHigh, " & _
                "  FlagMaleLow, FlagMaleHigh, " & _
                "  FlagFemaleLow, FlagFemaleHigh, " & _
                "  AutoValLow, AutoValHigh, " & _
                "  Analyser, " & _
          "  LIH, PrintRefRange, SampleType, InUse, H, S, L, O, G, J, Category, Hospital, ArchitectCode"
15820     sql = sql & " ) VALUES (" & _
                "'" & tb!Code & "', '" & tb!LongName & "', '" & tb!ShortName & "', " & _
                "'" & grdRange.TextMatrix(Y, 6) & "', '" & grdRange.TextMatrix(Y, 7) & "', '" & grdRange.TextMatrix(Y, 0) & "', '" & grdRange.TextMatrix(Y, 1) & "', " & _
                "'" & IIf(tb!KnownToAnalyser, 1, 0) & "', " & _
                "'" & IIf(tb!DoDelta, 1, 0) & "', '" & IIf(IsNull(tb!DeltaLimit), "", tb!DeltaLimit) & "', " & _
                "'" & tb!PrintPriority & "', '" & tb!DP & "', " & _
                "'" & IIf(IsNull(tb!BarCode), "", tb!BarCode) & "', '" & IIf(IsNull(tb!Units), "", tb!Units) & "', '" & IIf(tb!Printable, 1, 0) & "', " & _
                "'" & Replace(IIf(IsNull(tb!PlausibleLow), "", tb!PlausibleLow), ",", ".") & "', '" & Replace(IIf(IsNull(tb!PlausibleHigh), "", tb!PlausibleHigh), ",", ".") & "', " & _
                "'" & Replace(grdRange.TextMatrix(Y, 2), ",", ".") & "', '" & Replace(grdRange.TextMatrix(Y, 3), ",", ".") & "', " & _
                "'" & Replace(grdRange.TextMatrix(Y, 4), ",", ".") & "', '" & Replace(grdRange.TextMatrix(Y, 5), ",", ".") & "', " & _
                "'" & Replace(grdRange.TextMatrix(Y, 8), ",", ".") & "', '" & Replace(grdRange.TextMatrix(Y, 9), ",", ".") & "', " & _
                "'" & Replace(grdRange.TextMatrix(Y, 10), ",", ".") & "', '" & Replace(grdRange.TextMatrix(Y, 11), ",", ".") & "', " & _
                "'" & tb!AutoValLow & "', '" & tb!AutoValHigh & "', " & _
                "'" & tb!Analyser & "', " & _
                "'" & IIf(IsNull(tb!LIH), 0, tb!LIH) & "', " & _
                "'" & IIf(IsNull(tb!PrintRefRange), 0, tb!PrintRefRange) & "', " & _
                "'" & tb!SampleType & "', " & IIf(tb!InUse, 1, 0) & ", " & _
                IIf(tb!H, 1, 0) & ", " & IIf(tb!s, 1, 0) & ", " & IIf(tb!l, 1, 0) & ", " & _
                IIf(tb!o, 1, 0) & ", " & IIf(tb!g, 1, 0) & ", " & IIf(tb!J, 1, 0) & ", " & _
                "'" & tb!Category & "', '" & tb!Hospital & "', '" & tb!ArchitectCode & "')"
15830     Cnxn(0).Execute sql
15840   End If
15850 Next

15860 sql = "SELECT * FROM " & mDept & "TestDefinitions WHERE " & _
            "ShortName = '" & lstNames & "' " & _
            "AND SampleType = '" & pSampleType1 & "' " & _
            "AND InUse = 1 " & _
            "AND Category = '" & cmbCategory & "'"
15870 Set tb = New Recordset
15880 RecOpenClient 0, tb, sql
15890 Do While Not tb.EOF
15900   Found = False
15910   For Y = 2 To grdRange.Rows - 2
15920     If tb!AgeFromDays = Val(grdRange.TextMatrix(Y, 6)) And _
             tb!AgeToDays = Val(grdRange.TextMatrix(Y, 7)) Then
15930       Found = True
15940       Exit For
15950     End If
15960   Next
15970   If Not Found Then
15980     tb.Delete
15990   End If
16000   tb.MoveNext
16010 Loop
        
16020 EnableControls True
16030 cmdSave.Visible = False

16040 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

16050 intEL = Erl
16060 strES = Err.Description
16070 LogError "frmTestNormalRange", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()

16080 Me.Caption = "NetAcquire = Normal Ranges"

16090 Select Case mDept
        Case "TM":
16100   Case "CD4"
16110     cmbSampleType.Enabled = False
16120   Case "BIO"
16130   Case "HAEM"
16140     cmbSampleType.Enabled = False
16150 End Select

End Sub

Private Sub FillCategories()

      Dim sql As String
      Dim tb As Recordset

16160 On Error GoTo FillCategories_Error

16170 sql = "Select Cat from Categorys " & _
            "order by ListOrder"
16180 Set tb = New Recordset
16190 RecOpenServer 0, tb, sql

16200 cmbCategory.Clear
16210 Do While Not tb.EOF
16220   cmbCategory.AddItem tb!Cat & ""
16230   tb.MoveNext
16240 Loop

16250 If cmbCategory.ListCount > 0 Then
16260   cmbCategory.ListIndex = 0
16270 End If

16280 Exit Sub

FillCategories_Error:

      Dim strES As String
      Dim intEL As Integer

16290 intEL = Erl
16300 strES = Err.Description
16310 LogError "frmTestNormalRange", "FillCategories", intEL, strES, sql


End Sub

Private Sub Form_Load()

      Dim n As Integer

16320 On Error GoTo Form_Load_Error

16330 InitializeGrid
16340 cmbInput.Clear
16350 For n = 0 To 35
16360   arrDays(n).Text = Choose(n + 1, "1 Day", "2 Days", "3 Days", "4 Days", "5 Days", "6 Days", _
                                "1 Week", "2 Weeks", "3 Weeks", _
                                "1 Month", "2 Months", "3 Months", "6 Months", _
                                "1 Year", "2 Years", "3 Years", "4 Years", "5 Years", "6 Years", _
                                "7 Years", "8 Years", "9 Years", "10 Years", "12 Years", "14 Years", _
                                "15 Years", "16 Years", "20 Years", "25 Years", "49 Years", "50 Years", _
                                "60 Years", "70 Years", "75 Years", "80 Years", "120 Years")
        
16370   arrDays(n).Value = Choose(n + 1, 1, 2, 3, 4, 5, 6, _
                                  7, 14, 21, _
                                  30, 60, 90, 180, _
                                  365, 730, 1095, 1461, 1826, 2192, _
                                  2557, 2922, 3287, 3652, 4383, 5110, _
                                  5479, 5844, 7305, 9131, 17885, 18262, _
                                  21900, 25550, 27393, 29200, 43830)
16380   cmbInput.AddItem arrDays(n).Text
16390   cmbInput.ItemData(n) = arrDays(n).Value
16400 Next

16410 FillSampleTypes
16420 FillList
16430 FillCategories
16440 grdRange.Enabled = False

16450 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

16460 intEL = Erl
16470 strES = Err.Description
16480 LogError "frmTestNormalRange", "Form_Load", intEL, strES

End Sub
Private Function GetValueFromText(ByVal Text As String) As Long

      Dim n As Integer

16490 For n = 0 To 31
16500   If arrDays(n).Text = Text Then
16510     GetValueFromText = arrDays(n).Value
16520     Exit For
16530   End If
16540 Next

End Function


Private Function GetNextTextFromText(ByVal Text As String) As String

      Dim n As Integer
      Dim RetVal As String

16550 For n = 0 To 31
16560   If arrDays(n).Text = Text Then
16570     RetVal = arrDays(n + 1).Text
16580     Exit For
16590   End If
16600 Next
16610 If RetVal = "" Then
16620   RetVal = "120 Years"
16630 Else
16640   RetVal = arrDays(n + 1).Text
16650 End If

16660 GetNextTextFromText = RetVal

End Function



Private Sub grdRange_KeyUp(KeyCode As Integer, Shift As Integer)

16670 If grdRange.Col < 2 Or grdRange.Col = 6 Or grdRange.Col = 7 Then Exit Sub
16680 If grdRange.TextMatrix(grdRange.row, 0) = "" Then Exit Sub 'No Age From entered

16690 If EditGrid(grdRange, KeyCode, Shift) Then
16700   cmdSave.Visible = True
16710 End If

End Sub

Private Function EditGrid(ByVal g As MSFlexGrid, _
                         ByVal KeyCode As Integer, _
                         ByVal Shift As Integer) _
                         As Boolean

      'returns true if grid changed

      Dim ShiftDown As Boolean
      Dim RetVal As Boolean

16720 RetVal = False

16730 If g.row < g.FixedRows Then
16740   Exit Function
16750 ElseIf g.Col < g.FixedCols Then
16760   Exit Function
16770 End If
16780 ShiftDown = (Shift And vbShiftMask) > 0

16790 Select Case KeyCode
        Case vbKeyA To vbKeyZ:
16800     If ShiftDown Then
16810       g = g & Chr(KeyCode)
16820       RetVal = True
16830     Else
16840       g = g & Chr(KeyCode + 32)
16850       RetVal = True
16860     End If
        
16870   Case vbKey0 To vbKey9:
16880     g = g & Chr(KeyCode)
16890     RetVal = True
        
16900   Case vbKeyBack:
16910     If Len(g) > 0 Then
16920       g = Left$(g, Len(g) - 1)
16930       RetVal = True
16940     End If
        
16950   Case &HBE, vbKeyDecimal:
16960     g = g & "."
16970     RetVal = True
          
16980   Case vbKeySpace:
16990     g = g & " "
17000     RetVal = True
          
17010   Case vbKeyNumpad0 To vbKeyNumpad9:
17020   Case vbKeyDelete:
17030   Case vbKeyLeft:
17040   Case vbKeyRight:
17050   Case vbKeyUp:
17060   Case vbKeyDown:
17070   Case vbKeyTab:
17080 End Select

17090 If RetVal Then
17100   If g.Col < 8 Then
17110     g.TextMatrix(g.row, g.Col + 6) = g
17120   End If
17130 End If

17140 EditGrid = RetVal

End Function


Private Sub grdRange_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim sql As String

17150 On Error GoTo grdRange_MouseUp_Error

17160 With grdRange
17170   If .MouseRow = 0 Or .MouseRow = 1 Then Exit Sub
17180   If .MouseRow = 2 And .MouseCol = 0 Then Exit Sub
17190   If .MouseCol = 1 And .TextMatrix(.row, 1) = "120 Years" Then Exit Sub
17200   If .MouseCol = 1 And .MouseRow = .Rows - 1 Then Exit Sub
        
17210   Select Case .Col
          Case 0
17220       If .TextMatrix(.row, 0) <> "" Then Exit Sub
17230       cmbInput = ""
17240       cmbInput.Left = .Left + .CellLeft
17250       cmbInput.Top = .Top + .CellTop
17260       cmbInput.width = .CellWidth
17270       EnableControls False
17280       cmbInput.Visible = True
17290       cmbInput.SetFocus

17300     Case 1
17310       cmbInput = .TextMatrix(.row, .Col)
17320       cmbInput.Left = .Left + .CellLeft
17330       cmbInput.Top = .Top + .CellTop
17340       cmbInput.width = .CellWidth
17350       EnableControls False
17360       cmbInput.Visible = True
17370       cmbInput.SetFocus

17380     Case 12
17390       .TextMatrix(.row, 12) = IIf(.TextMatrix(.row, 12) = "No", "Yes", "No")
17400       sql = "UPDATE " & mDept & "TestDefinitions " & _
                  "SET PrintRefRange = '" & IIf(.TextMatrix(.row, 12) = "No", 0, 1) & "' " & _
                  "WHERE ShortName = '" & lstNames & "' " & _
                  "AND SampleType = '" & pSampleType1 & "'"
17410       Cnxn(0).Execute sql

17420     Case Else:
        
17430   End Select
        
17440 End With

17450 Exit Sub

grdRange_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

17460 intEL = Erl
17470 strES = Err.Description
17480 LogError "frmTestNormalRange", "grdRange_MouseUp", intEL, strES, sql

End Sub

Public Property Let Discipline(ByVal strNewValue As String)

17490 mDept = UCase$(strNewValue)

End Property

Public Property Let SampleType(ByVal strNewValue As String)

17500 pSampleType1 = UCase$(strNewValue)

End Property


Private Sub lstNames_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17510 FillGrid

17520 grdRange.Enabled = True

End Sub


Private Sub InitializeGrid()

      Dim i As Integer

17530 With grdRange
17540     .Rows = 3: .FixedRows = 2
17550     .Cols = 13: .FixedCols = 0
          '<Age From          |<Age To            |<Low           |<High           |<Low           |<High _
          |<AgeFromDays |<AgeToDays|<Flag Low|<Flag High |<Flag Low |<Flag High |Ref Range

17560     .TextMatrix(0, 0) = "Age Ranges"
17570     .TextMatrix(0, 1) = "Age Ranges"
17580     .TextMatrix(0, 2) = "Male"
17590     .TextMatrix(0, 3) = "Male"
17600     .TextMatrix(0, 4) = "Female"
17610     .TextMatrix(0, 5) = "Female"
17620     .TextMatrix(0, 6) = ""
17630     .TextMatrix(0, 7) = ""
17640     .TextMatrix(0, 8) = "Male"
17650     .TextMatrix(0, 9) = "Male"
17660     .TextMatrix(0, 10) = "Female"
17670     .TextMatrix(0, 11) = "Female"

17680 .MergeCells = flexMergeRestrictRows
17690 .MergeRow(0) = True

17700     .TextMatrix(1, 0) = "Age From": .ColWidth(0) = 1200: .ColAlignment(0) = flexAlignLeftCenter
17710     .TextMatrix(1, 1) = "Age To": .ColWidth(1) = 1200: .ColAlignment(1) = flexAlignLeftCenter
17720     .TextMatrix(1, 2) = "Low": .ColWidth(2) = 1000: .ColAlignment(2) = flexAlignLeftCenter
17730     .TextMatrix(1, 3) = "High": .ColWidth(3) = 1000: .ColAlignment(3) = flexAlignLeftCenter
17740     .TextMatrix(1, 4) = "Low": .ColWidth(4) = 1000: .ColAlignment(4) = flexAlignLeftCenter
17750     .TextMatrix(1, 5) = "High": .ColWidth(5) = 1000: .ColAlignment(5) = flexAlignLeftCenter
17760     .TextMatrix(1, 8) = "Flag Low": .ColWidth(8) = 1000: .ColAlignment(2) = flexAlignLeftCenter
17770     .TextMatrix(1, 9) = "Flag High": .ColWidth(9) = 1000: .ColAlignment(3) = flexAlignLeftCenter
17780     .TextMatrix(1, 10) = "Flag Low": .ColWidth(10) = 1000: .ColAlignment(4) = flexAlignLeftCenter
17790     .TextMatrix(1, 11) = "Flag High": .ColWidth(11) = 1000: .ColAlignment(5) = flexAlignLeftCenter
17800     .TextMatrix(1, 12) = "Ref Ranges"

17810     For i = 0 To .Cols - 1
17820         If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
17830             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + Me.TextWidth(" ")
17840         End If
17850     Next i
17860     .TextMatrix(1, 6) = "Age From Days": .ColWidth(6) = 0: .ColAlignment(6) = flexAlignLeftCenter
17870     .TextMatrix(1, 7) = "Age To Days": .ColWidth(7) = 0: .ColAlignment(7) = flexAlignLeftCenter
17880 End With

End Sub

