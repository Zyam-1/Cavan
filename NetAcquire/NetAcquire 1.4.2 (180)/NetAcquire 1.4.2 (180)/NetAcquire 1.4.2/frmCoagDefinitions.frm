VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCoagDefinitions 
   Caption         =   "NetAcquire - Coagulation Definitions"
   ClientHeight    =   6990
   ClientLeft      =   465
   ClientTop       =   1590
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   8370
   Begin VB.Frame Frame7 
      Caption         =   "Specifics (Applies to all age ranges)"
      Height          =   1845
      Left            =   1440
      TabIndex        =   5
      Top             =   4950
      Width           =   5325
      Begin VB.CheckBox chkShowRefRange 
         Caption         =   "Print Ref Range"
         Height          =   225
         Left            =   3630
         TabIndex        =   57
         Top             =   1290
         Width           =   1605
      End
      Begin VB.CheckBox chkInUse 
         Caption         =   "In Use"
         Height          =   195
         Left            =   3630
         TabIndex        =   56
         Top             =   960
         Width           =   765
      End
      Begin VB.TextBox tTestName 
         Height          =   285
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cUnits 
         Height          =   315
         Left            =   3630
         TabIndex        =   8
         Text            =   "cUnits"
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox tCode 
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox cPrintable 
         Caption         =   "Printable"
         Height          =   225
         Left            =   3630
         TabIndex        =   6
         Top             =   630
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.Label lblBarCode 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1230
         TabIndex        =   55
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "BarCode"
         Height          =   195
         Left            =   540
         TabIndex        =   54
         Top             =   1290
         Width           =   615
      End
      Begin VB.Label lDP 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   1230
         TabIndex        =   14
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Decimal Points"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Test Name"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   630
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   3240
         TabIndex        =   11
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   780
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   1410
      TabIndex        =   2
      Top             =   90
      Width           =   6615
      Begin VB.CommandButton bAmendAgeRange 
         Caption         =   "Amend Age Range"
         Height          =   525
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   870
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   1725
         Left            =   570
         TabIndex        =   4
         Top             =   270
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3043
         _Version        =   393216
         Cols            =   3
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         FormatString    =   "^Range #  |^Age From        |^Age To           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   7020
      Picture         =   "frmCoagDefinitions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5730
      Width           =   1155
   End
   Begin VB.ListBox lstParameter 
      Height          =   4665
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2625
      Left            =   1440
      TabIndex        =   15
      Top             =   2250
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4630
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Normal Range"
      TabPicture(0)   =   "frmCoagDefinitions.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tMaleHigh"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tFemaleHigh"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tMaleLow"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tFemaleLow"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmCoagDefinitions.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tFlagFemaleLow"
      Tab(1).Control(1)=   "tFlagMaleHigh"
      Tab(1).Control(2)=   "tFlagFemaleHigh"
      Tab(1).Control(3)=   "tFlagMaleLow"
      Tab(1).Control(4)=   "Label7(1)"
      Tab(1).Control(5)=   "Label12(2)"
      Tab(1).Control(6)=   "Label13(1)"
      Tab(1).Control(7)=   "Label14(2)"
      Tab(1).Control(8)=   "Label15(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Plausible"
      TabPicture(2)   =   "frmCoagDefinitions.frx":0F02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tPlausibleLow"
      Tab(2).Control(1)=   "tPlausibleHigh"
      Tab(2).Control(2)=   "Label10(2)"
      Tab(2).Control(3)=   "Label9(1)"
      Tab(2).Control(4)=   "Label8(1)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Delta Check"
      TabPicture(3)   =   "frmCoagDefinitions.frx":0F1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tDeltaDaysBackLimit"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "tDelta"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "oDelta"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label17"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label20"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.TextBox tDeltaDaysBackLimit 
         Height          =   285
         Left            =   -72690
         MaxLength       =   5
         TabIndex        =   58
         Top             =   1650
         Width           =   555
      End
      Begin VB.TextBox tFlagFemaleLow 
         Height          =   315
         Left            =   -73500
         TabIndex        =   53
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   0
         Left            =   -71640
         TabIndex        =   30
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   2
         Left            =   -73170
         TabIndex        =   29
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   1
         Left            =   -71640
         TabIndex        =   28
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   3
         Left            =   -73170
         TabIndex        =   27
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tFemaleLow 
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   26
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox tMaleLow 
         Height          =   315
         Left            =   3030
         MaxLength       =   5
         TabIndex        =   25
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox tFemaleHigh 
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   24
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tMaleHigh 
         Height          =   315
         Left            =   3030
         MaxLength       =   5
         TabIndex        =   23
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -73080
         TabIndex        =   22
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -73080
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tDelta 
         Height          =   285
         Left            =   -72690
         MaxLength       =   5
         TabIndex        =   20
         Top             =   1320
         Width           =   555
      End
      Begin VB.CheckBox oDelta 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   -73050
         TabIndex        =   19
         Top             =   930
         Width           =   915
      End
      Begin VB.TextBox tFlagMaleHigh 
         Height          =   315
         Left            =   -71970
         TabIndex        =   18
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tFlagFemaleHigh 
         Height          =   315
         Left            =   -73500
         TabIndex        =   17
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tFlagMaleLow 
         Height          =   315
         Left            =   -71970
         TabIndex        =   16
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Back Days Limit"
         Height          =   195
         Left            =   -73920
         TabIndex        =   59
         Top             =   1650
         Width           =   1140
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   52
         Top             =   1590
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   -73110
         TabIndex        =   51
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   0
         Left            =   -74190
         TabIndex        =   50
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "These are the normal range values printed on the report forms."
         Height          =   195
         Left            =   510
         TabIndex        =   49
         Top             =   1980
         Width           =   4395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Do not print the result if the sample is :-"
         Height          =   195
         Index           =   1
         Left            =   -74070
         TabIndex        =   48
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   1
         Left            =   -73020
         TabIndex        =   47
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   0
         Left            =   -71370
         TabIndex        =   46
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Index           =   1
         Left            =   -73950
         TabIndex        =   45
         Top             =   1290
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   0
         Left            =   -73920
         TabIndex        =   44
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   930
         TabIndex        =   43
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   930
         TabIndex        =   42
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   1650
         TabIndex        =   41
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   40
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Results outside this range will be marked as implausible."
         Height          =   195
         Index           =   2
         Left            =   -74220
         TabIndex        =   39
         Top             =   1950
         Width           =   3930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -73560
         TabIndex        =   38
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   -73590
         TabIndex        =   37
         Top             =   750
         Width           =   330
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   -73200
         TabIndex        =   36
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   35
         Top             =   1980
         Width           =   4410
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   2
         Left            =   -73350
         TabIndex        =   34
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   -71700
         TabIndex        =   33
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   -74070
         TabIndex        =   32
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -74070
         TabIndex        =   31
         Top             =   1530
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmCoagDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FromDays() As Long
Private ToDays() As Long



Private Sub FillAges()

          Dim tb As Recordset
          Dim s As String
          Dim n As Integer
          Dim sql As String

16860     On Error GoTo FillAges_Error

16870     g.Rows = 2
16880     g.AddItem ""
16890     g.RemoveItem 1

16900     ReDim FromDays(0 To 0)
16910     ReDim ToDays(0 To 0)

16920     sql = "Select * from CoagTestDefinitions where " & _
              "TestName = '" & lstParameter & "' " & _
              "and Hospital = '" & HospName(0) & "' " & _
              "Order by cast(AgetoDays as numeric) asc"
16930     Set tb = New Recordset
16940     RecOpenClient 0, tb, sql
16950     If Not tb.EOF Then
16960         ReDim FromDays(0 To tb.RecordCount - 1)
16970         ReDim ToDays(0 To tb.RecordCount - 1)
16980         n = 0
16990         Do While Not tb.EOF
17000             FromDays(n) = tb!AgeFromDays
17010             ToDays(n) = tb!AgeToDays
17020             s = Format$(n) & vbTab & _
                      dmyFromCount(FromDays(n)) & vbTab & _
                      dmyFromCount(ToDays(n))
17030             g.AddItem s
17040             n = n + 1
17050             tb.MoveNext
17060         Loop
17070     End If

17080     If g.Rows > 2 Then
17090         g.RemoveItem 1
17100     End If

17110     g.Col = 0
17120     g.row = 1
17130     g.CellBackColor = vbYellow
17140     g.CellForeColor = vbBlue

17150     Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

17160     intEL = Erl
17170     strES = Err.Description
17180     LogError "fCoagDefinitions", "FillAges", intEL, strES, sql

End Sub

Private Sub SaveDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Integer

17190     On Error GoTo SaveDetails_Error

17200     g.Col = 0
17210     For Y = 1 To g.Rows - 1
17220         g.row = Y
17230         If g.CellBackColor = vbYellow Then
        
17240             sql = "Select * from CoagTestDefinitions where " & _
                      "TestName = '" & lstParameter & "' " & _
                      "and AgeFromDays = '" & FromDays(Y - 1) & "' " & _
                      "and AgeToDays = '" & ToDays(Y - 1) & "'"
17250             Set tb = New Recordset
17260             RecOpenClient 0, tb, sql
17270             With tb
17280                 If .EOF Then .AddNew

17290                 !MaleLow = Val(tMaleLow)
17300                 !MaleHigh = Val(tMaleHigh)
17310                 !FemaleLow = Val(tFemaleLow)
17320                 !FemaleHigh = Val(tFemaleHigh)
17330                 !FlagMaleLow = Val(tFlagMaleLow)
17340                 !FlagMaleHigh = Val(tFlagMaleHigh)
17350                 !FlagFemaleLow = Val(tFlagFemaleLow)
17360                 !FlagFemaleHigh = Val(tFlagFemaleHigh)
17370                 !PlausibleLow = Val(tPlausibleLow)
17380                 !PlausibleHigh = Val(tPlausibleHigh)
17390                 !AgeFromDays = FromDays(Y - 1)
17400                 !AgeToDays = ToDays(Y - 1)
17410                 !InUse = 1
17420                 !DoDelta = 0
17430                 !DeltaDaysBackLimit = Val(tDeltaDaysBackLimit & "")
17440                 !Displayable = 1
17450                 !Printable = 1
17460                 .Update
17470             End With
          
17480             sql = "Update CoagTestDefinitions " & _
                      "Set Code = '" & tCode & "', " & _
                      "DoDelta = " & IIf(oDelta = 1, 1, 0) & ", " & _
                      "DeltaDaysBackLimit = " & Val(tDeltaDaysBackLimit & "") & ", " & _
                      "DeltaLimit = '" & tDelta & "', " & _
                      "PrintPriority = '" & lstParameter.ListIndex & "', " & _
                      "DP = '" & lDP & "', " & _
                      "Units = '" & cUnits & "', " & _
                      "Printable = " & IIf(cPrintable = 1, 1, 0) & ", " & _
                      "BarCode = '" & Trim$(lblBarCode) & "', " & _
                      "InUse = '" & IIf(chkInUse = 0, 0, 1) & "', " & _
                      "PrintRefRange = '" & IIf(chkShowRefRange = 0, 0, 1) & "' " & _
                      "where TestName = '" & lstParameter & "'"
17490             Cnxn(0).Execute sql
          
17500             Exit For
17510         End If
17520     Next

17530     Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

17540     intEL = Erl
17550     strES = Err.Description
17560     LogError "fCoagDefinitions", "SaveDetails", intEL, strES, sql

End Sub

Private Sub bAmendAgeRange_Click()

17570     On Error GoTo bAmendAgeRange_Click_Error

17580     If lstParameter = "" Then
17590         iMsg "Select Parameter", vbCritical
17600         Exit Sub
17610     End If

17620     With frmAges
17630         .Analyte = lstParameter
17640         .SampleType = "Coagulation"
17650         .Discipline = "Coagulation"
17660         .Show 1
17670     End With

17680     FillAges

17690     Exit Sub

bAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer

17700     intEL = Erl
17710     strES = Err.Description
17720     LogError "fCoagDefinitions", "bAmendAgeRange_Click", intEL, strES

End Sub

Private Sub bcancel_Click()

17730     Unload Me

End Sub


Private Sub chkInUse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17740     SaveDetails

End Sub


Private Sub chkShowRefRange_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
17750     SaveDetails
End Sub


Private Sub cPrintable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17760     SaveDetails

End Sub


Private Sub cUnits_Click()

17770     SaveDetails

End Sub


Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

17780     On Error GoTo Form_Load_Error

17790     cUnits.Clear
17800     sql = "Select * from Lists where " & _
              "ListType = 'UN' and InUse = 1 " & _
              "order by ListOrder"
17810     Set tb = New Recordset
17820     RecOpenServer 0, tb, sql
17830     Do While Not tb.EOF
17840         cUnits.AddItem tb!Text & ""
17850         tb.MoveNext
17860     Loop
17870     If cUnits.ListCount > 0 Then
17880         cUnits.ListIndex = 0
17890     End If

17900     FillParameters
17910     FillAges
17920     FillDetails

17930     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

17940     intEL = Erl
17950     strES = Err.Description
17960     LogError "fCoagDefinitions", "Form_Load", intEL, strES, sql

End Sub

Private Sub FillParameters()

          Dim sql As String
          Dim tb As Recordset
          Dim InList As Boolean
          Dim n As Integer

17970     On Error GoTo FillParameters_Error

17980     lstParameter.Clear

17990     sql = "Select * from CoagTestDefinitions " & _
              "Order by PrintPriority"
18000     Set tb = New Recordset
18010     RecOpenServer 0, tb, sql
18020     Do While Not tb.EOF
18030         InList = False
18040         For n = 0 To lstParameter.ListCount - 1
18050             If lstParameter.List(n) = tb!TestName Then
18060                 InList = True
18070                 Exit For
18080             End If
18090         Next
18100         If Not InList Then
18110             lstParameter.AddItem tb!TestName
18120         End If
18130         tb.MoveNext
18140     Loop

18150     If lstParameter.ListCount > 0 Then
18160         lstParameter.Selected(0) = True
18170     End If

18180     Exit Sub

FillParameters_Error:

          Dim strES As String
          Dim intEL As Integer

18190     intEL = Erl
18200     strES = Err.Description
18210     LogError "fCoagDefinitions", "FillParameters", intEL, strES, sql

End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim n As Integer
          Dim ySave As Integer

18220     If g.MouseRow = 0 Then Exit Sub

18230     ySave = g.row

18240     g.Col = 0
18250     For n = 1 To g.Rows - 1
18260         g.row = n
18270         g.CellBackColor = 0
18280         g.CellForeColor = 0
18290     Next
        
18300     g.row = ySave
18310     g.CellBackColor = vbYellow
18320     g.CellForeColor = vbBlue

18330     FillDetails

End Sub

Private Sub lblBarCode_Click()

18340     lblBarCode = iBOX("Scan in Bar Code", , lblBarCode)

18350     SaveDetails

End Sub

Private Sub lDP_Click()

18360     lDP = Format$(Val(lDP) + 1)

18370     If Val(lDP) > 3 Then lDP = "0"

18380     SaveDetails

End Sub

Private Sub FillDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim Filled As Boolean
          Dim AgeNumber As Integer
          Dim Y As Integer

18390     On Error GoTo FillDetails_Error

18400     Filled = False

18410     AgeNumber = -1
18420     g.Col = 0
18430     For Y = 1 To g.Rows - 1
18440         g.row = Y
18450         If g.CellBackColor = vbYellow Then
18460             AgeNumber = Y - 1
18470             Exit For
18480         End If
18490     Next
18500     If AgeNumber = -1 Then
18510         iMsg "Select Age Range", vbCritical
18520         Exit Sub
18530     End If

18540     tCode = ""
18550     tTestName = ""
18560     oDelta = 0
18570     tDelta = ""
18580     lDP = "0"
18590     cUnits = ""
18600     cPrintable = 0
18610     tPlausibleLow = ""
18620     tPlausibleHigh = ""
18630     tMaleHigh = ""
18640     tFemaleHigh = ""
18650     tMaleLow = ""
18660     tFemaleLow = ""
18670     lblBarCode = ""
18680     chkInUse = 0
18690     chkShowRefRange = 0

18700     sql = "Select * from CoagTestDefinitions where " & _
              "TestName = '" & lstParameter & "' " & _
              "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
              "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
              "And Hospital = '" & HospName(0) & "'"
18710     Set tb = New Recordset
18720     RecOpenClient 0, tb, sql
18730     If Not tb.EOF Then
18740         tCode = tb!Code
18750         tTestName = lstParameter
18760         oDelta = IIf(tb!DoDelta, 1, 0)
18770         tDelta = IIf(IsNull(tb!DeltaLimit), "", tb!DeltaLimit)
18780         tDeltaDaysBackLimit = IIf(IsNull(tb!DeltaDaysBackLimit), "", tb!DeltaDaysBackLimit)
18790         lDP = tb!DP
18800         cUnits = tb!Units & ""
18810         cPrintable = IIf(tb!Printable, 1, 0)
18820         tPlausibleLow = IIf(IsNull(tb!PlausibleLow), "", tb!PlausibleLow)
18830         tPlausibleHigh = IIf(IsNull(tb!PlausibleHigh), "", tb!PlausibleHigh)
18840         tMaleHigh = IIf(IsNull(tb!MaleHigh), "", tb!MaleHigh)
18850         tFemaleHigh = IIf(IsNull(tb!FemaleHigh), "", tb!FemaleHigh)
18860         tMaleLow = IIf(IsNull(tb!MaleLow), "", tb!MaleLow)
18870         tFemaleLow = IIf(IsNull(tb!FemaleLow), "", tb!FemaleLow)
18880         tFlagMaleHigh = IIf(IsNull(tb!FlagMaleHigh), "", tb!FlagMaleHigh)
18890         tFlagFemaleHigh = IIf(IsNull(tb!FlagFemaleHigh), "", tb!FlagFemaleHigh)
18900         tFlagMaleLow = IIf(IsNull(tb!FlagMaleLow), "", tb!FlagMaleLow)
18910         tFlagFemaleLow = IIf(IsNull(tb!FlagFemaleLow), "", tb!FlagFemaleLow)
18920         lblBarCode = tb!BarCode & ""
18930         If Not IsNull(tb!InUse) Then
18940             chkInUse = IIf(tb!InUse = 0, 0, 1)
18950         End If
18960         chkShowRefRange = IIf(tb!PrintRefRange = 0, 0, 1)
18970     End If

18980     Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

18990     intEL = Erl
19000     strES = Err.Description
19010     LogError "fCoagDefinitions", "FillDetails", intEL, strES, sql

End Sub


Private Sub lstParameter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19020     FillAges
19030     FillDetails

End Sub

Private Sub odelta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19040     SaveDetails

End Sub


Private Sub tCode_KeyUp(KeyCode As Integer, Shift As Integer)

19050     SaveDetails

End Sub


Private Sub tDelta_KeyUp(KeyCode As Integer, Shift As Integer)

19060     SaveDetails

End Sub


Private Sub tDeltaDaysBackLimit_KeyUp(KeyCode As Integer, Shift As Integer)
19070     SaveDetails
End Sub

Private Sub tFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

19080     SaveDetails

End Sub


Private Sub tFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

19090     SaveDetails

End Sub


Private Sub tFlagFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

19100     SaveDetails

End Sub


Private Sub tFlagFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

19110     SaveDetails

End Sub


Private Sub tFlagMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

19120     SaveDetails

End Sub


Private Sub tFlagMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

19130     SaveDetails

End Sub


Private Sub tMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

19140     SaveDetails

End Sub


Private Sub tMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

19150     SaveDetails

End Sub


Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

19160     SaveDetails

End Sub


Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

19170     SaveDetails

End Sub


