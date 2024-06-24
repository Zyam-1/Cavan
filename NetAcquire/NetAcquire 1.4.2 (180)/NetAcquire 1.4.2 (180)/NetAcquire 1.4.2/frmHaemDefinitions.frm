VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHaemDefinitions 
   Caption         =   "NetAcquire - Haematology Definitions"
   ClientHeight    =   7515
   ClientLeft      =   1200
   ClientTop       =   885
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8535
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   7140
      Picture         =   "frmHaemDefinitions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6030
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Specifics (Applies to all age ranges)"
      Height          =   1245
      Left            =   1500
      TabIndex        =   8
      Top             =   5760
      Width           =   4635
      Begin VB.TextBox tTestName 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   2925
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Test Name"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Decimal Points"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label lDP 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   750
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   1500
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      Begin VB.CommandButton bAmendAgeRange 
         Caption         =   "Amend Age Range"
         Height          =   525
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   870
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   1725
         Left            =   570
         TabIndex        =   3
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
   Begin VB.ListBox lstParameter 
      Height          =   5235
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   0
      Top             =   450
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   1500
      TabIndex        =   9
      Top             =   2700
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5265
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmHaemDefinitions.frx":066A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tFemaleLow"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tMaleLow"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tFemaleHigh"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tMaleHigh"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmHaemDefinitions.frx":0686
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7(1)"
      Tab(1).Control(1)=   "Label12(2)"
      Tab(1).Control(2)=   "Label13(1)"
      Tab(1).Control(3)=   "Label14(2)"
      Tab(1).Control(4)=   "Label15(1)"
      Tab(1).Control(5)=   "tFlagFemaleLow"
      Tab(1).Control(6)=   "tFlagMaleHigh"
      Tab(1).Control(7)=   "tFlagFemaleHigh"
      Tab(1).Control(8)=   "tFlagMaleLow"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Plausible"
      TabPicture(2)   =   "frmHaemDefinitions.frx":06A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10(2)"
      Tab(2).Control(1)=   "Label9(1)"
      Tab(2).Control(2)=   "Label8(1)"
      Tab(2).Control(3)=   "tPlausibleLow"
      Tab(2).Control(4)=   "tPlausibleHigh"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Delta Check"
      TabPicture(3)   =   "frmHaemDefinitions.frx":06BE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tDeltaBackDays"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "oDelta"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "tDelta"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label22"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label20"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Controls"
      TabPicture(4)   =   "frmHaemDefinitions.frx":06DA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label18"
      Tab(4).Control(1)=   "Label19"
      Tab(4).Control(2)=   "Label21"
      Tab(4).Control(3)=   "Label5"
      Tab(4).Control(4)=   "Label16"
      Tab(4).Control(5)=   "Label17"
      Tab(4).Control(6)=   "txtMean(0)"
      Tab(4).Control(7)=   "txt1SD(0)"
      Tab(4).Control(8)=   "cmdSaveControls"
      Tab(4).Control(9)=   "cmbLotNumber(0)"
      Tab(4).Control(10)=   "txt1SD(1)"
      Tab(4).Control(11)=   "txtMean(1)"
      Tab(4).Control(12)=   "cmbLotNumber(1)"
      Tab(4).Control(13)=   "txt1SD(2)"
      Tab(4).Control(14)=   "txtMean(2)"
      Tab(4).Control(15)=   "cmbLotNumber(2)"
      Tab(4).ControlCount=   16
      Begin VB.TextBox tDeltaBackDays 
         Height          =   285
         Left            =   -72210
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1830
         Width           =   555
      End
      Begin VB.ComboBox cmbLotNumber 
         Height          =   315
         Index           =   2
         Left            =   -74040
         TabIndex        =   65
         Text            =   "cmbLotNumber"
         Top             =   1890
         Width           =   1845
      End
      Begin VB.TextBox txtMean 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   -72060
         TabIndex        =   64
         Top             =   1890
         Width           =   795
      End
      Begin VB.TextBox txt1SD 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   -71100
         TabIndex        =   63
         Top             =   1890
         Width           =   795
      End
      Begin VB.ComboBox cmbLotNumber 
         Height          =   315
         Index           =   1
         Left            =   -74040
         TabIndex        =   62
         Text            =   "cmbLotNumber"
         Top             =   1410
         Width           =   1845
      End
      Begin VB.TextBox txtMean 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   -72060
         TabIndex        =   61
         Top             =   1410
         Width           =   795
      End
      Begin VB.TextBox txt1SD 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   -71100
         TabIndex        =   60
         Top             =   1410
         Width           =   795
      End
      Begin VB.ComboBox cmbLotNumber 
         Height          =   315
         Index           =   0
         Left            =   -74040
         TabIndex        =   56
         Text            =   "cmbLotNumber"
         Top             =   930
         Width           =   1845
      End
      Begin VB.CommandButton cmdSaveControls 
         Caption         =   "Save Changes"
         Height          =   705
         Left            =   -70080
         Picture         =   "frmHaemDefinitions.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1470
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txt1SD 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   -71100
         TabIndex        =   52
         Top             =   930
         Width           =   795
      End
      Begin VB.TextBox txtMean 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   -72060
         TabIndex        =   51
         Top             =   930
         Width           =   795
      End
      Begin VB.TextBox tFlagMaleLow 
         Height          =   315
         Left            =   -71460
         TabIndex        =   26
         Top             =   1770
         Width           =   915
      End
      Begin VB.TextBox tFlagFemaleHigh 
         Height          =   315
         Left            =   -72990
         TabIndex        =   25
         Top             =   1170
         Width           =   915
      End
      Begin VB.TextBox tFlagMaleHigh 
         Height          =   315
         Left            =   -71460
         TabIndex        =   24
         Top             =   1170
         Width           =   915
      End
      Begin VB.CheckBox oDelta 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   -72570
         TabIndex        =   23
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox tDelta 
         Height          =   285
         Left            =   -72210
         MaxLength       =   5
         TabIndex        =   21
         Top             =   1530
         Width           =   555
      End
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -72600
         TabIndex        =   20
         Top             =   900
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -72600
         TabIndex        =   19
         Top             =   1470
         Width           =   1215
      End
      Begin VB.TextBox tMaleHigh 
         Height          =   315
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   18
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tFemaleHigh 
         Height          =   315
         Left            =   1830
         MaxLength       =   5
         TabIndex        =   17
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tMaleLow 
         Height          =   315
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   16
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tFemaleLow 
         Height          =   315
         Left            =   1830
         MaxLength       =   5
         TabIndex        =   15
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   3
         Left            =   -73170
         TabIndex        =   14
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   1
         Left            =   -71640
         TabIndex        =   13
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   2
         Left            =   -73170
         TabIndex        =   12
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   0
         Left            =   -71640
         TabIndex        =   11
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tFlagFemaleLow 
         Height          =   315
         Left            =   -72990
         TabIndex        =   10
         Top             =   1770
         Width           =   915
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Back Days Limit"
         Height          =   195
         Left            =   -73560
         TabIndex        =   66
         Top             =   1860
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74520
         TabIndex        =   59
         Top             =   1950
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   -74775
         TabIndex        =   58
         Top             =   1470
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "High"
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
         Height          =   195
         Left            =   -74580
         TabIndex        =   57
         Top             =   990
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Lot Number"
         Height          =   195
         Left            =   -73680
         TabIndex        =   54
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "1 SD"
         Height          =   195
         Left            =   -70920
         TabIndex        =   50
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Mean"
         Height          =   195
         Left            =   -71940
         TabIndex        =   49
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   1
         Left            =   -73740
         TabIndex        =   48
         Top             =   1830
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Index           =   2
         Left            =   -73770
         TabIndex        =   47
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   -71190
         TabIndex        =   46
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   2
         Left            =   -72840
         TabIndex        =   45
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   1
         Left            =   -74010
         TabIndex        =   44
         Top             =   2490
         Width           =   4410
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   -72720
         TabIndex        =   43
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   -73110
         TabIndex        =   42
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -73080
         TabIndex        =   41
         Top             =   1500
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Results outside this range will be marked as implausible."
         Height          =   195
         Index           =   2
         Left            =   -73740
         TabIndex        =   40
         Top             =   2340
         Width           =   3930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   3630
         TabIndex        =   39
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   1980
         TabIndex        =   38
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1260
         TabIndex        =   37
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1290
         TabIndex        =   36
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   0
         Left            =   -73920
         TabIndex        =   35
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Index           =   1
         Left            =   -73950
         TabIndex        =   34
         Top             =   1290
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   0
         Left            =   -71370
         TabIndex        =   33
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   1
         Left            =   -73020
         TabIndex        =   32
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Do not print the result if the sample is :-"
         Height          =   195
         Index           =   1
         Left            =   -74070
         TabIndex        =   31
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "These are the normal range values printed on the report forms."
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   2520
         Width           =   4395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   0
         Left            =   -74190
         TabIndex        =   29
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   -73110
         TabIndex        =   28
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   27
         Top             =   1590
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmHaemDefinitions"
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

200       On Error GoTo FillAges_Error

210       g.Rows = 2
220       g.AddItem ""
230       g.RemoveItem 1

240       If lstParameter = "" Then Exit Sub

250       ReDim FromDays(0 To 0)
260       ReDim ToDays(0 To 0)

          'sql = "Select * from HaemTestDefinitions where " & _
          '      "AnalyteName = '" & lstParameter & "' " & _
          '      "and Hospital = '" & cmbHospital & "' " & _
          '      "Order by cast(AgetoDays as numeric) asc"
270       sql = "Select * from HaemTestDefinitions where " & _
              "AnalyteName = '" & lstParameter & "' " & _
              "Order by cast(AgetoDays as numeric) asc"
280       Set tb = New Recordset
290       RecOpenClient 0, tb, sql

300       ReDim FromDays(0 To tb.RecordCount - 1)
310       ReDim ToDays(0 To tb.RecordCount - 1)
320       n = 0
330       Do While Not tb.EOF
340           FromDays(n) = tb!AgeFromDays
350           ToDays(n) = tb!AgeToDays
360           s = Format$(n) & vbTab & _
                  dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
370           g.AddItem s
380           n = n + 1
390           tb.MoveNext
400       Loop

410       If g.Rows > 2 Then
420           g.RemoveItem 1
430       End If

440       g.Col = 0
450       g.row = 1
460       g.CellBackColor = vbYellow
470       g.CellForeColor = vbBlue

480       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "fHaemDefinitions", "FillAges", intEL, strES, sql


End Sub

Private Sub FillControls()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer

520       On Error GoTo FillControls_Error

530       For n = 0 To 2
540           If Trim$(cmbLotNumber(n)) <> "" Then
550               sql = "SELECT TOP 1 " & _
                      "COALESCE(Mean,1) AS Mean, " & _
                      "COALESCE(SD1,0.1) AS SD1 " & _
                      "FROM HaemControlDefinitions " & _
                      "WHERE LotNumber = '" & cmbLotNumber(n) & "' " & _
                      "AND LNH = '" & Choose(n + 1, "H", "N", "L") & "' " & _
                      "AND Analyte = '" & lstParameter & "' " & _
                      "ORDER BY DateEntered DESC"
560               Set tb = New Recordset
570               RecOpenServer 0, tb, sql
580               If Not tb.EOF Then
590                   txtMean(n) = tb!mean
600                   txt1SD(n) = tb!SD1
610               Else
620                   txtMean(n) = ""
630                   txt1SD(n) = ""
640               End If
650           End If
660       Next

670       cmdSaveControls.Visible = False

680       Exit Sub

FillControls_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "fHaemDefinitions", "FillControls", intEL, strES, sql


End Sub

Private Sub FillLotNumbers()

          Dim sql As String
          Dim tb As Recordset

720       On Error GoTo FillLotNumbers_Error

730       cmbLotNumber(0).Clear
740       cmbLotNumber(1).Clear
750       cmbLotNumber(2).Clear

760       sql = "SELECT LotNumber, LNH " & _
              "FROM HaemControlDefinitions " & _
              "GROUP BY LotNumber, LNH "    ' & _
              "Order by DateEntered"
770       Set tb = New Recordset
780       RecOpenServer 0, tb, sql
790       Do While Not tb.EOF
800           Select Case tb!LNH
                  Case "H": cmbLotNumber(0).AddItem tb!LotNumber
810               Case "N": cmbLotNumber(1).AddItem tb!LotNumber
820               Case "L": cmbLotNumber(2).AddItem tb!LotNumber
830           End Select
840           tb.MoveNext
850       Loop

860       Exit Sub

FillLotNumbers_Error:

          Dim strES As String
          Dim intEL As Integer

870       intEL = Erl
880       strES = Err.Description
890       LogError "fHaemDefinitions", "FillLotNumbers", intEL, strES, sql


End Sub


Private Sub FillParameters()

          Dim tb As Recordset
          Dim sql As String

900       lstParameter.Clear
          'sql = "Select distinct AnalyteName from HaemTestDefinitions " & _
          '      "Where Hospital = '" & hospname(0)& "' " & _
          '      "order by AnalyteName asc"
910       sql = "Select distinct AnalyteName from HaemTestDefinitions " & _
              "order by AnalyteName asc"
920       Set tb = New Recordset
930       RecOpenClient 0, tb, sql

940       Do While Not tb.EOF
950           lstParameter.AddItem tb!AnalyteName & ""
960           tb.MoveNext
970       Loop

980       If lstParameter.ListCount > 0 Then
990           lstParameter.Selected(0) = True
1000      End If

End Sub

Private Sub FillDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim Filled As Boolean
          Dim AgeNumber As Integer
          Dim Y As Integer

1010      On Error GoTo FillDetails_Error

1020      Filled = False

1030      AgeNumber = -1
1040      g.Col = 0
1050      For Y = 1 To g.Rows - 1
1060          g.row = Y
1070          If g.CellBackColor = vbYellow Then
1080              AgeNumber = Y - 1
1090              Exit For
1100          End If
1110      Next
1120      If AgeNumber = -1 Then
1130          iMsg "Select Age Range", vbCritical
1140          Exit Sub
1150      End If

1160      tTestName = ""
1170      oDelta = 0
1180      tDelta = ""
1190      tDeltaBackDays = ""
1200      lDP = "0"
1210      tPlausibleLow = ""
1220      tPlausibleHigh = ""
1230      tMaleHigh = ""
1240      tFemaleHigh = ""
1250      tMaleLow = ""
1260      tFemaleLow = ""

1270      sql = "Select * from HaemTestDefinitions where " & _
              "AnalyteName = '" & lstParameter & "' " & _
              "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
              "and AgeToDays = '" & ToDays(AgeNumber) & "' "    '& _
              "And Hospital = '" & hospname(0)& "'"
1280      Set tb = New Recordset
1290      RecOpenClient 0, tb, sql
1300      If Not tb.EOF Then
1310          tTestName = lstParameter
1320          oDelta = IIf(tb!DoDelta, 1, 0)
1330          tDelta = tb!DeltaValue
1340          tDeltaBackDays = tb!DeltaDaysBackLimit & ""
1350          If Not IsNull(tb!Printformat) Then
1360              lDP = tb!Printformat
1370          Else
1380              lDP = "1"
1390          End If
1400          tPlausibleLow = IIf(IsNull(tb!PlausibleLow), "0", tb!PlausibleLow)
1410          tPlausibleHigh = IIf(IsNull(tb!PlausibleHigh), "9999", tb!PlausibleHigh)
1420          tMaleHigh = tb!MaleHigh
1430          tFemaleHigh = tb!FemaleHigh
1440          tMaleLow = tb!MaleLow
1450          tFemaleLow = tb!FemaleLow

1460      End If

1470      Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

1480      intEL = Erl
1490      strES = Err.Description
1500      LogError "fHaemDefinitions", "FillDetails", intEL, strES, sql


End Sub

Private Sub SaveDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Integer

1510      On Error GoTo SaveDetails_Error

1520      g.Col = 0
1530      For Y = 1 To g.Rows - 1
1540          g.row = Y
1550          If g.CellBackColor = vbYellow Then

1560              sql = "Select * from HaemTestDefinitions where " & _
                      "AnalyteName = '" & lstParameter & "' " & _
                      "and AgeFromDays = '" & FromDays(Y - 1) & "' " & _
                      "and AgeToDays = '" & ToDays(Y - 1) & "'"
1570              Set tb = New Recordset
1580              RecOpenClient 0, tb, sql
1590              With tb

1600                  If .EOF Then .AddNew
1610                  !AnalyteName = lstParameter
1620                  !DoDelta = oDelta = 1
1630                  !DeltaValue = Val(tDelta)
1640                  !DeltaDaysBackLimit = Val(tDeltaBackDays)
1650                  !Printformat = lDP
1660                  !MaleLow = Val(tMaleLow)
1670                  !MaleHigh = Val(tMaleHigh)
1680                  !FemaleLow = Val(tFemaleLow)
1690                  !FemaleHigh = Val(tFemaleHigh)
1700                  !PlausibleLow = Val(tPlausibleLow)
1710                  !PlausibleHigh = Val(tPlausibleHigh)
1720                  !AgeFromDays = FromDays(Y - 1)
1730                  !AgeToDays = ToDays(Y - 1)
1740                  .Update
1750              End With
1760              Exit For
1770          End If
1780      Next

1790      Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

1800      intEL = Erl
1810      strES = Err.Description
1820      LogError "fHaemDefinitions", "SaveDetails", intEL, strES, sql


End Sub






Private Sub bAmendAgeRange_Click()

1830      If lstParameter = "" Then
1840          iMsg "Select Parameter", vbCritical
1850          Exit Sub
1860      End If

1870      With frmAges
1880          .Analyte = lstParameter
1890          .SampleType = "Haematology"
1900          .Discipline = "Haematology"
1910          .Show 1
1920      End With

1930      FillAges

End Sub

Private Sub bcancel_Click()

1940      Unload Me

End Sub


Private Sub cmbLotNumber_Click(Index As Integer)

1950      FillControls

End Sub

Private Sub cmbLotNumber_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

1960      cmdSaveControls.Visible = True

End Sub


Private Sub cmdSaveControls_Click()

          Dim sql As String
          Dim n As Integer
          Dim LotNos(0 To 2) As String

1970      On Error GoTo cmdSaveControls_Click_Error

1980      For n = 0 To 2
1990          LotNos(n) = cmbLotNumber(n)
2000          If Trim$(cmbLotNumber(n)) <> "" And Trim$(txtMean(n)) <> "" And Trim$(txt1SD(n)) <> "" Then
2010              sql = "INSERT INTO HaemControlDefinitions " & _
                      "(Analyte, LotNumber, Mean, SD1, LNH, DateEntered ) VALUES " & _
                      "('" & lstParameter & "', " & _
                      " '" & cmbLotNumber(n) & "', " & _
                      " '" & Val(txtMean(n)) & "', " & _
                      " '" & Val(txt1SD(n)) & "', " & _
                      " '" & Choose(n + 1, "H", "N", "L") & "', " & _
                      " '" & Format$(Now, "dd/MMM/yyyy HH:nn:ss") & "') "
2020              Cnxn(0).Execute sql
2030          End If
2040      Next

2050      FillLotNumbers
2060      For n = 0 To 2
2070          cmbLotNumber(n) = LotNos(n)
2080      Next

2090      cmdSaveControls.Visible = False

2100      Exit Sub

cmdSaveControls_Click_Error:

          Dim strES As String
          Dim intEL As Integer

2110      intEL = Erl
2120      strES = Err.Description
2130      LogError "fHaemDefinitions", "cmdSaveControls_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

2140      SSTab1.TabVisible(1) = False

2150      FillParameters
2160      FillAges
2170      FillDetails
2180      FillLotNumbers

End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim n As Integer
          Dim ySave As Integer

2190      If g.MouseRow = 0 Then Exit Sub

2200      ySave = g.row

2210      g.Col = 0
2220      For n = 1 To g.Rows - 1
2230          g.row = n
2240          g.CellBackColor = 0
2250          g.CellForeColor = 0
2260      Next

2270      g.row = ySave
2280      g.CellBackColor = vbYellow
2290      g.CellForeColor = vbBlue

2300      FillDetails

End Sub


Private Sub lDP_Click()

2310      lDP = Format$(Val(lDP) + 1)

2320      If Val(lDP) > 3 Then lDP = "0"

2330      SaveDetails

End Sub


Private Sub lstParameter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

2340      FillAges
2350      FillDetails

2360      FillControls

End Sub


Private Sub odelta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim sql As String

2370      On Error GoTo odelta_MouseUp_Error

2380      sql = "Update HaemTestDefinitions " & _
              "set DoDelta = " & IIf(oDelta = 1, 1, 0) & " " & _
              "where AnalyteName = '" & lstParameter & "'"
2390      Cnxn(0).Execute sql

2400      Exit Sub

odelta_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

2410      intEL = Erl
2420      strES = Err.Description
2430      LogError "fHaemDefinitions", "odelta_MouseUp", intEL, strES, sql


End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

2440      FillControls

End Sub

Private Sub tDelta_KeyUp(KeyCode As Integer, Shift As Integer)

2450      SaveDetails

End Sub


Private Sub tDeltaBackDays_KeyUp(KeyCode As Integer, Shift As Integer)
2460      SaveDetails
End Sub

Private Sub tFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

2470      SaveDetails

End Sub

Private Sub tFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

2480      SaveDetails

End Sub

Private Sub tFlagFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

2490      SaveDetails

End Sub


Private Sub tFlagFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

2500      SaveDetails

End Sub


Private Sub tFlagMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

2510      SaveDetails

End Sub


Private Sub tFlagMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

2520      SaveDetails

End Sub


Private Sub tMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

2530      SaveDetails

End Sub

Private Sub tMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

2540      SaveDetails

End Sub

Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

2550      SaveDetails

End Sub

Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

2560      SaveDetails

End Sub

Private Sub txt1SD_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

2570      cmdSaveControls.Visible = True

End Sub


Private Sub txtMean_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

2580      cmdSaveControls.Visible = True

End Sub


