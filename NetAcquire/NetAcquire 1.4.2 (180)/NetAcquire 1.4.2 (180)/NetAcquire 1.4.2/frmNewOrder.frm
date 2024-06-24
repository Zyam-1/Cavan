VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Test Order"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeQuestions 
      Caption         =   "Questions"
      Height          =   4515
      Left            =   698
      TabIndex        =   58
      Top             =   4410
      Visible         =   0   'False
      Width           =   12495
      Begin VB.CommandButton cmdHide 
         Cancel          =   -1  'True
         Caption         =   "Hide"
         Height          =   285
         Left            =   4830
         TabIndex        =   59
         Top             =   4140
         Width           =   2925
      End
      Begin MSFlexGridLib.MSFlexGrid flxQuestion 
         Height          =   3975
         Left            =   30
         TabIndex        =   60
         Top             =   180
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   5
      End
   End
   Begin VB.CommandButton cmdQuestion 
      Appearance      =   0  'Flat
      Caption         =   "&View Questions"
      Height          =   1035
      Left            =   12540
      Picture         =   "frmNewOrder.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      Tag             =   "save"
      Top             =   5880
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelAMSOrder 
      Caption         =   "Cancel AMS Order"
      Height          =   375
      Left            =   3600
      TabIndex        =   56
      Top             =   6960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstHaeTests 
      Height          =   450
      Left            =   12540
      MultiSelect     =   1  'Simple
      TabIndex        =   54
      Top             =   4800
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.ListBox LstHaePanel 
      Height          =   2205
      Left            =   12480
      MultiSelect     =   1  'Simple
      TabIndex        =   53
      Top             =   1980
      Width           =   1300
   End
   Begin VB.CheckBox chkGBottle 
      Caption         =   "Glucose bottle is in use"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   840
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.ListBox lstBloodCodes 
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   7050
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstBloodTests 
      Height          =   4935
      Left            =   5820
      TabIndex        =   48
      Top             =   1980
      Width           =   1305
   End
   Begin VB.ListBox lstExistingBio 
      Height          =   2205
      Left            =   60
      TabIndex        =   46
      Top             =   4380
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CheckBox chkAddOn 
      Caption         =   "Add on Request"
      Height          =   435
      Left            =   3750
      TabIndex        =   45
      Top             =   270
      Width           =   1005
   End
   Begin VB.ListBox lstSweatCodes 
      Height          =   255
      Left            =   9930
      TabIndex        =   43
      Top             =   7080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ListBox lstSweatTests 
      Height          =   2115
      IntegralHeight  =   0   'False
      Left            =   8760
      TabIndex        =   42
      Top             =   4800
      Width           =   1080
   End
   Begin VB.ListBox lstCSFCodes 
      Height          =   255
      Left            =   8730
      TabIndex        =   41
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstImmnoCodes 
      Height          =   255
      Left            =   9930
      TabIndex        =   40
      Top             =   3930
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ListBox lstUrineCodes 
      Height          =   255
      Left            =   7260
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox lstSerumCodes 
      Height          =   255
      Left            =   1650
      TabIndex        =   38
      Top             =   7050
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CheckBox chkUrgent 
      Caption         =   "Urgent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1380
      TabIndex        =   36
      Top             =   1080
      Width           =   1035
   End
   Begin VB.ListBox lstSerumPanel 
      Height          =   4935
      Left            =   330
      MultiSelect     =   1  'Simple
      TabIndex        =   20
      Top             =   1980
      Width           =   1245
   End
   Begin VB.ListBox lstSerumTests 
      Columns         =   4
      Height          =   4935
      Left            =   1650
      MultiSelect     =   1  'Simple
      TabIndex        =   19
      Top             =   1980
      Width           =   4155
   End
   Begin VB.ListBox lstUrineTests 
      Height          =   3960
      Left            =   7260
      MultiSelect     =   1  'Simple
      TabIndex        =   18
      Top             =   2970
      Width           =   1425
   End
   Begin VB.ListBox lstUrinePanel 
      Height          =   645
      Left            =   7260
      MultiSelect     =   1  'Simple
      TabIndex        =   17
      Top             =   1980
      Width           =   1425
   End
   Begin VB.ListBox lstCSFTests 
      Height          =   2475
      IntegralHeight  =   0   'False
      Left            =   8760
      MultiSelect     =   1  'Simple
      TabIndex        =   16
      Top             =   1980
      Width           =   1080
   End
   Begin VB.ListBox lstCoag 
      Height          =   2115
      IntegralHeight  =   0   'False
      Left            =   11190
      MultiSelect     =   1  'Simple
      TabIndex        =   15
      Top             =   4800
      Width           =   1185
   End
   Begin VB.ListBox lstImmunoTests 
      Height          =   2025
      IntegralHeight  =   0   'False
      Left            =   9930
      MultiSelect     =   1  'Simple
      TabIndex        =   14
      Top             =   1980
      Width           =   1185
   End
   Begin VB.ListBox lstHaem 
      Height          =   2130
      IntegralHeight  =   0   'False
      ItemData        =   "frmNewOrder.frx":0639
      Left            =   11190
      List            =   "frmNewOrder.frx":063B
      MultiSelect     =   1  'Simple
      TabIndex        =   13
      Top             =   1980
      Width           =   1185
   End
   Begin VB.ListBox lstCoagPanel 
      Height          =   2115
      IntegralHeight  =   0   'False
      Left            =   9930
      MultiSelect     =   1  'Simple
      TabIndex        =   12
      Top             =   4800
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Names"
      Height          =   705
      Left            =   11250
      TabIndex        =   9
      Top             =   240
      Width           =   1155
      Begin VB.OptionButton optShort 
         Caption         =   "Short"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   450
         Width           =   675
      End
      Begin VB.OptionButton optLong 
         Caption         =   "Long"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   480
      Top             =   7230
   End
   Begin VB.OptionButton oSorF 
      Caption         =   "Fasting"
      Height          =   225
      Index           =   1
      Left            =   1350
      TabIndex        =   8
      Top             =   780
      Width           =   825
   End
   Begin VB.OptionButton oSorF 
      Alignment       =   1  'Right Justify
      Caption         =   "Random"
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   7
      Top             =   780
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton bsave 
      Appearance      =   0  'Flat
      Caption         =   "&Save Requests"
      Height          =   1035
      Left            =   6990
      Picture         =   "frmNewOrder.frx":063D
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "save"
      Top             =   150
      Width           =   1275
   End
   Begin VB.TextBox txtSampleID 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   390
      MaxLength       =   15
      TabIndex        =   0
      Top             =   330
      Width           =   1605
   End
   Begin VB.TextBox tinput 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Top             =   330
      Width           =   1365
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Caption         =   "&Exit"
      Height          =   1035
      Left            =   9870
      Picture         =   "frmNewOrder.frx":1507
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "cancel"
      Top             =   150
      Width           =   1275
   End
   Begin VB.CommandButton bclear 
      Appearance      =   0  'Flat
      Caption         =   "Cle&ar"
      Height          =   1035
      Left            =   8430
      Picture         =   "frmNewOrder.frx":23D1
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   61
      Top             =   7380
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   12540
      TabIndex        =   55
      Top             =   4440
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Extended IPU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   52
      Top             =   1680
      Width           =   1300
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Haematology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   2
      Left            =   12480
      TabIndex        =   51
      Top             =   1350
      Width           =   1300
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Blood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Left            =   5820
      TabIndex        =   47
      Top             =   1350
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sweat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   44
      Top             =   4500
      Width           =   1080
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   30
      Left            =   9000
      TabIndex        =   37
      Top             =   510
      Width           =   285
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Left            =   345
      TabIndex        =   35
      Top             =   1350
      Width           =   5475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Urine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Left            =   7260
      TabIndex        =   34
      Top             =   1350
      Width           =   1440
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CSF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   0
      Left            =   8760
      TabIndex        =   33
      Top             =   1350
      Width           =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   32
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   1650
      TabIndex        =   31
      Top             =   1680
      Width           =   5490
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   7260
      TabIndex        =   30
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   7260
      TabIndex        =   29
      Top             =   2670
      Width           =   1440
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   28
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Coagulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   0
      Left            =   9930
      TabIndex        =   27
      Top             =   4170
      Width           =   2445
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Immuno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   1
      Left            =   9930
      TabIndex        =   26
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   9930
      TabIndex        =   25
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   2
      Left            =   11190
      TabIndex        =   24
      Top             =   4500
      Width           =   1185
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Haematology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   1
      Left            =   11190
      TabIndex        =   23
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sapphire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   3
      Left            =   11190
      TabIndex        =   22
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   9930
      TabIndex        =   21
      Top             =   4500
      Width           =   1170
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sample Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Test Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmNewOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mFromEdit As Boolean

Private mSampleID As String
Private mChart As String

Private CoagChanged As Boolean
Private HaemChanged As Boolean
Private BioChanged As Boolean
Private ImmChanged As Boolean
Private HaeChanged As Boolean

Private AnalyserID As String

Private Activated As Boolean

Private Type udtBarCode
    BarCodeType As String
    Name As String
    Code As String
End Type
Private BarCodes() As udtBarCode

Private Type udtQBNames
    Short As String
    Long As String
End Type
Private QuickBioNames() As udtQBNames

Private mCancelled As Boolean

Private Const fcsLine_NO = 0
Private Const fcsSr = 1
Private Const fcsRID = 2
Private Const fcsQus = 3
Private Const fcsAns = 4

Private Function CheckCodes() As Boolean

          Dim n As Integer
          Dim Y As Integer

60800     On Error GoTo CheckCodes_Error

60810     CheckCodes = False

60820     For n = 0 To UBound(BarCodes)
60830         With BarCodes(n)
60840             If .Code = tinput Then
60850                 If .BarCodeType = "Control" Then
60860                     Select Case .Name
                          Case "CTLCANCEL": Unload Me: Exit Function
60870                     Case "CTLSAVE":
60880                         bsave = True
60890                         If mFromEdit Then
60900                             mFromEdit = False
60910                             Unload Me
60920                             Exit Function
60930                         End If
60940                     Case "CTLCLEAR": ClearRequests
60950                     Case "CTLRANDOM": oSorF(0) = True
60960                     Case "CTLFASTING": oSorF(1) = True
60970                     Case "CTLR": AnalyserID = "R"
60980                     Case "CTLFBC": lstHaem.Selected(0) = Not lstHaem.Selected(0): HaemChanged = True
60990                     Case "CTLESR": lstHaem.Selected(1) = Not lstHaem.Selected(1): HaemChanged = True
61000                     Case "CTLRETICS": lstHaem.Selected(2) = Not lstHaem.Selected(2): HaemChanged = True
61010                     Case "CTLMONOSPOT": lstHaem.Selected(3) = Not lstHaem.Selected(3): HaemChanged = True
61020                     End Select
61030                     tinput = ""
61040                     tinput.SetFocus
61050                     CheckCodes = True
61060                     Exit Function

61070                 ElseIf .BarCodeType = "Coag" Then

61080                     For Y = 0 To lstCoag.ListCount - 1
61090                         If .Name = UCase$(lstCoag.List(Y)) Then
61100                             lstCoag.Selected(Y) = Not lstCoag.Selected(Y)
61110                             CoagChanged = True
61120                             CheckCodes = True
61130                             Exit Function
61140                         End If
61150                     Next

61160                 ElseIf .BarCodeType = "Immuno" Then

61170                     For Y = 0 To lstImmunoTests.ListCount - 1
61180                         If .Name = UCase$(lstImmunoTests.List(Y)) Then
61190                             lstImmunoTests.Selected(Y) = Not lstImmunoTests.Selected(Y)
61200                             ImmChanged = True
61210                             CheckCodes = True
61220                             Exit Function
61230                         End If
61240                     Next

61250                 End If
61260             End If
61270         End With
61280     Next

61290     Exit Function

CheckCodes_Error:

          Dim strES As String
          Dim intEL As Integer

61300     intEL = Erl
61310     strES = Err.Description
61320     LogError "frmNewOrder", "CheckCodes", intEL, strES

End Function

Private Sub CheckRecentBio()

          Dim n As Integer
          Dim s As String
          Dim PrevDate As String

61330     On Error GoTo CheckRecentBio_Error

61340     For n = 0 To lstSerumTests.ListCount - 1
61350         If lstSerumTests.Selected(n) Then
61360             PrevDate = RecentlyOrdered(lstSerumTests.List(n))
61370             If PrevDate <> "" Then
61380                 s = lstSerumTests.List(n) & " was processed on " & PrevDate & "." & vbCrLf & _
                          "Do you want to order now?"
61390                 If iMsg(s, vbQuestion + vbYesNo) = vbNo Then
61400                     lstSerumTests.Selected(n) = False
61410                 End If
61420             End If
61430         End If
61440     Next

61450     Exit Sub

CheckRecentBio_Error:

          Dim strES As String
          Dim intEL As Integer

61460     intEL = Erl
61470     strES = Err.Description
61480     LogError "frmNewOrder", "CheckRecentBio", intEL, strES

End Sub

Private Sub FillKnownBioOrders()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Found As Boolean
      Dim LongOrShort As String

61490 On Error GoTo FillKnownBioOrders_Error

61500 Found = False

61510 LongOrShort = IIf(optLong, "Long", "Short")

61520 oSorF(0) = True    'Random
61530 chkUrgent.Value = 0
61540 sql = "Select Fasting,COALESCE(Urgent, 0) Urgent from Demographics where " & _
            "SampleID = '" & Val(txtSampleID) & "'"
61550 Set tb = New Recordset
61560 RecOpenClient 0, tb, sql
61570 If Not tb.EOF Then
61580     If Not IsNull(tb!Fasting) Then
61590         If tb!Fasting Then oSorF(1) = True
61600         If tb!Urgent Then chkUrgent.Value = 1
61610     End If
61620 End If

61630 lstExistingBio.Clear

61640 sql = "SELECT T." & LongOrShort & "Name Name, R.Code " & _
            "FROM BioRequests R " & _
            "JOIN BioTestDefinitions T " & _
            "ON R.Code = T.Code " & _
            "WHERE SampleID = '" & Val(txtSampleID) & "' " & _
            "AND InUse = 1 " & _
            "AND R.SampleType = 'S'"
61650 Set tb = New Recordset
61660 RecOpenClient 0, tb, sql
61670 Do While Not tb.EOF
61680     lstExistingBio.AddItem tb!Code & ""
61690     For n = 0 To lstSerumTests.ListCount - 1
61700         If tb!Name = lstSerumTests.List(n) Then
61710             lstSerumTests.Selected(n) = True
61720             Found = True
61730             Exit For
61740         End If
61750     Next
61760     tb.MoveNext
61770 Loop

61780 If Not Found Then
61790     sql = "SELECT T." & LongOrShort & "Name Name " & _
                "FROM BioRequests R JOIN BioTestDefinitions T ON R.Code = T.Code WHERE " & _
                "SampleID = '" & Val(txtSampleID) & "' " & _
                "AND InUse = 1 AND R.SampleType = 'U'"
61800     Set tb = New Recordset
61810     RecOpenClient 0, tb, sql
61820     Do While Not tb.EOF
61830         For n = 0 To lstUrineTests.ListCount - 1
61840             If tb!Name = lstUrineTests.List(n) Then
61850                 lstUrineTests.Selected(n) = True
61860                 Found = True
61870                 Exit For
61880             End If
61890         Next
61900         tb.MoveNext
61910     Loop
61920 End If

61930 If Not Found Then
61940     sql = "SELECT T." & LongOrShort & "Name Name " & _
                "FROM BioRequests R JOIN BioTestDefinitions T ON R.Code = T.Code WHERE " & _
                "SampleID = '" & Val(txtSampleID) & "' " & _
                "AND InUse = 1 and R.SampleType = 'C'"
61950     Set tb = New Recordset
61960     RecOpenClient 0, tb, sql
61970     Do While Not tb.EOF
61980         For n = 0 To lstCSFTests.ListCount - 1
61990             If tb!Name = lstCSFTests.List(n) Then
62000                 lstCSFTests.Selected(n) = True
62010                 Found = True
62020                 Exit For
62030             End If
62040         Next
62050         tb.MoveNext
62060     Loop
62070 End If

      'if biorequests are found
62080 cmdCancelAMSOrder.Visible = Found

62090 sql = "SELECT T." & LongOrShort & "Name Name " & _
            "FROM ImmRequests R JOIN ImmTestDefinitions T ON R.Code = T.Code where " & _
            "SampleID = '" & Val(txtSampleID) & "' " & _
            "AND InUse = 1"
62100 Set tb = New Recordset
62110 RecOpenClient 0, tb, sql
62120 Do While Not tb.EOF
62130     For n = 0 To lstImmunoTests.ListCount - 1
62140         If tb!Name = lstImmunoTests.List(n) Then
62150             lstImmunoTests.Selected(n) = True
62160             Exit For
62170         End If
62180     Next
62190     tb.MoveNext
62200 Loop



62210 Exit Sub

FillKnownBioOrders_Error:

      Dim strES As String
      Dim intEL As Integer

62220 intEL = Erl
62230 strES = Err.Description
62240 LogError "fNewOrder", "FillKnownBioOrders", intEL, strES, sql

End Sub

Private Sub GetChartNumber()

          Dim tb As Recordset
          Dim sql As String

62250     On Error GoTo GetChartNumber_Error

62260     mChart = ""

62270     If Val(txtSampleID) <> 0 Then
62280         sql = "SELECT Chart FROM Demographics WHERE " & _
                    "SampleID = '" & Val(txtSampleID) & "'"
62290         Set tb = New Recordset
62300         RecOpenServer 0, tb, sql
62310         If Not tb.EOF Then
62320             mChart = Trim$(tb!Chart & "")
62330         End If
62340     End If

62350     Exit Sub

GetChartNumber_Error:

          Dim strES As String
          Dim intEL As Integer

62360     intEL = Erl
62370     strES = Err.Description
62380     LogError "frmNewOrder", "GetChartNumber", intEL, strES, sql

End Sub

Private Sub LoadBarCodes()

          Dim sql As String
          Dim tb As Recordset
          Dim intCurrentUpper As Integer
62390     ReDim BarCodes(0 To 0) As udtBarCode
          Dim CodeAdded As Boolean
          Dim LongOrShort As String

62400     On Error GoTo LoadBarCodes_Error

62410     LongOrShort = IIf(optLong, "Long", "Short")

62420     CodeAdded = False
62430     sql = "Select * from BarCodeControl"
62440     Set tb = New Recordset
62450     RecOpenClient 0, tb, sql
62460     With tb
62470         Do While Not .EOF
62480             If Trim$(!Text) <> "" And Trim$(!Code) <> "" Then
62490                 If CodeAdded Then
62500                     intCurrentUpper = UBound(BarCodes)
62510                     ReDim Preserve BarCodes(0 To intCurrentUpper + 1)
62520                     intCurrentUpper = intCurrentUpper + 1
62530                 End If
62540                 BarCodes(intCurrentUpper).Name = Trim$(UCase$(!Text))
62550                 BarCodes(intCurrentUpper).Code = Trim$(UCase$(!Code))
62560                 BarCodes(intCurrentUpper).BarCodeType = "Control"
62570                 CodeAdded = True
62580             End If
62590             .MoveNext
62600         Loop
62610     End With

62620     sql = "Select Distinct TestName, Code from CoagTestDefinitions"
62630     Set tb = New Recordset
62640     RecOpenClient 0, tb, sql
62650     With tb
62660         Do While Not .EOF
62670             If Trim$(!TestName) <> "" And Trim$(!Code) <> "" Then
62680                 If CodeAdded Then
62690                     intCurrentUpper = UBound(BarCodes)
62700                     ReDim Preserve BarCodes(0 To intCurrentUpper + 1)
62710                     intCurrentUpper = intCurrentUpper + 1
62720                 End If
62730                 BarCodes(intCurrentUpper).Name = Trim$(UCase$(!TestName))
62740                 BarCodes(intCurrentUpper).Code = Trim$(UCase$(!Code))
62750                 BarCodes(intCurrentUpper).BarCodeType = "Coag"
62760                 CodeAdded = True
62770             End If
62780             .MoveNext
62790         Loop
62800     End With

62810     sql = "Select Distinct " & LongOrShort & "Name as name, BarCode from BioTestDefinitions where " & _
                "Analyser = '4' " & _
                "and Hospital = '" & HospName(0) & "'"
62820     Set tb = New Recordset
62830     RecOpenClient 0, tb, sql
62840     With tb
62850         Do While Not .EOF
62860             If Trim$(!Name & "") <> "" And Trim$(!BarCode & "") <> "" Then
62870                 If CodeAdded Then
62880                     intCurrentUpper = UBound(BarCodes)
62890                     ReDim Preserve BarCodes(0 To intCurrentUpper + 1)
62900                     intCurrentUpper = intCurrentUpper + 1
62910                 End If
62920                 BarCodes(intCurrentUpper).Name = Trim$(UCase$(!Name))
62930                 BarCodes(intCurrentUpper).Code = Trim$(UCase$(!BarCode))
62940                 BarCodes(intCurrentUpper).BarCodeType = "Immuno"
62950                 CodeAdded = True
62960             End If
62970             .MoveNext
62980         Loop
62990     End With

63000     Exit Sub

LoadBarCodes_Error:

          Dim strES As String
          Dim intEL As Integer

63010     intEL = Erl
63020     strES = Err.Description
63030     LogError "frmNewOrder", "LoadBarCodes", intEL, strES, sql

End Sub

Private Function RecentlyOrdered(ByVal TestName As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim LongOrShort As String
          Dim RetVal As String

63040     On Error GoTo RecentlyOrdered_Error

63050     If optLong.Value = True Then
63060         LongOrShort = "LongName"
63070     Else
63080         LongOrShort = "ShortName"
63090     End If

63100     RetVal = ""

63110     sql = "SELECT TOP 1 " & _
                "CASE WHEN ReRunDays IS NULL THEN '' " & _
              "     WHEN (DATEDIFF(DD, R.RunDate, getdate()) > ReRunDays) THEN '' " & _
              "     ELSE CONVERT(nvarchar(10), R.RunDate, 3) END PassFail " & _
                "FROM BioResults R, BioTestDefinitions T, Demographics D " & _
                "WHERE T.Code = R.Code " & _
                "AND T." & LongOrShort & " = '" & TestName & "' " & _
                "AND R.SampleID = D.SampleID " & _
                "AND D.Chart = '" & mChart & "' " & _
                "ORDER BY R.RunDate DESC"
63120     Set tb = New Recordset
63130     RecOpenServer 0, tb, sql
63140     If Not tb.EOF Then
63150         RetVal = tb!PassFail
63160     End If

63170     RecentlyOrdered = RetVal

63180     Exit Function

RecentlyOrdered_Error:

          Dim strES As String
          Dim intEL As Integer

63190     intEL = Erl
63200     strES = Err.Description
63210     LogError "fNewOrder", "RecentlyOrdered", intEL, strES, sql

End Function

Public Property Let SampleID(ByVal sNewValue As String)

63220     mSampleID = sNewValue

End Property

Private Sub FillCoagList()

          Dim sql As String
          Dim tb As Recordset

63230     On Error GoTo FillCoagList_Error

63240     lstCoag.Clear
63250     lstCoagPanel.Clear

63260     sql = "SELECT DISTINCT TestName FROM CoagTestDefinitions WHERE " & _
                "Hospital = '" & HospName(0) & "' " & _
                "AND InUse = 1"
63270     Set tb = New Recordset
63280     RecOpenClient 0, tb, sql
63290     Do While Not tb.EOF
63300         If Trim$(UCase$(tb!TestName & "")) <> "FIB" Then
63310             lstCoag.AddItem tb!TestName
63320         End If
63330         tb.MoveNext
63340     Loop

63350     sql = "Select distinct PanelName, ListOrder from CoagPanels " & _
                "Order by ListOrder"
63360     Set tb = New Recordset
63370     RecOpenServer 0, tb, sql
63380     Do While Not tb.EOF
63390         lstCoagPanel.AddItem tb!PanelName & ""
63400         tb.MoveNext
63410     Loop

63420     Exit Sub

FillCoagList_Error:

          Dim strES As String
          Dim intEL As Integer

63430     intEL = Erl
63440     strES = Err.Description
63450     LogError "frmNewOrder", "FillCoagList", intEL, strES, sql

End Sub

Public Property Let FromEdit(ByVal bFromEdit As Boolean)

63460     mFromEdit = bFromEdit

End Property

Public Property Let Chart(ByVal strNewValue As String)

63470     mChart = Trim$(strNewValue)

End Property


Function CheckCSF() As Boolean

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim LongOrShort As String

63480     On Error GoTo CheckCSF_Error

63490     LongOrShort = IIf(optLong, "Long", "Short")

63500     CheckCSF = False
63510     sql = "Select " & LongOrShort & "Name as Name from BioTestDefinitions where " & _
                "SampleType = 'C' " & _
                "and BarCode =  '" & tinput & "'"
63520     Set tb = New Recordset
63530     RecOpenServer 0, tb, sql
63540     If Not tb.EOF Then
63550         CheckCSF = True
63560         For Y = 0 To lstCSFTests.ListCount - 1
63570             If lstCSFTests.List(Y) = tb!Name Then
63580                 lstCSFTests.Selected(Y) = Not lstCSFTests.Selected(Y)
63590                 BioChanged = True
63600                 Exit For
63610             End If
63620         Next
63630     End If

63640     Exit Function

CheckCSF_Error:

          Dim strES As String
          Dim intEL As Integer

63650     intEL = Erl
63660     strES = Err.Description
63670     LogError "frmNewOrder", "CheckCSF", intEL, strES, sql

End Function

Function CheckSerum() As Boolean

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim LongOrShort As String

63680     On Error GoTo CheckSerum_Error

63690     LongOrShort = IIf(optLong, "Long", "Short")

63700     CheckSerum = False
63710     sql = "Select " & LongOrShort & "Name as Name from BioTestDefinitions where " & _
                "SampleType = 'S' " & _
                "and BarCode = '" & tinput & "'"
63720     Set tb = New Recordset
63730     RecOpenServer 0, tb, sql
63740     If Not tb.EOF Then
63750         CheckSerum = True
63760         For Y = 0 To lstSerumTests.ListCount - 1
63770             If lstSerumTests.List(Y) = tb!Name Then
63780                 lstSerumTests.Selected(Y) = Not lstSerumTests.Selected(Y)
63790                 Exit For
63800             End If
63810         Next
63820         BioChanged = True
63830     End If

63840     Exit Function

CheckSerum_Error:

          Dim strES As String
          Dim intEL As Integer

63850     intEL = Erl
63860     strES = Err.Description
63870     LogError "frmNewOrder", "CheckSerum", intEL, strES, sql

End Function

Function CheckBlood() As Boolean

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim LongOrShort As String

63880     LongOrShort = IIf(optLong, "Long", "Short")

63890     CheckBlood = False
63900     sql = "Select " & LongOrShort & "Name as Name from BioTestDefinitions where " & _
                "SampleType = 'B' " & _
                "and BarCode = '" & tinput & "'"
63910     Set tb = New Recordset
63920     RecOpenServer 0, tb, sql
63930     If Not tb.EOF Then
63940         CheckBlood = True
63950         For Y = 0 To lstBloodTests.ListCount - 1
63960             If lstBloodTests.List(Y) = tb!Name Then
63970                 lstBloodTests.Selected(Y) = Not lstBloodTests.Selected(Y)
63980                 Exit For
63990             End If
64000         Next
64010         BioChanged = True
64020     End If

End Function


Function CheckSerumPanel() As Boolean

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String

64030     On Error GoTo CheckSerumPanel_Error

64040     CheckSerumPanel = False

64050     sql = "Select * from Panels where " & _
                "BarCode = '" & tinput & "' " & _
                "and PanelType = 'S' "
64060     Set tb = New Recordset
64070     RecOpenServer 0, tb, sql

64080     Do While Not tb.EOF
64090         CheckSerumPanel = True
64100         For Y = 0 To lstSerumPanel.ListCount - 1
64110             If lstSerumPanel.List(Y) = tb!PanelName Then
64120                 lstSerumPanel.Selected(Y) = Not lstSerumPanel.Selected(Y)
64130                 Exit For
64140             End If
64150         Next
64160         tb.MoveNext
64170     Loop

64180     If CheckSerumPanel Then
64190         BioChanged = True
64200     End If

64210     Exit Function

CheckSerumPanel_Error:

          Dim strES As String
          Dim intEL As Integer

64220     intEL = Erl
64230     strES = Err.Description
64240     LogError "frmNewOrder", "CheckSerumPanel", intEL, strES, sql

End Function
Function CheckUrine() As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Integer
          Dim LongOrShort As String

64250     On Error GoTo CheckUrine_Error

64260     LongOrShort = IIf(optLong, "Long", "Short")

64270     CheckUrine = False
64280     sql = "Select " & LongOrShort & "Name as Name from BioTestDefinitions where " & _
                "SampleType = 'U' " & _
                "and BarCode = '" & tinput & "'"
64290     Set tb = New Recordset
64300     RecOpenServer 0, tb, sql
64310     If Not tb.EOF Then
64320         CheckUrine = True
64330         For Y = 0 To lstUrineTests.ListCount - 1
64340             If lstUrineTests.List(Y) = tb!Name Then
64350                 lstUrineTests.Selected(Y) = Not lstUrineTests.Selected(Y)
64360                 Exit For
64370             End If
64380         Next
64390         BioChanged = True
64400     End If

64410     Exit Function

CheckUrine_Error:

          Dim strES As String
          Dim intEL As Integer

64420     intEL = Erl
64430     strES = Err.Description
64440     LogError "frmNewOrder", "CheckUrine", intEL, strES, sql

End Function

Function CheckUrinePanel() As Boolean

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String

64450     On Error GoTo CheckUrinePanel_Error

64460     CheckUrinePanel = False
64470     sql = "Select * from Panels where " & _
                "BarCode = '" & tinput & "' " & _
                "and PanelType = 'U' "
64480     Set tb = New Recordset
64490     RecOpenServer 0, tb, sql
64500     Do While Not tb.EOF
64510         CheckUrinePanel = True
64520         For Y = 0 To lstUrinePanel.ListCount - 1
64530             If lstUrinePanel.List(Y) = tb!PanelName Then
64540                 lstUrinePanel.Selected(Y) = Not lstUrinePanel.Selected(Y)
64550                 Exit For
64560             End If
64570         Next
64580         tb.MoveNext
64590     Loop
64600     If CheckUrinePanel Then
64610         BioChanged = True
64620     End If

64630     Exit Function

CheckUrinePanel_Error:

          Dim strES As String
          Dim intEL As Integer

64640     intEL = Erl
64650     strES = Err.Description
64660     LogError "frmNewOrder", "CheckUrinePanel", intEL, strES, sql

End Function
Sub ClearRequests()

          Dim n As Integer

64670     On Error GoTo ClearRequests_Error

64680     chkUrgent.Value = 0

64690     For n = 0 To lstSerumPanel.ListCount - 1
64700         lstSerumPanel.Selected(n) = False
64710     Next

64720     For n = 0 To lstSerumTests.ListCount - 1
64730         lstSerumTests.Selected(n) = False
64740     Next

64750     For n = 0 To lstUrinePanel.ListCount - 1
64760         lstUrinePanel.Selected(n) = False
64770     Next

64780     For n = 0 To lstUrineTests.ListCount - 1
64790         lstUrineTests.Selected(n) = False
64800     Next

64810     For n = 0 To lstCSFTests.ListCount - 1
64820         lstCSFTests.Selected(n) = False
64830     Next

64840     For n = 0 To lstImmunoTests.ListCount - 1
64850         lstImmunoTests.Selected(n) = False
64860     Next

64870     For n = 0 To lstCoag.ListCount - 1
64880         lstCoag.Selected(n) = False
64890     Next

64900     For n = 0 To lstCoagPanel.ListCount - 1
64910         lstCoagPanel.Selected(n) = False
64920     Next

64930     For n = 0 To lstHaem.ListCount - 1
64940         lstHaem.Selected(n) = False
64950     Next

64960     For n = 0 To lstBloodTests.ListCount - 1
64970         lstBloodTests.Selected(n) = False
64980     Next

64990     Exit Sub

ClearRequests_Error:

          Dim strES As String
          Dim intEL As Integer

65000     intEL = Erl
65010     strES = Err.Description
65020     LogError "frmNewOrder", "ClearRequests", intEL, strES


End Sub

Private Sub FillLists()

          Dim tb As Recordset
          Dim sql As String
          Dim LongOrShort As String

65030     On Error GoTo FillLists_Error

65040     FillCoagList

65050     LongOrShort = IIf(optLong, "Long", "Short")

65060     lstSerumPanel.Clear
65070     lstSerumTests.Clear
65080     lstBloodTests.Clear
65090     lstUrinePanel.Clear
65100     lstUrineTests.Clear
65110     lstCSFTests.Clear
65120     lstSerumCodes.Clear
65130     lstBloodCodes.Clear
65140     lstUrineCodes.Clear



65150     With lstHaem
65160         .Clear
      '150           .AddItem "FBC"
65170         .AddItem "ESR"
      '170           .AddItem "Retics"
65180         .AddItem "Monospot"
65190         .AddItem "Malaria"
65200         .AddItem "Sickledex"
              '160     .AddItem "CD3/4/8"
65210         If sysOptAlwaysRequestFBC(0) Then
65220             .Selected(0) = True
65230         End If
65240     End With

65250     sql = "Select distinct PanelName, ListOrder from Panels where " & _
                "PanelType = 'S' " & _
                "and Hospital = '" & HospName(0) & "' " & _
                "Order by ListOrder"
65260     Set tb = New Recordset
65270     RecOpenServer 0, tb, sql
65280     Do While Not tb.EOF
65290         lstSerumPanel.AddItem tb!PanelName
65300         lstSerumPanel.AddItem ""
65310         tb.MoveNext
65320     Loop

65330     sql = "Select distinct PanelName, ListOrder from Panels where " & _
                "PanelType = 'U' " & _
                "and Hospital = '" & HospName(0) & "' " & _
                "Order by ListOrder"
65340     Set tb = New Recordset
65350     RecOpenServer 0, tb, sql
65360     Do While Not tb.EOF
65370         lstUrinePanel.AddItem tb!PanelName
65380         tb.MoveNext
65390     Loop

65400     sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                "PrintPriority FROM BioTestDefinitions WHERE " & _
                "COALESCE(Analyser, '') <> '' " & _
                "AND SampleType = 'S' " & _
                "AND ISNULL(KnownToAnalyser,0) = 1 " & _
                "AND InUse = 1 ORDER BY PrintPriority"
65410     Set tb = Cnxn(0).Execute(sql)
65420     Do While Not tb.EOF
65430         lstSerumTests.AddItem tb!Name
65440         lstSerumCodes.AddItem tb!Code
65450         tb.MoveNext
65460     Loop

65470     sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                "PrintPriority FROM BioTestDefinitions WHERE " & _
                "COALESCE(Analyser, '') <> '' " & _
                "AND SampleType = 'B' " & _
                "AND KnownToAnalyser = 1 " & _
                "ORDER BY PrintPriority"
65480     Set tb = Cnxn(0).Execute(sql)
65490     Do While Not tb.EOF
65500         lstBloodTests.AddItem tb!Name
65510         lstBloodCodes.AddItem tb!Code
65520         tb.MoveNext
65530     Loop

10        sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                "PrintPriority from BioTestDefinitions WHERE " & _
                "COALESCE(Analyser, '')  <> '' " & _
                "AND SampleType = 'U' " & _
                "AND KnownToAnalyser = 1 " & _
                "ORDER BY PrintPriority"
20        Set tb = Cnxn(0).Execute(sql)
30        Do While Not tb.EOF
40            lstUrineTests.AddItem tb!Name
50            lstUrineCodes.AddItem tb!Code
60            tb.MoveNext
70        Loop

80        sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                "PrintPriority from BioTestDefinitions WHERE " & _
                "COALESCE(Analyser, '') <> '' " & _
                "AND SampleType = 'C' " & _
                "AND KnownToAnalyser = 1 " & _
                "ORDER BY PrintPriority"
90        Set tb = Cnxn(0).Execute(sql)
100       Do While Not tb.EOF
110           lstCSFTests.AddItem tb!Name
120           lstCSFCodes.AddItem tb!Code
130           tb.MoveNext
140       Loop

150       lstSweatCodes.Clear
160       lstSweatTests.Clear
170       sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                "PrintPriority from BioTestDefinitions WHERE " & _
                "COALESCE(Analyser, '') <> '' " & _
                "AND SampleType = 'SW' AND KnownToAnalyser = 1 " & _
                "ORDER BY PrintPriority"
180       Set tb = Cnxn(0).Execute(sql)
190       Do While Not tb.EOF
200           lstSweatTests.AddItem tb!Name
210           lstSweatCodes.AddItem tb!Code
220           tb.MoveNext
230       Loop

240       lstImmunoTests.Clear
250       sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                "PrintPriority FROM ImmTestDefinitions " & _
                "WHERE COALESCE(Analyser, '') <>'' AND KnownToAnalyser = 1 " & _
                "ORDER BY PrintPriority"
260       Set tb = Cnxn(0).Execute(sql)
270       Do While Not tb.EOF
280           lstImmunoTests.AddItem tb!Name
290           lstImmnoCodes.AddItem tb!Code
300           tb.MoveNext
310       Loop

320       With LstHaePanel
330           .Clear
340           sql = "Select distinct PanelName, ListOrder from HaePanels where " & _
                  " Hospital = '" & HospName(0) & "' " & _
                    "Order by ListOrder"
350           Set tb = New Recordset
360           RecOpenServer 0, tb, sql

370           Do While Not tb.EOF
380               .AddItem tb!PanelName
                  '            .AddItem ""
390               tb.MoveNext
400           Loop
410       End With

420       With lstHaeTests
430           .Clear
440           sql = "SELECT DISTINCT " & LongOrShort & "Name Name, Code, " & _
                    "PrintPriority FROM HaemTestDefinitions " & _
                    "WHERE ISNULL(KnownToAnalyser,0) = 1 " & _
                    "ORDER BY PrintPriority"
450           Set tb = Cnxn(0).Execute(sql)
460           Do While Not tb.EOF
                  '            .AddItem tb!Name
470               .AddItem tb!Name
480               tb.MoveNext
490           Loop

500       End With







510       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

520       intEL = Erl
530       strES = Err.Description
540       LogError "frmNewOrder", "FillLists", intEL, strES, sql

End Sub

Private Sub SaveBioResult(ByVal Code As String, ByVal Result As String)

      Dim SampleDate As String
      Dim ReceiveDate As String
      Dim sql        As String
      Dim tb         As Recordset
      Dim l_Hours As Integer

550   On Error GoTo SaveBioResult_Error

560   sql = "SELECT * FROM Demographics WHERE SampleID = " & txtSampleID
570   Set tb = New Recordset
580   RecOpenServer 0, tb, sql
590   If Not tb.EOF Then

600       SampleDate = ConvertNull(tb!SampleDate, Now)
610       ReceiveDate = ConvertNull(tb!RecDate, Now)
      '          +++Junaid
      '          If DateDiff("H", SampleDate, ReceiveDate) > 8 Then
620       If Code = "1102" Then
630         l_Hours = 6
640       Else
650         l_Hours = 24
660       End If
670       If DateDiff("H", SampleDate, ReceiveDate) >= l_Hours Then
      '          ---Junaid
680           sql = "IF EXISTS(SELECT * FROM BioResults " & _
                    "          WHERE SampleID = @sampleid0 " & _
                    "          AND Code = '@Code1' ) " & _
                    "UPDATE [dbo].[BioResults] SET " & _
                    "[sampleid] = @sampleid0 , " & _
                    "[Code] = '@Code1', " & _
                    "[result] = '@result2', " & _
                    "[valid] = @valid3, " & _
                    "[printed] = @printed4, " & _
                    "[RunTime] = @RunTime5, " & _
                    "[RunDate] = @RunDate6, " & _
                    "[Units] = '@Units9', " & _
                    "[SampleType] = '@SampleType10', " & _
                    "[Analyser] = '@Analyser11', " & _
                    "[Faxed] = '@Faxed12', " & _
                    "[Healthlink] = '@Healthlink18'" & _
                    "WHERE SampleID = @sampleid0 AND Code = '@Code1' " & _
                    "ELSE " & _
                    "  INSERT INTO BioResults " & _
                    "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
                    "   Units, SampleType, Analyser, Faxed, " & _
                    "   Healthlink) VALUES " & _
                    "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
                    "  '@Units9', '@SampleType10', '@Analyser11', " & _
                    "  @Faxed12, @Healthlink18) "

690           sql = Replace(sql, "@sampleid0", txtSampleID)
700           sql = Replace(sql, "@Code1", Code)
710           sql = Replace(sql, "@result2", Result)
720           sql = Replace(sql, "@valid3", 1)
730           sql = Replace(sql, "@printed4", 0)
740           sql = Replace(sql, "@RunTime5", Format$(Now, "'dd/mmm/yyyy hh:mm:ss'"))
750           sql = Replace(sql, "@RunDate6", Format$(Now, "'dd/mmm/yyyy'"))
760           sql = Replace(sql, "@Units9", "mMol/L")
770           sql = Replace(sql, "@SampleType10", "S")
780           sql = Replace(sql, "@Analyser11", "Manual")
790           sql = Replace(sql, "@Faxed12", 0)
800           sql = Replace(sql, "@Healthlink18", 0)

810           Cnxn(0).Execute sql

820       End If



830   End If

840   Exit Sub
SaveBioResult_Error:

850   LogError "frmNewOrder", "SaveBioResult", Erl, Err.Description, sql


End Sub

Private Sub SaveBio()

      Dim n          As Integer
      Dim Code       As String
      Dim sql        As String
      Dim tb         As Recordset
      Dim ExistingCodeList As String
      Dim CurrentCodeList As String
      Dim e          As Integer
      Dim Found      As Boolean
      Dim GlucoseCode As String

860   On Error GoTo SaveBio_Error

870   ExistingCodeList = ""
880   For n = 0 To lstExistingBio.ListCount - 1
890       ExistingCodeList = ExistingCodeList & "'" & lstExistingBio.List(n) & "',"
900   Next

910   If Len(ExistingCodeList) > 0 Then
920       ExistingCodeList = Left(ExistingCodeList, Len(ExistingCodeList) - 1)

930       Cnxn(0).Execute ("DELETE FROM BioRequests WHERE " & _
                           "SampleID = '" & txtSampleID & "' " & _
                           "AND Code NOT IN (" & ExistingCodeList & ") " & _
                           "AND Programmed = 0")
940   End If

950   CurrentCodeList = "'x',"
960   For n = 0 To lstSerumTests.ListCount - 1
970       If lstSerumTests.Selected(n) Then
980           CurrentCodeList = CurrentCodeList & "'" & lstSerumCodes.List(n) & "',"
990       End If
1000  Next
1010  If Len(CurrentCodeList) > 0 Then
1020      CurrentCodeList = Left(CurrentCodeList, Len(CurrentCodeList) - 1)
1030      Cnxn(0).Execute ("DELETE FROM BioRequests WHERE " & _
                           "SampleID = '" & txtSampleID & "' " & _
                           "AND Code NOT IN (" & CurrentCodeList & ") " & _
                           "AND Programmed = 0")
1040  End If

      '' MASOOD 22-01-2015
      'sql = "SELECT Contents FROM Options WHERE Description Like 'GlucoseCode%'"
      'Set tb = New Recordset
      'RecOpenServer 0, tb, sql
      'If tb.EOF Then
      '    GlucoseCode = ""
      'Else
      '    While Not tb.EOF
      '        If Code = tb!Contents & "" Then
      '            GlucoseCode = tb!Contents
      '        End If
      '        tb.MoveNext
      '    Wend
      'End If
      '' MASOOD 22-01-2015

1050  For n = 0 To lstSerumTests.ListCount - 1
1060      If lstSerumTests.Selected(n) Then
1070          Code = lstSerumCodes.List(n)
1080          Found = False
1090          For e = 0 To lstExistingBio.ListCount - 1
1100              If lstExistingBio.List(e) = Code Then
1110                  Found = True
1120                  Exit For
1130              End If
1140          Next
1150          If Not Found Then
1160              UpDateRequests "Bio", Code, "S", IIf(Code = FndOptionSettingGlucose(Code), chkGBottle.Value, 0)
1170              If Code = GetOptionSetting("BioCodeForPotassium", "1102") Then
1180                  SaveBioResult Code, "Old"
1190                  If mFromEdit Then frmEditAll.LoadBiochemistry
1200              End If
                  '+++ Junaid
1210              If Code = GetOptionSetting("BioCodeForPhosphate", "1011") Then
1220                  SaveBioResult Code, "Old"
1230                  If mFromEdit Then frmEditAll.LoadBiochemistry
1240              End If
                  '--- Junaid
1250          End If
1260      End If
1270  Next

1280  For n = 0 To lstSweatTests.ListCount - 1
1290      If lstSweatTests.Selected(n) Then
1300          Code = lstSweatCodes.List(n)
1310          UpDateRequests "Bio", Code, "SW"
1320      End If
1330  Next

1340  For n = 0 To lstBloodTests.ListCount - 1
1350      If lstBloodTests.Selected(n) Then
1360          Code = lstBloodCodes.List(n)
1370          UpDateRequests "Bio", Code, "B"
1380      End If
1390  Next

1400  For n = 0 To lstUrineTests.ListCount - 1
1410      If lstUrineTests.Selected(n) Then
1420          Code = lstUrineCodes.List(n)
1430          UpDateRequests "Bio", Code, "U"
1440      End If
1450  Next

1460  For n = 0 To lstCSFTests.ListCount - 1
1470      If lstCSFTests.Selected(n) Then
1480          Code = lstCSFCodes.List(n)
1490          UpDateRequests "Bio", Code, "C"
1500      End If
1510  Next


1520  sql = "SELECT * FROM demographics WHERE " & _
            "SampleID = '" & txtSampleID & "'"
1530  Set tb = New Recordset
1540  RecOpenClient 0, tb, sql
1550  If tb.EOF Then
1560      tb.AddNew
1570      tb!Rundate = Format$(Now, "dd/mmm/yyyy")
1580      tb!SampleID = txtSampleID
1590      tb!FAXed = 0
1600      tb!RooH = 0
1610  End If
1620  If chkUrgent.Value = 1 Then
1630      tb!Urgent = 1
1640  Else
1650      tb!Urgent = 0
1660  End If
1670  tb!Fasting = IIf(oSorF(1), 1, 0)
1680  tb.Update

1690  Exit Sub

SaveBio_Error:

      Dim strES      As String
      Dim intEL      As Integer

1700  intEL = Erl
1710  strES = Err.Description
1720  LogError "fNewOrder", "SaveBio", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : DeleteOldEnteries
' Author    : Masood
' Date      : 29/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub DeleteOldEnteries(Discipline As String)
1730      On Error GoTo DeleteOldEnteries_Error

          Dim sql As String

1740      sql = "Delete from  " & Discipline & "Requests "
1750      sql = sql & " WHERE SAMPLEID ='" & txtSampleID & "'"

1760      Cnxn(0).Execute sql


1770      Exit Sub


DeleteOldEnteries_Error:

          Dim strES As String
          Dim intEL As Integer

1780      intEL = Erl
1790      strES = Err.Description
1800      LogError "frmNewOrder", "DeleteOldEnteries", intEL, strES
End Sub




Private Sub SaveImm()

          Dim n As Integer
          Dim Code As String
          Dim sql As String
          Dim tb As Recordset

1810      On Error GoTo SaveImm_Error

1820      Cnxn(0).Execute ("DELETE FROM ImmRequests WHERE " & _
                           "SampleID = '" & txtSampleID & "' " & _
                           "AND Programmed = 0")

1830      For n = 0 To lstImmunoTests.ListCount - 1
1840          If lstImmunoTests.Selected(n) Then
1850              Code = lstImmnoCodes.List(n)
1860              UpDateRequests "Imm", Code, "S"
1870          End If
1880      Next

1890      sql = "SELECT * FROM demographics WHERE " & _
                "SampleID = '" & txtSampleID & "'"
1900      Set tb = New Recordset
1910      RecOpenClient 0, tb, sql
1920      If tb.EOF Then
1930          tb.AddNew
1940          tb!Rundate = Format$(Now, "dd/mmm/yyyy")
1950          tb!SampleID = txtSampleID
1960          tb!FAXed = 0
1970          tb!RooH = 0
1980      End If
1990      If chkUrgent.Value = 1 Then
2000          tb!Urgent = 1
2010      Else
2020          tb!Urgent = 0
2030      End If
2040      tb!Fasting = IIf(oSorF(1), 1, 0)
2050      tb.Update

2060      Exit Sub

SaveImm_Error:

          Dim strES As String
          Dim intEL As Integer

2070      intEL = Erl
2080      strES = Err.Description
2090      LogError "fNewOrder", "SaveImm", intEL, strES, sql

End Sub

Private Sub SaveCoag()

      Dim sql As String
      Dim n As Integer
      Dim TestCode As String

2100  On Error GoTo SaveCoag_Error

2110  sql = "UPDATE demographics " & _
            "SET Urgent = " & IIf(chkUrgent.Value = 1, 1, 0) & " " & _
            "WHERE SampleID = '" & txtSampleID & "'"
2120  Cnxn(0).Execute sql
2130  sql = "Delete from CoagRequests where " & _
            "SampleID = '" & mSampleID & "'"
2140  Cnxn(0).Execute sql

2150  For n = 0 To lstCoag.ListCount - 1
2160      If lstCoag.Selected(n) Then
2170          TestCode = CoagCodeForTestName(lstCoag.List(n))
2180          sql = "Insert into CoagRequests " & _
                    "(SampleID, Code) VALUES " & _
                    "('" & mSampleID & "', " & _
                    "'" & TestCode & "')"
2190          Cnxn(0).Execute sql
2200      End If
2210  Next

2220  For n = 0 To lstCoag.ListCount - 1
2230      lstCoag.Selected(n) = False
2240  Next

2250  Exit Sub

SaveCoag_Error:

      Dim strES As String
      Dim intEL As Integer

2260  intEL = Erl
2270  strES = Err.Description
2280  LogError "fNewOrder", "SaveCoag", intEL, strES, sql

End Sub

Private Sub SetDisplay(ByVal LongOrShort As String)

          Dim sql As String
          Dim LOrS As String

2290      On Error GoTo SetDisplay_Error

2300      LOrS = GetOptionSetting("LongOrShortBioNames", "Short")

2310      If LOrS <> LongOrShort Then
2320          If iMsg("Do you want to reset the Default Display to " & LongOrShort & " Names?", vbQuestion + vbYesNo) = vbYes Then
2330              SaveOptionSetting "LongOrShortBioNames", LongOrShort
2340          End If
2350      End If

2360      Exit Sub

SetDisplay_Error:

          Dim strES As String
          Dim intEL As Integer

2370      intEL = Erl
2380      strES = Err.Description
2390      LogError "frmNewOrder", "SetDisplay", intEL, strES, sql

End Sub

Private Sub UpDateRequests(ByVal Discipline As String, _
                           ByVal Code As String, _
                           ByVal SampleType As String, Optional Gbottle As Integer)

          Dim sql As String

2400      On Error GoTo UpDateRequests_Error

2410      sql = "INSERT INTO " & Discipline & "Requests " & _
                "(SampleID, Code, DateTime, SampleType, Programmed, AddOn, AnalyserID,Gbottle) " & _
                "SELECT DISTINCT '" & txtSampleID & "', " & _
              "       '" & Code & "', getdate(), " & _
              "       '" & SampleType & "', '0', '" & chkAddOn.Value & "', " & _
              "       Analyser ," & Gbottle & "  FROM " & Discipline & "TestDefinitions " & _
              "        " & _
              " WHERE Code = '" & Code & "' " & _
              " AND InUse = 1"
2420      Cnxn(0).Execute sql

2430      Exit Sub

UpDateRequests_Error:

          Dim strES As String
          Dim intEL As Integer

2440      intEL = Erl
2450      strES = Err.Description
2460      LogError "frmNewOrder", "UpDateRequests", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

2470      mCancelled = True
2480      Unload Me

End Sub

Private Sub bclear_Click()

2490      pBar = 0
2500      ClearRequests

End Sub

Private Sub bSave_Click()

2510      On Error GoTo bSave_Click_Error

2520      pBar = 0

2530      If Trim$(txtSampleID) = "" Then
2540          iMsg "Sample Number Required.", vbCritical
2550          Exit Sub
2560      End If

2570      If CoagChanged Then SaveCoag
2580      If HaemChanged Then SaveHaem
2590      If ImmChanged Then SaveImm

2600      If HaeChanged Then SaveHae

2610      If BioChanged Then
2620          If mChart <> "" Then
2630              CheckRecentBio
2640          End If
2650          SaveBio
2660      End If


2670      ClearRequests

2680      txtSampleID = Format$(Val(txtSampleID) + 1)
2690      txtSampleID.SelStart = 0
2700      txtSampleID.SelLength = Len(txtSampleID)
2710      txtSampleID.SetFocus

2720      If mFromEdit Then
2730          mFromEdit = False
2740          Unload Me
2750      End If

2760      Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

2770      intEL = Erl
2780      strES = Err.Description
2790      LogError "frmNewOrder", "bSave_Click", intEL, strES

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveHae
' Author    : Masood
' Date      : 29/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveHae()
2800      On Error GoTo SaveHae_Error

          Dim n As Integer
          Dim TestName As String
          Dim sql As String
2810      With LstHaePanel
2820          DeleteOldEnteries ("Hae")
2830          For n = 0 To .ListCount - 1
2840              If .Selected(n) Then
2850                  TestName = .List(n)
2860                  UpDateRequestsHae "Hae", TestName, IIf((optLong.Value = True), "Long", "Short")
2870              End If
2880          Next
2890      End With


        Dim tb As New ADODB.Recordset
2900      sql = "SELECT * FROM demographics WHERE " & _
                "SampleID = '" & txtSampleID & "'"
2910      Set tb = New Recordset
2920      RecOpenClient 0, tb, sql
2930      If tb.EOF Then
2940          tb.AddNew
2950          tb!Rundate = Format$(Now, "dd/mmm/yyyy")
2960          tb!SampleID = txtSampleID
2970          tb!FAXed = 0
2980          tb!RooH = 0
2990      End If
3000      If chkUrgent.Value = 1 Then
3010          tb!Urgent = 1
3020      Else
3030          tb!Urgent = 0
3040      End If

3050      tb.Update


3060      Exit Sub


SaveHae_Error:

          Dim strES As String
          Dim intEL As Integer

3070      intEL = Erl
3080      strES = Err.Description
3090      LogError "frmNewOrder", "SaveHae", intEL, strES
End Sub


Private Sub SaveHaem()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim FBCFound As Boolean
          Dim ESRFound As Boolean
          Dim ReticsFound As Boolean
          Dim MonoSpotFound As Boolean
          Dim MalariaFound As Boolean
          Dim SickleFound As Boolean
          Dim CD348Found As Boolean
          Dim OrderString As String
          Dim Young As Boolean

3100      On Error GoTo SaveHaem_Error

3110      Found = False
3120      FBCFound = False
3130      ESRFound = False
3140      ReticsFound = False
3150      MonoSpotFound = False
3160      MalariaFound = False
3170      SickleFound = False
3180      CD348Found = False

3190      For n = 0 To lstHaem.ListCount - 1
3200          If lstHaem.Selected(n) Then
3210              Select Case UCase$(lstHaem.List(n))
                  Case "FBC"
3220                  Found = True
3230                  FBCFound = True
3240              Case "ESR"
3250                  Found = True
3260                  ESRFound = True
3270              Case "RETICS"
3280                  Found = True
3290                  ReticsFound = True
3300              Case "MONOSPOT"
3310                  Found = True
3320                  MonoSpotFound = True
3330              Case "MALARIA"
3340                  Found = True
3350                  MalariaFound = True
3360              Case "SICKLEDEX"
3370                  Found = True
3380                  SickleFound = True
3390              Case "CD3/4/8"
3400                  Found = True
3410                  CD348Found = True
3420              End Select
3430          End If
3440      Next

3450      If Found Then

3460          Young = False
3470          sql = "SELECT DoB FROM Demographics WHERE " & _
                    "SampleID = '" & txtSampleID & "'"
3480          Set tb = New Recordset
3490          RecOpenServer 0, tb, sql
3500          If Not tb.EOF Then
3510              If IsDate(tb!DoB) Then
3520                  If DateDiff("d", tb!DoB & "", Now) < 120 Then
3530                      Young = True
3540                  End If
3550              End If
3560          End If

3570          sql = "UPDATE demographics " & _
                    "SET Urgent = " & IIf(chkUrgent.Value = 1, 1, 0) & " " & _
                    "WHERE SampleID = '" & txtSampleID & "'"
3580          Cnxn(0).Execute sql

3590          If ESRFound Or ReticsFound Or MonoSpotFound Or MalariaFound Or SickleFound Then
3600              sql = "SELECT * FROM HaemResults WHERE " & _
                        "SampleID = '" & txtSampleID & "'"
3610              Set tb = New Recordset
3620              RecOpenClient 0, tb, sql
3630              If tb.EOF Then
3640                  tb.AddNew
3650                  tb!Rundate = Format$(Now, "dd/mmm/yyyy")
3660                  tb!RunDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
3670                  tb!SampleID = txtSampleID
3680                  tb!FAXed = 0
3690                  tb!Printed = 0
3700                  tb!ccoag = 0
3710                  tb!Valid = 0
3720                  tb!Printed = 0
3730              End If

3740              tb!cESR = IIf(ESRFound, 1, 0)
3750              tb!cRetics = IIf(ReticsFound, 1, 0)
3760              tb!cMonospot = IIf(MonoSpotFound, 1, 0)
3770              tb!cMalaria = IIf(MalariaFound, 1, 0)
3780              tb!cSickledex = IIf(SickleFound, 1, 0)

3790              tb.Update
3800              tb.Close
3810          Else
3820              sql = "UPDATE HaemResults SET " & _
                        "cESR = 0, cRetics = 0, cMonospot = 0, cMalaria = 0, cSickledex = 0 " & _
                        "WHERE SampleID = '" & txtSampleID & "'"
3830              Cnxn(0).Execute sql
3840          End If
3850      End If

3860      If FBCFound Or ReticsFound Or CD348Found Then
3870          If CD348Found Then
3880              OrderString = "CBCEC+CD3/4/8"
3890          ElseIf ReticsFound And Young Then
3900              OrderString = "CBCEL+RETC"
3910          ElseIf ReticsFound And Not Young Then
3920              OrderString = "CBC+RETC"
3930          ElseIf Young Then
3940              OrderString = "CBCEL"
3950          Else
3960              OrderString = "CBC"
3970          End If
3980          sql = "SELECT * FROM HaemRequests WHERE " & _
                    "SampleID = '" & txtSampleID.Text & "'"
3990          Set tb = New Recordset
4000          RecOpenServer 0, tb, sql
4010          If tb.EOF Then
4020              tb.AddNew
4030          End If
4040          tb!SampleID = Val(txtSampleID.Text)
4050          tb!OrderString = OrderString
4060          tb!Programmed = 0
4070          tb.Update
4080          tb.Close
4090      End If

4100      Screen.MousePointer = vbDefault

4110      Exit Sub

SaveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

4120      intEL = Erl
4130      strES = Err.Description
4140      LogError "fNewOrder", "SaveHaem", intEL, strES, sql

4150      Screen.MousePointer = vbDefault

End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkGBottle_Click
' Author    : XPMUser
' Date      : 03/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGBottle_Click()
4160      On Error GoTo chkGBottle_Click_Error


4170      If chkGBottle.Value = 0 Then
4180          chkGBottle.Caption = "Glucose bottle is NOT in use"
4190          chkGBottle.BackColor = vbRed
4200      ElseIf chkGBottle.Value = 1 Then
4210          chkGBottle.Caption = "Glucose bottle is in use"
4220          chkGBottle.BackColor = &H8000000F
4230      End If


4240      Exit Sub


chkGBottle_Click_Error:

          Dim strES As String
          Dim intEL As Integer

4250      intEL = Erl
4260      strES = Err.Description
4270      LogError "frmNewOrder", "chkGBottle_Click", intEL, strES
End Sub



Private Sub cmdCancelAMSOrder_Click()

      Dim sql As String

4280  On Error GoTo cmdCancelAMSOrder_Click_Error


4290  If iMsg("Are you sure you want to cancel this order?", vbYesNo, "Cancel Order") = vbYes Then
4300      sql = "UPDATE BioRequests SET AMS = 0 WHERE SampleID = " & txtSampleID
4310      Cnxn(0).Execute sql

4320  End If



4330  Exit Sub

cmdCancelAMSOrder_Click_Error:
      Dim strES As String
      Dim intEL As Integer

4340  intEL = Erl
4350  strES = Err.Description
4360  LogError "frmNewOrder", "cmdCancelAMSOrder_Click", intEL, strES

End Sub

Private Sub cmdHide_Click()
4370      fmeQuestions.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdQuestion_Click()
              
4380      On Error GoTo cmdQuestion_Click_Error
          
          Dim sql As String
          Dim tb As Recordset
          Dim str As String
          
4390      If txtSampleID.Text = "" Then
4400          MsgBox "Sample ID not present.", vbInformation
4410          Exit Sub
4420      End If
          
4430      fmeQuestions.Visible = True
          
4440      flxQuestion.Rows = 1
4450      flxQuestion.row = 0
          
4460      sql = "Select Q.RID, Q.Question, Q.Answer from ocmQuestions "
4470      sql = sql & " Inner Join ocmRequestDetails R ON R.RequestID = Q.RID "
4480      sql = sql & " Where SampleID = '" & txtSampleID.Text & "'"
          
4490      Set tb = New Recordset
4500      RecOpenServer 0, tb, sql
          
4510      If Not tb Is Nothing Then
4520          If Not tb.EOF Then
4530              While Not tb.EOF
4540                  str = "" & vbTab & "" & vbTab & tb!RID & vbTab & tb!Question & vbTab & tb!Answer
4550                  flxQuestion.AddItem (str)
4560                  tb.MoveNext
4570              Wend
4580          End If
4590      End If
          
4600      Exit Sub
cmdQuestion_Click_Error:

          Dim strES As String
          Dim intEL As Integer
          
4610      fmeQuestions.Visible = False
4620      intEL = Erl
4630      strES = Err.Description
4640      LogError "frmNewOrder", "Form_Activate", intEL, strES
          
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Activate
' Author    : XPMUser
' Date      : 03/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Activate()

          Dim LOrS As String

4650      On Error GoTo Form_Activate_Error


4660      pBar = 0
4670      pBar.max = LogOffDelaySecs
4680      TimerBar.Enabled = True

4690      If Activated Then Exit Sub
4700      Activated = True

4710      If sysOptUrgent(0) Then chkUrgent.Visible = True

4720      txtSampleID = mSampleID

4730      If mSampleID <> "" Then
4740          FillKnownCoagOrders
4750          FillKnownHaemOrders
4760          FillKnownHaeOrders
4770          FillKnownBioOrders

4780          tinput.SetFocus
4790      End If

4800      If sysOptAlwaysRequestFBC(0) Then
4810          lstHaem.Selected(0) = True
4820      End If

4830      LOrS = GetOptionSetting("LongOrShortBioNames", "Short")

4840      If LOrS = "Long" Then
4850          optLong = True
4860      Else
4870          optShort = True
4880      End If

4890      CoagChanged = False
4900      HaemChanged = False
4910      BioChanged = False
4920      ImmChanged = False
4930      HaeChanged = False
4940      chkAddOn.Value = 0
4950      If GetOptionSetting("DisableGBottleDetection", 0) = 1 Then
4960          chkGBottle.Value = 0
4970      Else
4980          chkGBottle.Value = 1
4990      End If


5000      Exit Sub


Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

5010      intEL = Erl
5020      strES = Err.Description
5030      LogError "frmNewOrder", "Form_Activate", intEL, strES
End Sub



'---------------------------------------------------------------------------------------
' Procedure : FillKnownHaeOrders
' Author    : Masood
' Date      : 29/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FillKnownHaeOrders()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim TestName As String
      Dim LongOrShort As String


5040  On Error GoTo FillKnownHaeOrders_Error

5050  LongOrShort = IIf(optLong, "Long", "Short")
      'urgent
5060  chkUrgent.Value = 0
5070  sql = "Select Urgent from Demographics where " & _
            "SampleID = '" & Val(txtSampleID) & "'"
5080  Set tb = New Recordset
5090  RecOpenClient 0, tb, sql
5100  If Not tb.EOF Then
5110      If Not IsNull(tb!Urgent) Then
5120          If tb!Urgent Then chkUrgent.Value = 1
5130      End If
5140  End If
5150  With LstHaePanel
5160      For n = 0 To .ListCount - 1
5170          .Selected(n) = False
5180      Next

5190      sql = "SELECT T." & LongOrShort & "Name Name, R.Code " & _
                "FROM HaeRequests R " & _
                "JOIN HaemTestDefinitions T " & _
                "ON R.Code = T.Code " & _
                "WHERE SampleID = '" & Val(txtSampleID) & "' " & _
                "AND InUse = 1 "



5200      sql = "SELECT Code as PanelName   FROM HaeRequests WHERE SampleID = '" & Val(txtSampleID) & "' "



5210      Set tb = New Recordset
5220      RecOpenClient 0, tb, sql
5230      Do While Not tb.EOF
5240          TestName = tb!PanelName    ' CoagNameFor(tb!Code & "")
5250          For n = 0 To .ListCount - 1
5260              If .List(n) = TestName Then
5270                  .Selected(n) = True
5280                  Exit For
5290              End If
5300          Next
5310          tb.MoveNext
5320      Loop
5330  End With


5340  Exit Sub


FillKnownHaeOrders_Error:

      Dim strES As String
      Dim intEL As Integer

5350  intEL = Erl
5360  strES = Err.Description
5370  LogError "frmNewOrder", "FillKnownHaeOrders", intEL, strES, sql

End Sub



Private Sub Form_Deactivate()

5380      TimerBar.Enabled = False

End Sub


Private Sub Form_Load()

5390      FillLists
5400      LoadBarCodes
5410      FillQuickBioNames
5420      Call FormatGrid

5430      Activated = False

End Sub

Private Sub FillKnownCoagOrders()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim TestName As String

5440  On Error GoTo FillKnownCoagOrders_Error
      'Urgent
5450  chkUrgent.Value = 0
5460  sql = "Select Urgent from Demographics where " & _
            "SampleID = '" & Val(txtSampleID) & "'"
5470  Set tb = New Recordset
5480  RecOpenClient 0, tb, sql
5490  If Not tb.EOF Then
5500      If Not IsNull(tb!Urgent) Then
5510          If tb!Urgent Then chkUrgent.Value = 1
5520      End If
5530  End If
5540  For n = 0 To lstCoag.ListCount - 1
5550      lstCoag.Selected(n) = False
5560  Next

5570  sql = "select * from CoagRequests where " & _
            "SampleID = '" & Val(mSampleID) & "'"
5580  Set tb = New Recordset
5590  RecOpenClient 0, tb, sql
5600  Do While Not tb.EOF
5610      TestName = CoagNameFor(tb!Code & "")
5620      For n = 0 To lstCoag.ListCount - 1
5630          If lstCoag.List(n) = TestName Then
5640              lstCoag.Selected(n) = True
5650              Exit For
5660          End If
5670      Next
5680      tb.MoveNext
5690  Loop

5700  Exit Sub

FillKnownCoagOrders_Error:

      Dim strES As String
      Dim intEL As Integer

5710  intEL = Erl
5720  strES = Err.Description
5730  LogError "fNewOrder", "FillKnownCoagOrders", intEL, strES, sql


End Sub

Private Sub FillKnownHaemOrders()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

5740  On Error GoTo FillKnownHaemOrders_Error
      'Urgent
5750  chkUrgent.Value = 0
5760  sql = "Select Urgent from Demographics where " & _
            "SampleID = '" & Val(txtSampleID) & "'"
5770  Set tb = New Recordset
5780  RecOpenClient 0, tb, sql
5790  If Not tb.EOF Then
5800      If Not IsNull(tb!Urgent) Then
5810          If tb!Urgent Then chkUrgent.Value = 1
5820      End If
5830  End If
5840  For n = 0 To lstHaem.ListCount - 1
5850      lstHaem.Selected(n) = False
5860  Next

5870  sql = "select * from HaemResults where " & _
            "SampleID = '" & Val(mSampleID) & "'"
5880  Set tb = New Recordset
5890  RecOpenClient 0, tb, sql
5900  If Not tb.EOF Then
5910      lstHaem.Selected(0) = True
5920      If tb!cESR <> 0 Then lstHaem.Selected(1) = True
5930      If tb!cRetics <> 0 Then lstHaem.Selected(2) = True
5940      If tb!cMonospot <> 0 Then lstHaem.Selected(3) = True
5950      If tb!cMalaria <> 0 Then lstHaem.Selected(4) = True
5960      If tb!cSickledex <> 0 Then lstHaem.Selected(5) = True
5970  End If

5980  Exit Sub

FillKnownHaemOrders_Error:

      Dim strES As String
      Dim intEL As Integer

5990  intEL = Erl
6000  strES = Err.Description
6010  LogError "fNewOrder", "FillKnownHaemOrders", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6020      pBar = 0

End Sub


Private Sub Form_Paint()

6030      If sysOptAlwaysRequestFBC(0) Then
6040          lstHaem.Selected(0) = True
6050      End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

6060      Activated = False
6070      mChart = ""

End Sub

Private Sub lstBloodTests_Click()

6080      pBar = 0

End Sub

Private Sub lstBloodTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6090      BioChanged = True

End Sub


Private Sub lstCSFTests_Click()

6100      pBar = 0

End Sub

Private Sub lstCSFTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6110      BioChanged = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LstHaePanel_Click
' Author    : Masood
' Date      : 29/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LstHaePanel_Click()

          Dim t As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim Found As Boolean
6120      HaeChanged = True
6130  On Error GoTo LstHaePanel_Click_Error

6140  With lstHaeTests
6150    sql = "Select * from HaePanels WHERE PanelName ='" & LstHaePanel.List(LstHaePanel.ListIndex) & "'"
6160    Set tb = New Recordset
6170    RecOpenServer 0, tb, sql
6180    While tb.EOF = False
6190        Found = False
6200        For t = 0 To .ListCount - 1
6210            If UCase(.List(t)) = UCase((tb!Content)) Then
6220                .Selected(t) = True
6230                Found = True
6240                Exit For
6250            End If
6260        Next
6270        tb.MoveNext
6280    Wend

6290  End With

6300  HaeChanged = True


6310  Exit Sub


LstHaePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

6320  intEL = Erl
6330  strES = Err.Description
6340  LogError "frmNewOrder", "LstHaePanel_Click", intEL, strES, sql
End Sub

Private Sub lstHaeTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
6350      HaeChanged = True
End Sub

Private Sub lstSerumPanel_Click()

          Dim t As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim Found As Boolean

6360      On Error GoTo lstSerumPanel_Click_Error

6370      pBar = 0

6380      If lstSerumPanel = "" Then
6390          lstSerumPanel.Selected(lstSerumPanel.ListIndex) = False
6400          Exit Sub
6410      End If

6420      sql = "Select * from Panels where " & _
                "PanelName = '" & lstSerumPanel.Text & "'" & _
                "and PanelType = 'S' "
6430      Set tb = New Recordset
6440      RecOpenServer 0, tb, sql
6450      Do While Not tb.EOF
6460          Found = False
6470          For t = 0 To lstSerumTests.ListCount - 1
6480              If UCase(QuickBioLongNameFor(lstSerumTests.List(t))) = UCase(QuickBioLongNameFor(tb!Content)) Then
6490                  lstSerumTests.Selected(t) = True
6500                  Found = True
6510                  Exit For
6520              End If
6530          Next
6540          If Not Found Then
6550              For t = 0 To lstImmunoTests.ListCount - 1
6560                  If UCase(QuickBioLongNameFor(lstImmunoTests.List(t))) = UCase(QuickBioLongNameFor(tb!Content)) Then
6570                      lstImmunoTests.Selected(t) = True
6580                      Exit For
6590                  End If
6600              Next
6610          End If
6620          tb.MoveNext
6630      Loop

6640      BioChanged = True

6650      Exit Sub

lstSerumPanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

6660      intEL = Erl
6670      strES = Err.Description
6680      LogError "frmNewOrder", "lstSerumPanel_Click", intEL, strES, sql

End Sub

Private Function QuickBioLongNameFor(ByVal LongOrShortName As String) As String

          Dim n As Integer

6690      On Error GoTo QuickBioLongNameFor_Error

6700      For n = 1 To UBound(QuickBioNames)
6710          If LongOrShortName = QuickBioNames(n).Long Or LongOrShortName = QuickBioNames(n).Short Then
6720              QuickBioLongNameFor = QuickBioNames(n).Long
6730              Exit For
6740          End If
6750      Next

6760      Exit Function

QuickBioLongNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

6770      intEL = Erl
6780      strES = Err.Description
6790      LogError "frmNewOrder", "QuickBioLongNameFor", intEL, strES

End Function
Private Function FillQuickBioNames() As String

          Dim tb As Recordset
          Dim sql As String
          Dim UB As Integer

6800      On Error GoTo FillQuickBioNames_Error

6810      ReDim Preserve QuickBioNames(0 To 0)

6820      sql = "SELECT DISTINCT LongName, ShortName " & _
                "FROM BioTestDefinitions "
6830      Set tb = New Recordset
6840      RecOpenServer 0, tb, sql
6850      Do While Not tb.EOF
6860          UB = UBound(QuickBioNames) + 1
6870          ReDim Preserve QuickBioNames(0 To UB)
6880          QuickBioNames(UB).Short = tb!ShortName & ""
6890          QuickBioNames(UB).Long = tb!LongName & ""
6900          tb.MoveNext
6910      Loop

6920      Exit Function

FillQuickBioNames_Error:

          Dim strES As String
          Dim intEL As Integer

6930      intEL = Erl
6940      strES = Err.Description
6950      LogError "frmNewOrder", "FillQuickBioNames", intEL, strES, sql

End Function

Private Sub lstSerumTests_Click()

6960      pBar = 0

End Sub


Private Sub lstSerumTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6970      BioChanged = True

End Sub


Private Sub lstCoag_Click()

6980      pBar = 0

End Sub

Private Sub lstCoag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6990      CoagChanged = True

End Sub


Private Sub lstCoagPanel_Click()

          Dim t As Integer
          Dim tb As Recordset
          Dim sql As String

7000      On Error GoTo lstCoagPanel_Click_Error

7010      pBar = 0

7020      If lstCoagPanel = "" Then
7030          lstCoagPanel.Selected(lstCoagPanel.ListIndex) = False
7040          Exit Sub
7050      End If

7060      sql = "Select * from CoagPanels where " & _
                "PanelName = '" & lstCoagPanel.Text & "'"
7070      Set tb = New Recordset
7080      RecOpenServer 0, tb, sql
7090      Do While Not tb.EOF
7100          For t = 0 To lstCoag.ListCount - 1
7110              If UCase(lstCoag.List(t)) = UCase(tb!Content) Then
7120                  lstCoag.Selected(t) = True
7130                  Exit For
7140              End If
7150          Next
7160          tb.MoveNext
7170      Loop

7180      CoagChanged = True

7190      Exit Sub

lstCoagPanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

7200      intEL = Erl
7210      strES = Err.Description
7220      LogError "frmNewOrder", "lstCoagPanel_Click", intEL, strES, sql


End Sub

Private Sub lstHaem_Click()

7230      pBar = 0

End Sub

Private Sub lstHaem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7240      HaemChanged = True

End Sub


Private Sub lstImmunoTests_Click()

7250      pBar = 0

End Sub

Private Sub lstImmunoTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7260      ImmChanged = True

End Sub


Private Sub lstSweatTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7270      BioChanged = True

End Sub

Private Sub lstUrinePanel_Click()

         Dim n As Integer
          Dim tb As Recordset
          Dim sql As String

7280      On Error GoTo lstUrinePanel_Click_Error

7290      pBar = 0
      '+++ Junaid 12-10-2023
      '30        For n = 0 To lstUrineTests.ListCount - 1
      '40            lstUrineTests.Selected(n) = False
      '50        Next
      '---Junaid
7300      sql = "Select * from Panels where " & _
                "PanelName = '" & lstUrinePanel.Text & "'" & _
                "and PanelType = 'U' "
7310      Set tb = New Recordset
7320      RecOpenServer 0, tb, sql
7330      Do While Not tb.EOF
7340          For n = 0 To lstUrineTests.ListCount - 1
7350              If UCase(BioLongNameFor(lstUrineTests.List(n))) = UCase(BioLongNameFor(tb!Content)) Then
7360                  lstUrineTests.Selected(n) = True
7370                  Exit For
7380              End If
7390          Next
7400          tb.MoveNext
7410      Loop

7420      Exit Sub


lstUrinePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

7430      intEL = Erl
7440      strES = Err.Description
7450      LogError "frmNewOrder", "lstUrinePanel_Click", intEL, strES, sql


End Sub

Private Sub lstUrinePanel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7460      BioChanged = True

7470      bsave.Enabled = True

End Sub


Private Sub lstUrineTests_Click()

7480      pBar = 0

End Sub


Private Sub lstUrineTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7490      BioChanged = True

End Sub


Private Sub optLong_Click()

7500      SetDisplay "Long"

7510      FillLists
7520      FillKnownBioOrders
7530      FillKnownHaeOrders
End Sub

Private Sub optShort_Click()

7540      SetDisplay "Short"

7550      FillLists
7560      FillKnownBioOrders
7570      FillKnownHaeOrders
End Sub


Private Sub TimerBar_Timer()

7580      pBar = pBar + 1

7590      If pBar = pBar.max Then
7600          Unload Me
7610          Exit Sub
7620      End If

End Sub

Private Sub tinput_KeyPress(KeyAscii As Integer)

7630      If KeyAscii = 13 Then
7640          KeyAscii = 0
7650          tinput_LostFocus
7660      End If

End Sub


Private Sub tinput_LostFocus()

7670      If Trim$(tinput) = "" Then Exit Sub

7680      tinput = UCase$(Trim$(tinput))

7690      If Not CheckCodes() Then
7700          If Not CheckSerumPanel() Then
7710              If Not CheckUrinePanel() Then
7720                  If Not CheckSerum() Then
7730                      If Not CheckUrine() Then
7740                          If CheckCSF() Then

7750                              BioChanged = True
7760                          End If
7770                      End If
7780                  End If
7790              End If
7800          End If
7810      End If
7820      tinput = ""
7830      If tinput.Visible Then
7840          tinput.SetFocus
7850      End If

End Sub

Private Sub txtsampleid_GotFocus()

7860      ClearRequests
7870      If sysOptAlwaysRequestFBC(0) Then
7880          lstHaem.Selected(0) = True
7890          lstHaem.Refresh
7900      End If

End Sub


Private Sub txtsampleid_KeyPress(KeyAscii As Integer)

7910      If KeyAscii = 13 Then
7920          KeyAscii = 0
7930          tinput.SetFocus
7940      End If

End Sub


Private Sub txtsampleid_LostFocus()

          Dim sql As String
          Dim tb As Recordset

7950      On Error GoTo txtsampleid_LostFocus_Error

7960      If Val(txtSampleID) = 0 Then Exit Sub

7970      txtSampleID = Trim$(txtSampleID)

7980      mSampleID = txtSampleID

7990      If sysOptUrgent(0) Then
8000          sql = "Select * from demographics where " & _
                    "sampleid = '" & txtSampleID & "'"
8010          Set tb = New Recordset
8020          RecOpenServer 0, tb, sql
8030          If Not tb.EOF Then
8040              If tb!Urgent = 1 Then chkUrgent.Value = 1
8050          End If
8060      End If

8070      FillKnownBioOrders
8080      FillKnownHaeOrders
8090      FillKnownCoagOrders
8100      FillKnownHaemOrders

8110      CoagChanged = False
8120      HaemChanged = False
8130      BioChanged = False
8140      ImmChanged = False

8150      GetChartNumber

8160      Exit Sub

txtsampleid_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

8170      intEL = Erl
8180      strES = Err.Description
8190      LogError "frmNewOrder", "txtsampleid_LostFocus", intEL, strES, sql


End Sub



Public Property Get Cancelled() As Boolean

8200      Cancelled = mCancelled

End Property




'---------------------------------------------------------------------------------------
' Procedure : UpDateRequestsHae
' Author    : Masood
' Date      : 29/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub UpDateRequestsHae(ByVal Discipline As String, _
                              ByVal TestName As String, _
                              ShortNameOrLongName As String)

          Dim sql As String

8210      On Error GoTo UpDateRequestsHae_Error

          '
          'sql = "INSERT INTO " & Discipline & "Requests " & vbNewLine
          '    sql = sql & "(SampleId, Code, DateTimeOfRecord, SampleType, Units,  UserName) " & vbNewLine
          '    sql = sql & "SELECT DISTINCT '" & txtSampleID & "', " & vbNewLine
          '    sql = sql & "        Code , getdate(), " & vbNewLine
          '    sql = sql & "        SampleType , Units,  " & vbNewLine
          '    sql = sql & "       '" & UserName & "'  FROM " & IIf((Discipline = "Hae"), "Haem", Discipline) & "TestDefinitions " & vbNewLine
          '    sql = sql & "        " & vbNewLine
          '
          '  If ShortNameOrLongName = "Short" Then
          '     sql = sql & " WHERE ShortName = '" & TestName & "' " & vbNewLine
          '  Else
          '      sql = sql & " WHERE LongName = '" & TestName & "' " & vbNewLine
          '  End If
          '    sql = sql & " AND InUse = 1" & vbNewLine




8220      sql = "INSERT INTO " & Discipline & "Requests " & vbNewLine
8230      sql = sql & "(SampleId, Code, DateTimeOfRecord, SampleType, Units,  UserName,Analyser,Programmed) " & vbNewLine
8240      sql = sql & " VALUES ( " & vbNewLine
8250      sql = sql & " '" & txtSampleID & "', " & vbNewLine
8260      sql = sql & "        '" & TestName & "' , getdate(), " & vbNewLine
8270      sql = sql & "        'Blood EDTA' , '',  " & vbNewLine
8280      sql = sql & "       '" & UserName & "','IPU',0 " & vbNewLine
8290      sql = sql & " )"

8300      Cnxn(0).Execute sql




8310      Exit Sub


UpDateRequestsHae_Error:

          Dim strES As String
          Dim intEL As Integer

8320      intEL = Erl
8330      strES = Err.Description
8340      LogError "frmNewOrder", "UpDateRequestsHae", intEL, strES, sql

End Sub

Private Sub FormatGrid()


8350  On Error GoTo FormatGrid_Error

8360  flxQuestion.Rows = 1
8370  flxQuestion.row = 0

8380  flxQuestion.ColWidth(fcsLine_NO) = 100

8390  flxQuestion.TextMatrix(0, fcsSr) = "Sr. #"
8400  flxQuestion.ColWidth(fcsSr) = 550
8410  flxQuestion.ColAlignment(fcsSr) = flexAlignRightCenter

8420  flxQuestion.TextMatrix(0, fcsRID) = "Request ID"
8430  flxQuestion.ColWidth(fcsRID) = 6200
8440  flxQuestion.ColAlignment(fcsRID) = flexAlignLeftCenter

8450  flxQuestion.TextMatrix(0, fcsQus) = "Question"
8460  flxQuestion.ColWidth(fcsQus) = 1050
8470  flxQuestion.ColAlignment(fcsQus) = flexAlignRightCenter

8480  flxQuestion.TextMatrix(0, fcsAns) = "Answer"
8490  flxQuestion.ColWidth(fcsAns) = 1050
8500  flxQuestion.ColAlignment(fcsAns) = flexAlignCenterCenter

8510  Exit Sub
FormatGrid_Error:

8520  LogError "frmNewOrder", "FormatGrid", Erl, Err.Description




End Sub
