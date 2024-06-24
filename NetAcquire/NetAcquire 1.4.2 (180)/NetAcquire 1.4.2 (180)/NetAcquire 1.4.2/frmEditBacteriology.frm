VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditMicrobiology 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   14535
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   750
      Top             =   8490
   End
   Begin VB.Frame fmeQuestions 
      Caption         =   "Questions"
      Height          =   3195
      Left            =   2880
      TabIndex        =   326
      Top             =   2640
      Width           =   6495
      Begin VB.CommandButton btnHideQuestions 
         Caption         =   "Hide"
         Height          =   315
         Left            =   5250
         TabIndex        =   328
         Top             =   2850
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid flxQuestions 
         Height          =   2595
         Left            =   60
         TabIndex        =   327
         Top             =   270
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   4
      End
   End
   Begin VB.CommandButton cmdValidateMicro 
      Caption         =   "&Validate"
      Height          =   675
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   308
      Top             =   2340
      Width           =   1272
   End
   Begin VB.CommandButton cmdIntrim 
      Caption         =   "&Interim "
      Height          =   675
      Left            =   13140
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmEditBacteriology.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   307
      Top             =   3060
      Width           =   1272
   End
   Begin VB.TextBox txtCommentMicro 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   306
      Top             =   2280
      Width           =   4515
   End
   Begin VB.CommandButton cmdOrderBiomnis 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   915
      Picture         =   "frmEditBacteriology.frx":0B2C
      Style           =   1  'Graphical
      TabIndex        =   302
      ToolTipText     =   "Order via Biomnis"
      Top             =   2220
      Width           =   645
   End
   Begin VB.TextBox txtMultiSeltdDemoForLabNoUpd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   15060
      TabIndex        =   298
      Top             =   420
      Width           =   3480
   End
   Begin VB.TextBox txtLabNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15060
      MaxLength       =   8
      TabIndex        =   297
      Top             =   120
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid gMDemoLabNoUpd 
      Height          =   1185
      Left            =   15000
      TabIndex        =   296
      Top             =   5280
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   2090
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdOrderBiomnis 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   210
      Picture         =   "frmEditBacteriology.frx":13F6
      Style           =   1  'Graphical
      TabIndex        =   294
      ToolTipText     =   "Order via Biomnis"
      Top             =   2220
      Width           =   645
   End
   Begin VB.CommandButton cmdUnLock 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Unlock Demographics"
      Height          =   1215
      Left            =   13110
      Picture         =   "frmEditBacteriology.frx":1CC0
      Style           =   1  'Graphical
      TabIndex        =   293
      Top             =   930
      Width           =   1245
   End
   Begin VB.CommandButton cmdHealthLink 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1620
      Picture         =   "frmEditBacteriology.frx":2B8A
      Style           =   1  'Graphical
      TabIndex        =   283
      Top             =   2220
      Width           =   645
   End
   Begin VB.CommandButton cmdViewReports 
      BackColor       =   &H00FFFF00&
      Caption         =   "Reports"
      Height          =   615
      Left            =   13740
      Picture         =   "frmEditBacteriology.frx":3454
      Style           =   1  'Graphical
      TabIndex        =   282
      Top             =   5160
      Width           =   675
   End
   Begin VB.CommandButton cmdSetValid 
      Caption         =   "Set Valid Date"
      Height          =   765
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":3D1E
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   5850
      Width           =   1275
   End
   Begin VB.CommandButton cmdSensArchive 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sens Archive"
      Height          =   255
      Left            =   11745
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   2220
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSensRepeat 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sens Repeat"
      Height          =   255
      Left            =   10470
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   2220
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdIsoArchive 
      BackColor       =   &H0000FFFF&
      Caption         =   "ISO Archive"
      Height          =   255
      Left            =   9195
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   2220
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdIsoRepeat 
      BackColor       =   &H0000FFFF&
      Caption         =   "ISO Repeat"
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   2220
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   705
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":4160
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Log as Phoned"
      Top             =   4455
      Width           =   1275
   End
   Begin VB.TextBox txtNOPAS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15900
      TabIndex        =   82
      Top             =   1770
      Width           =   1245
   End
   Begin VB.CommandButton cmdSaveHold 
      Caption         =   "Save && Hold"
      Height          =   645
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":45A2
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   6660
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveMicro 
      Caption         =   "&Save Details"
      Enabled         =   0   'False
      Height          =   705
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":49E4
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   7320
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   615
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":504E
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   12450
      Top             =   -90
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   450
      TabIndex        =   52
      Top             =   0
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2085
      Index           =   0
      Left            =   210
      TabIndex        =   28
      Top             =   150
      Width           =   14235
      Begin VB.CommandButton cmdViewLog 
         Caption         =   "...."
         Height          =   270
         Left            =   1815
         TabIndex        =   305
         Top             =   1695
         Width           =   570
      End
      Begin VB.CommandButton cmdViewScan 
         Caption         =   "&View Scan"
         Height          =   840
         Left            =   2460
         Picture         =   "frmEditBacteriology.frx":5358
         Style           =   1  'Graphical
         TabIndex        =   301
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdscan 
         Caption         =   "&Scan"
         Height          =   840
         Left            =   2460
         Picture         =   "frmEditBacteriology.frx":AB46
         Style           =   1  'Graphical
         TabIndex        =   300
         Top             =   1140
         Width           =   885
      End
      Begin VB.CommandButton cmdPatientNotePad 
         Height          =   500
         Left            =   3540
         Picture         =   "frmEditBacteriology.frx":BBC8
         Style           =   1  'Graphical
         TabIndex        =   299
         Tag             =   "bprint"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton bsearchDob 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Left            =   12930
         TabIndex        =   295
         Top             =   270
         Width           =   405
      End
      Begin VB.TextBox txtForeName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7860
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtSurName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5670
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdAddToConsultantList 
         Caption         =   "Remove from  Consultant List"
         Height          =   255
         Left            =   60
         TabIndex        =   87
         Top             =   1410
         Width           =   2325
      End
      Begin VB.ComboBox cmbConsultantVal 
         Height          =   315
         Left            =   90
         TabIndex        =   85
         Text            =   "cmbConsultantVal"
         Top             =   1680
         Width           =   1635
      End
      Begin VB.TextBox txtChart 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4110
         MaxLength       =   8
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   10620
         MaxLength       =   10
         TabIndex        =   5
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   10620
         MaxLength       =   4
         TabIndex        =   9
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   10620
         MaxLength       =   6
         TabIndex        =   10
         Top             =   990
         Width           =   1545
      End
      Begin VB.Frame fraSampleID 
         Height          =   1395
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   2385
         Begin VB.ComboBox cMRU 
            Height          =   315
            Left            =   570
            TabIndex        =   53
            Text            =   "cMRU"
            Top             =   1020
            Width           =   1605
         End
         Begin VB.TextBox txtSampleID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   180
            MaxLength       =   12
            TabIndex        =   0
            Top             =   540
            Width           =   1785
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   480
            Left            =   1920
            TabIndex        =   31
            Top             =   510
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   847
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtSampleID"
            BuddyDispid     =   196648
            OrigLeft        =   1920
            OrigTop         =   540
            OrigRight       =   2160
            OrigBottom      =   1020
            Max             =   2147483647
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "MRU"
            Height          =   195
            Index           =   21
            Left            =   150
            TabIndex        =   54
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image iRelevant 
            Height          =   480
            Index           =   1
            Left            =   1500
            Picture         =   "frmEditBacteriology.frx":C492
            Top             =   120
            Width           =   480
         End
         Begin VB.Image iRelevant 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   120
            Picture         =   "frmEditBacteriology.frx":C79C
            Top             =   120
            Width           =   480
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sample ID"
            Height          =   195
            Index           =   42
            Left            =   720
            TabIndex        =   32
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Searc&h"
         Height          =   345
         Left            =   9315
         TabIndex        =   4
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   285
         Left            =   12210
         TabIndex        =   29
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblSurNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "SurName"
         Height          =   195
         Left            =   5670
         TabIndex        =   168
         Top             =   390
         Width           =   660
      End
      Begin VB.Label lblForeNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "ForeName"
         Height          =   195
         Left            =   7920
         TabIndex        =   167
         Top             =   390
         Width           =   735
      End
      Begin VB.Label lblABsInUse 
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Left            =   10620
         TabIndex        =   74
         Top             =   1350
         Width           =   2235
      End
      Begin VB.Label label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   4110
         TabIndex        =   73
         Top             =   1710
         Width           =   5865
      End
      Begin VB.Label lblSiteDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4110
         TabIndex        =   69
         Top             =   1410
         Width           =   5865
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Chart #"
         Height          =   285
         Left            =   4110
         TabIndex        =   58
         ToolTipText     =   "Click to change Location"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label lAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4110
         TabIndex        =   57
         Top             =   1050
         Width           =   5865
      End
      Begin VB.Label lNoPrevious 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Previous Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6420
         TabIndex        =   55
         Top             =   150
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   20
         Left            =   10170
         TabIndex        =   35
         Top             =   300
         Width           =   405
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   19
         Left            =   10260
         TabIndex        =   34
         Top             =   660
         Width           =   285
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   18
         Left            =   10290
         TabIndex        =   33
         Top             =   1020
         Width           =   270
      End
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Height          =   885
      Left            =   15690
      Picture         =   "frmEditBacteriology.frx":CAA6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2340
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bHistory 
      Caption         =   "&History"
      Height          =   675
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":CDB0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8010
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":D1F2
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "bprint"
      Top             =   3720
      Width           =   1272
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "FAX"
      Height          =   825
      Index           =   0
      Left            =   11340
      Picture         =   "frmEditBacteriology.frx":D8DC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   13140
      Picture         =   "frmEditBacteriology.frx":DD1E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8685
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6435
      Left            =   180
      TabIndex        =   17
      Top             =   2910
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   11351
      _Version        =   393216
      Style           =   1
      Tabs            =   13
      TabsPerRow      =   16
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmEditBacteriology.frx":E388
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRequestID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSaveDemographics"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSaveInc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOrderTests"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkPregnant"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCopyFromPrevious"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdTagRepeat"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdOrderExt"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fmeAntibiotics"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fmeAntibio"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "btnOtherQuestions"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Urine"
      TabPicture(1)   =   "frmEditBacteriology.frx":E3A4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtUrineComment"
      Tab(1).Control(1)=   "cmdLock(1)"
      Tab(1).Control(2)=   "fraMicroResult(1)"
      Tab(1).Control(3)=   "label1(1)"
      Tab(1).Control(4)=   "lblPrinted(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Identification"
      TabPicture(2)   =   "frmEditBacteriology.frx":E3C0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Faeces"
      TabPicture(3)   =   "frmEditBacteriology.frx":E3DC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "label1(102)"
      Tab(3).Control(1)=   "lblViewOrganism"
      Tab(3).Control(2)=   "grdDay(3)"
      Tab(3).Control(3)=   "grdDay(2)"
      Tab(3).Control(4)=   "udHistoricalFaecesView"
      Tab(3).Control(5)=   "grdDay(1)"
      Tab(3).Control(6)=   "Frame2(2)"
      Tab(3).Control(7)=   "Frame2(1)"
      Tab(3).Control(8)=   "Frame2(0)"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "C && S"
      TabPicture(4)   =   "frmEditBacteriology.frx":E3F8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "imgSquareCross"
      Tab(4).Control(1)=   "imgSquareTick"
      Tab(4).Control(2)=   "lblPrinted(4)"
      Tab(4).Control(3)=   "cmdLock(4)"
      Tab(4).Control(4)=   "fraMicroResult(4)"
      Tab(4).Control(5)=   "grdAB(1)"
      Tab(4).Control(6)=   "grdAB(2)"
      Tab(4).Control(7)=   "grdAB(3)"
      Tab(4).Control(8)=   "grdAB(4)"
      Tab(4).Control(9)=   "cmdConC"
      Tab(4).Control(10)=   "txtConC"
      Tab(4).Control(11)=   "cmdMSC"
      Tab(4).Control(12)=   "txtMSC"
      Tab(4).Control(13)=   "cmbMSC"
      Tab(4).Control(14)=   "cmbConC"
      Tab(4).Control(15)=   "cmdUseSecondary(1)"
      Tab(4).Control(16)=   "cmdUseSecondary(2)"
      Tab(4).Control(17)=   "cmdUseSecondary(3)"
      Tab(4).Control(18)=   "cmdUseSecondary(4)"
      Tab(4).ControlCount=   19
      TabCaption(5)   =   "FOB"
      TabPicture(5)   =   "frmEditBacteriology.frx":E414
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblPrinted(5)"
      Tab(5).Control(1)=   "cmdLock(5)"
      Tab(5).Control(2)=   "fraMicroResult(5)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Rota/Adeno"
      TabPicture(6)   =   "frmEditBacteriology.frx":E430
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblPrinted(6)"
      Tab(6).Control(1)=   "cmdLock(6)"
      Tab(6).Control(2)=   "fraMicroResult(6)"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Red/Sub"
      TabPicture(7)   =   "frmEditBacteriology.frx":E44C
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblPrinted(7)"
      Tab(7).Control(1)=   "cmdLock(7)"
      Tab(7).Control(2)=   "fraMicroResult(7)"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "RSV"
      TabPicture(8)   =   "frmEditBacteriology.frx":E468
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "fraMicroResult(8)"
      Tab(8).Control(1)=   "cmdLock(8)"
      Tab(8).Control(2)=   "lblPrinted(8)"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "CSF"
      TabPicture(9)   =   "frmEditBacteriology.frx":E484
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lblPrinted(9)"
      Tab(9).Control(1)=   "cmdLock(9)"
      Tab(9).Control(2)=   "fraMicroResult(9)"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "C.diff"
      TabPicture(10)  =   "frmEditBacteriology.frx":E4A0
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "cmdLock(10)"
      Tab(10).Control(1)=   "fraMicroResult(10)"
      Tab(10).Control(2)=   "lblPrinted(10)"
      Tab(10).ControlCount=   3
      TabCaption(11)  =   "O/P"
      TabPicture(11)  =   "frmEditBacteriology.frx":E4BC
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "cmdLock(11)"
      Tab(11).Control(1)=   "fraMicroResult(11)"
      Tab(11).Control(2)=   "lblPrinted(11)"
      Tab(11).ControlCount=   3
      TabCaption(12)  =   "Identification"
      TabPicture(12)  =   "frmEditBacteriology.frx":E4D8
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "fraIdentification(0)"
      Tab(12).Control(1)=   "fraIdentification(1)"
      Tab(12).Control(2)=   "fraIdentification(2)"
      Tab(12).Control(3)=   "fraIdentification(3)"
      Tab(12).Control(4)=   "cmdGramPrep"
      Tab(12).ControlCount=   5
      Begin VB.Frame Frame2 
         Caption         =   "Day 1"
         Height          =   2085
         Index           =   0
         Left            =   -74910
         TabIndex        =   353
         Top             =   390
         Width           =   6765
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   11
            Left            =   1110
            TabIndex        =   362
            Text            =   "cmbDay1"
            Top             =   480
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   12
            Left            =   2970
            TabIndex        =   361
            Text            =   "cmbDay1"
            Top             =   480
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   13
            Left            =   4830
            TabIndex        =   360
            Text            =   "cmbDay1"
            Top             =   480
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   21
            Left            =   1110
            TabIndex        =   359
            Text            =   "cmbDay1"
            Top             =   840
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   22
            Left            =   2970
            TabIndex        =   358
            Text            =   "cmbDay1"
            Top             =   840
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   23
            Left            =   4830
            TabIndex        =   357
            Text            =   "cmbDay1"
            Top             =   840
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   31
            Left            =   1110
            TabIndex        =   356
            Text            =   "cmbDay1"
            Top             =   1200
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   32
            Left            =   2970
            TabIndex        =   355
            Text            =   "cmbDay1"
            Top             =   1200
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay1 
            Height          =   315
            Index           =   33
            Left            =   4830
            TabIndex        =   354
            Text            =   "cmbDay1"
            Top             =   1200
            Width           =   1755
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "XLD"
            Height          =   195
            Index           =   62
            Left            =   345
            TabIndex        =   368
            Top             =   540
            Width           =   315
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "DCA"
            Height          =   195
            Index           =   69
            Left            =   330
            TabIndex        =   367
            Top             =   900
            Width           =   330
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "SMAC"
            Height          =   195
            Index           =   91
            Left            =   210
            TabIndex        =   366
            Top             =   1260
            Width           =   450
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 1"
            Height          =   195
            Index           =   93
            Left            =   1620
            TabIndex        =   365
            Top             =   0
            Width           =   795
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 2"
            Height          =   195
            Index           =   96
            Left            =   3390
            TabIndex        =   364
            Top             =   0
            Width           =   795
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 3"
            Height          =   195
            Index           =   99
            Left            =   5280
            TabIndex        =   363
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Day 2"
         Height          =   2235
         Index           =   1
         Left            =   -74910
         TabIndex        =   337
         Top             =   2490
         Width           =   6765
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   11
            Left            =   1110
            TabIndex        =   346
            Text            =   "cmbDay2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   12
            Left            =   2970
            TabIndex        =   345
            Text            =   "cmbDay2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   13
            Left            =   4830
            TabIndex        =   344
            Text            =   "cmbDay2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   21
            Left            =   1110
            TabIndex        =   343
            Text            =   "cmbDay2"
            Top             =   990
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   22
            Left            =   2970
            TabIndex        =   342
            Text            =   "cmbDay2"
            Top             =   990
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   23
            Left            =   4830
            TabIndex        =   341
            Text            =   "cmbDay2"
            Top             =   990
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   31
            Left            =   1110
            TabIndex        =   340
            Text            =   "cmbDay2"
            Top             =   1350
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   32
            Left            =   2970
            TabIndex        =   339
            Text            =   "cmbDay2"
            Top             =   1350
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay2 
            Height          =   315
            Index           =   33
            Left            =   4830
            TabIndex        =   338
            Text            =   "cmbDay2"
            Top             =   1350
            Width           =   1755
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "XLD SUB"
            Height          =   195
            Index           =   63
            Left            =   345
            TabIndex        =   352
            Top             =   690
            Width           =   690
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "DCA SUB"
            Height          =   195
            Index           =   70
            Left            =   330
            TabIndex        =   351
            Top             =   1050
            Width           =   705
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "PRESTON"
            Height          =   195
            Index           =   92
            Left            =   255
            TabIndex        =   350
            Top             =   1410
            Width           =   780
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 1"
            Height          =   195
            Index           =   94
            Left            =   1620
            TabIndex        =   349
            Top             =   0
            Width           =   795
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 2"
            Height          =   195
            Index           =   97
            Left            =   3390
            TabIndex        =   348
            Top             =   0
            Width           =   795
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 3"
            Height          =   195
            Index           =   100
            Left            =   5280
            TabIndex        =   347
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Day 3"
         Height          =   1605
         Index           =   2
         Left            =   -74910
         TabIndex        =   329
         Top             =   4740
         Width           =   6765
         Begin VB.ComboBox cmbDay3 
            Height          =   315
            Index           =   1
            Left            =   1110
            TabIndex        =   332
            Text            =   "cmbDay3"
            Top             =   540
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay3 
            Height          =   315
            Index           =   2
            Left            =   2970
            TabIndex        =   331
            Text            =   "cmbDay3"
            Top             =   540
            Width           =   1755
         End
         Begin VB.ComboBox cmbDay3 
            Height          =   315
            Index           =   3
            Left            =   4830
            TabIndex        =   330
            Text            =   "cmbDay3"
            Top             =   540
            Width           =   1755
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "CCDA"
            Height          =   195
            Index           =   64
            Left            =   345
            TabIndex        =   336
            Top             =   630
            Width           =   435
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 1"
            Height          =   195
            Index           =   95
            Left            =   1620
            TabIndex        =   335
            Top             =   0
            Width           =   795
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 2"
            Height          =   195
            Index           =   98
            Left            =   3390
            TabIndex        =   334
            Top             =   0
            Width           =   795
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Organism 3"
            Height          =   195
            Index           =   101
            Left            =   5280
            TabIndex        =   333
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.CommandButton btnOtherQuestions 
         Caption         =   "Other Questions"
         Height          =   1095
         Left            =   5640
         Picture         =   "frmEditBacteriology.frx":E4F4
         Style           =   1  'Graphical
         TabIndex        =   325
         Top             =   3810
         Width           =   855
      End
      Begin VB.Frame fmeAntibio 
         Height          =   3285
         Left            =   10050
         TabIndex        =   321
         Top             =   510
         Visible         =   0   'False
         Width           =   2505
         Begin VB.CommandButton btnAdd 
            Caption         =   "&Add"
            Height          =   270
            Left            =   1260
            TabIndex        =   324
            Top             =   2970
            Width           =   1170
         End
         Begin VB.CommandButton btnHide 
            Caption         =   "&Hide"
            Height          =   270
            Left            =   60
            TabIndex        =   323
            Top             =   2970
            Width           =   1170
         End
         Begin VB.ListBox lstAntibio 
            Height          =   2790
            Left            =   60
            TabIndex        =   322
            Top             =   180
            Width           =   2385
         End
      End
      Begin VB.Frame fmeAntibiotics 
         Caption         =   "Patients Antibiotics/Intended Antibiotics"
         Height          =   1035
         Left            =   6570
         TabIndex        =   314
         Top             =   3180
         Visible         =   0   'False
         Width           =   4515
         Begin VB.CommandButton btnIntAntibiotics 
            Caption         =   "...."
            Height          =   270
            Left            =   4170
            TabIndex        =   320
            Top             =   600
            Width           =   300
         End
         Begin VB.CommandButton btnAntiBiotics 
            Caption         =   "...."
            Height          =   270
            Left            =   4170
            TabIndex        =   319
            Top             =   270
            Width           =   300
         End
         Begin VB.TextBox txtIntAntibiotics 
            Height          =   315
            Left            =   1050
            TabIndex        =   316
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txtAntibiotics 
            Height          =   315
            Left            =   1050
            TabIndex        =   315
            Top             =   270
            Width           =   3105
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Int.Antibiotics"
            Height          =   195
            Index           =   3
            Left            =   45
            TabIndex        =   318
            Top             =   630
            Width           =   945
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Antibiotics"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   317
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   4
         Left            =   -66570
         Picture         =   "frmEditBacteriology.frx":EB7A
         Style           =   1  'Graphical
         TabIndex        =   313
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   3
         Left            =   -69360
         Picture         =   "frmEditBacteriology.frx":EE84
         Style           =   1  'Graphical
         TabIndex        =   312
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   2
         Left            =   -72150
         Picture         =   "frmEditBacteriology.frx":F18E
         Style           =   1  'Graphical
         TabIndex        =   311
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   1
         Left            =   -74940
         Picture         =   "frmEditBacteriology.frx":F498
         Style           =   1  'Graphical
         TabIndex        =   310
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdOrderExt 
         Caption         =   "Order External Tests"
         Height          =   1215
         Left            =   5640
         Picture         =   "frmEditBacteriology.frx":F7A2
         Style           =   1  'Graphical
         TabIndex        =   309
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton cmdTagRepeat 
         Caption         =   "Tag as Repeat"
         Height          =   915
         Left            =   11160
         Picture         =   "frmEditBacteriology.frx":FAAC
         Style           =   1  'Graphical
         TabIndex        =   292
         Top             =   2400
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtUrineComment 
         Height          =   1185
         Left            =   -73620
         MultiLine       =   -1  'True
         TabIndex        =   290
         Top             =   4500
         Width           =   7065
      End
      Begin VB.ComboBox cmbConC 
         Height          =   315
         Left            =   -68910
         TabIndex        =   284
         Text            =   "cmbConC"
         Top             =   5700
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.ComboBox cmbMSC 
         Height          =   315
         Left            =   -74490
         TabIndex        =   287
         Text            =   "cmbMSC"
         Top             =   5700
         Visible         =   0   'False
         Width           =   4605
      End
      Begin VB.TextBox txtMSC 
         Height          =   1005
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   289
         Text            =   "frmEditBacteriology.frx":10376
         Top             =   5340
         Width           =   4635
      End
      Begin VB.CommandButton cmdMSC 
         Caption         =   "..."
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
         Left            =   -69870
         TabIndex        =   288
         ToolTipText     =   "Choose a comment from a list"
         Top             =   5670
         Width           =   435
      End
      Begin VB.TextBox txtConC 
         Height          =   1005
         Left            =   -68940
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   286
         Text            =   "frmEditBacteriology.frx":10391
         Top             =   5340
         Width           =   4575
      End
      Begin VB.CommandButton cmdConC 
         Caption         =   "..."
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
         Left            =   -64350
         TabIndex        =   285
         ToolTipText     =   "Choose a comment from a list"
         Top             =   5670
         Width           =   435
      End
      Begin VB.CommandButton cmdGramPrep 
         Caption         =   "&Gram Stains  Wet Prep"
         Height          =   615
         Left            =   -64440
         TabIndex        =   279
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   945
         Index           =   1
         Left            =   -64500
         Picture         =   "frmEditBacteriology.frx":103A7
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   4740
         Width           =   1035
      End
      Begin VB.Frame fraMicroResult 
         BorderStyle     =   0  'None
         Caption         =   "fraMicroResult"
         Height          =   3555
         Index           =   1
         Left            =   -74400
         TabIndex        =   228
         Top             =   600
         Width           =   10635
         Begin VB.Frame fraDipStick 
            Caption         =   "Dip Stick"
            Height          =   3165
            Left            =   8190
            TabIndex        =   259
            Top             =   210
            Width           =   2415
            Begin VB.TextBox txtUrobilinogen 
               Height          =   285
               Left            =   1050
               MaxLength       =   10
               TabIndex        =   267
               Top             =   2100
               Width           =   1200
            End
            Begin VB.TextBox txtBilirubin 
               Height          =   285
               Left            =   1050
               MaxLength       =   10
               TabIndex        =   266
               Top             =   2400
               Width           =   1200
            End
            Begin VB.TextBox txtKetones 
               Height          =   285
               Left            =   1050
               MaxLength       =   10
               TabIndex        =   265
               Top             =   1800
               Width           =   1200
            End
            Begin VB.TextBox txtGlucose 
               Height          =   285
               Left            =   1050
               MaxLength       =   10
               TabIndex        =   264
               Top             =   1500
               Width           =   1200
            End
            Begin VB.TextBox txtProtein 
               Height          =   285
               Left            =   1050
               MaxLength       =   20
               TabIndex        =   263
               Top             =   1200
               Width           =   1200
            End
            Begin VB.TextBox txtpH 
               Height          =   285
               Left            =   1050
               MaxLength       =   10
               TabIndex        =   262
               Top             =   900
               Width           =   1200
            End
            Begin VB.TextBox txtBloodHb 
               Height          =   285
               Left            =   1050
               MaxLength       =   10
               TabIndex        =   261
               Top             =   2700
               Width           =   1200
            End
            Begin VB.CommandButton cmdNAD 
               Appearance      =   0  'Flat
               Caption         =   "NAD"
               Height          =   285
               Left            =   1050
               TabIndex        =   260
               Top             =   360
               Width           =   1185
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Urobilinogen"
               Height          =   195
               Index           =   39
               Left            =   120
               TabIndex        =   274
               Top             =   2130
               Width           =   885
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Bilirubin"
               Height          =   195
               Index           =   40
               Left            =   450
               TabIndex        =   273
               Top             =   2430
               Width           =   540
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Ketones"
               Height          =   195
               Index           =   38
               Left            =   405
               TabIndex        =   272
               Top             =   1830
               Width           =   585
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Glucose"
               Height          =   195
               Index           =   37
               Left            =   405
               TabIndex        =   271
               Top             =   1530
               Width           =   585
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Protein"
               Height          =   195
               Index           =   36
               Left            =   495
               TabIndex        =   270
               Top             =   1230
               Width           =   495
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "pH"
               Height          =   195
               Index           =   35
               Left            =   780
               TabIndex        =   269
               Top             =   960
               Width           =   210
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Blood/Hb"
               Height          =   195
               Index           =   41
               Left            =   300
               TabIndex        =   268
               Top             =   2730
               Width           =   690
            End
         End
         Begin VB.Frame fraMicroscopy 
            Caption         =   "Microscopy"
            Height          =   3165
            Left            =   30
            TabIndex        =   242
            Top             =   210
            Width           =   3705
            Begin VB.TextBox txtRCC 
               Height          =   285
               Left            =   2070
               MaxLength       =   10
               TabIndex        =   251
               Top             =   1080
               Width           =   1200
            End
            Begin VB.TextBox txtWCC 
               Height          =   285
               Left            =   2070
               MaxLength       =   10
               TabIndex        =   250
               Top             =   750
               Width           =   780
            End
            Begin VB.ComboBox cmbCasts 
               Height          =   315
               Left            =   750
               TabIndex        =   249
               Text            =   "cmbCasts"
               Top             =   1710
               Width           =   2775
            End
            Begin VB.ComboBox cmbCrystals 
               Height          =   315
               Left            =   750
               TabIndex        =   248
               Text            =   "cmbCrystals"
               Top             =   1380
               Width           =   2775
            End
            Begin VB.ComboBox cmbMisc 
               Height          =   315
               Index           =   0
               Left            =   750
               TabIndex        =   247
               Text            =   "cmbMisc"
               Top             =   2040
               Width           =   2775
            End
            Begin VB.ComboBox cmbMisc 
               Height          =   315
               Index           =   1
               Left            =   750
               TabIndex        =   246
               Text            =   "cmbMisc"
               Top             =   2370
               Width           =   2775
            End
            Begin VB.ComboBox cmbMisc 
               Height          =   315
               Index           =   2
               Left            =   750
               TabIndex        =   245
               Text            =   "cmbMisc"
               Top             =   2700
               Width           =   2775
            End
            Begin VB.TextBox txtBacteria 
               Height          =   285
               Left            =   2070
               TabIndex        =   244
               Top             =   420
               Width           =   1185
            End
            Begin VB.CommandButton cmdNADMicro 
               Caption         =   "NAD"
               Height          =   345
               Left            =   270
               TabIndex        =   243
               Top             =   390
               Width           =   615
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Crystals"
               Height          =   195
               Index           =   25
               Left            =   180
               TabIndex        =   258
               Top             =   1440
               Width           =   540
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Casts"
               Height          =   195
               Index           =   26
               Left            =   330
               TabIndex        =   257
               Top             =   1770
               Width           =   390
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Misc"
               Height          =   195
               Index           =   27
               Left            =   360
               TabIndex        =   256
               Top             =   2100
               Width           =   330
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "RCC"
               Height          =   195
               Index           =   24
               Left            =   1560
               TabIndex        =   255
               Top             =   1110
               Width           =   330
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "WCC"
               Height          =   195
               Index           =   23
               Left            =   1530
               TabIndex        =   254
               Top             =   810
               Width           =   375
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Bacteria"
               Height          =   195
               Index           =   22
               Left            =   1320
               TabIndex        =   253
               Top             =   450
               Width           =   585
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "/cmm"
               Height          =   195
               Index           =   30
               Left            =   2880
               TabIndex        =   252
               Top             =   810
               Width           =   435
            End
         End
         Begin VB.Frame fraUrineSpecific 
            Caption         =   "Specific"
            Height          =   1785
            Left            =   3990
            TabIndex        =   235
            Top             =   1620
            Width           =   3855
            Begin VB.TextBox txtSG 
               Height          =   285
               Left            =   1560
               MaxLength       =   5
               TabIndex        =   238
               Top             =   750
               Width           =   2055
            End
            Begin VB.TextBox txtBenceJones 
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   237
               Top             =   450
               Width           =   2055
            End
            Begin VB.TextBox txtFatGlobules 
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   236
               Top             =   1050
               Width           =   2055
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Specific Gravity"
               Height          =   195
               Index           =   33
               Left            =   390
               TabIndex        =   241
               Top             =   780
               Width           =   1125
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Bence Jones Protein"
               Height          =   195
               Index           =   32
               Left            =   30
               TabIndex        =   240
               Top             =   480
               Width           =   1485
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Fat Globules"
               Height          =   195
               Index           =   34
               Left            =   600
               TabIndex        =   239
               Top             =   1080
               Width           =   915
            End
         End
         Begin VB.Frame fraPregnancy 
            Caption         =   "Pregnancy"
            Height          =   1245
            Left            =   3990
            TabIndex        =   229
            Top             =   210
            Width           =   3855
            Begin VB.TextBox txtPregnancy 
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   1560
               MaxLength       =   20
               MultiLine       =   -1  'True
               TabIndex        =   231
               ToolTipText     =   "P-Positive N-Negative E-Equivocal U-Unsuitable"
               Top             =   420
               Width           =   2055
            End
            Begin VB.TextBox txtHCGLevel 
               Height          =   285
               Left            =   1560
               MaxLength       =   5
               TabIndex        =   230
               Top             =   750
               Width           =   1545
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "IU/L"
               Height          =   195
               Index           =   31
               Left            =   3120
               TabIndex        =   234
               Top             =   780
               Width           =   330
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Pregnancy Test"
               Height          =   195
               Index           =   28
               Left            =   360
               TabIndex        =   233
               Top             =   450
               Width           =   1155
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "HCG Level"
               Height          =   195
               Index           =   29
               Left            =   720
               TabIndex        =   232
               Top             =   780
               Width           =   795
            End
         End
      End
      Begin VB.Frame fraMicroResult 
         BorderStyle     =   0  'None
         Height          =   3945
         Index           =   9
         Left            =   -74700
         TabIndex        =   205
         Top             =   630
         Width           =   9645
         Begin VB.TextBox txtCSFWCCDiff 
            Height          =   315
            Index           =   1
            Left            =   4230
            TabIndex        =   225
            Top             =   1320
            Width           =   2445
         End
         Begin VB.TextBox txtCSFWCCDiff 
            Height          =   315
            Index           =   0
            Left            =   4245
            TabIndex        =   224
            Top             =   930
            Width           =   2445
         End
         Begin VB.TextBox txtCSFRCC 
            Height          =   285
            Index           =   2
            Left            =   6840
            TabIndex        =   223
            Top             =   3180
            Width           =   2445
         End
         Begin VB.TextBox txtCSFRCC 
            Height          =   285
            Index           =   1
            Left            =   4245
            TabIndex        =   222
            Top             =   3180
            Width           =   2445
         End
         Begin VB.TextBox txtCSFRCC 
            Height          =   285
            Index           =   0
            Left            =   1665
            TabIndex        =   221
            Top             =   3180
            Width           =   2445
         End
         Begin VB.TextBox txtCSFWCC 
            Height          =   285
            Index           =   2
            Left            =   6855
            TabIndex        =   220
            Top             =   2820
            Width           =   2445
         End
         Begin VB.TextBox txtCSFWCC 
            Height          =   285
            Index           =   1
            Left            =   4245
            TabIndex        =   219
            Top             =   2820
            Width           =   2445
         End
         Begin VB.TextBox txtCSFWCC 
            Height          =   285
            Index           =   0
            Left            =   1665
            TabIndex        =   218
            Top             =   2820
            Width           =   2445
         End
         Begin VB.ComboBox cmbCSFAppearance 
            Height          =   315
            Index           =   2
            Left            =   6855
            TabIndex        =   217
            Top             =   2430
            Width           =   2445
         End
         Begin VB.ComboBox cmbCSFAppearance 
            Height          =   315
            Index           =   1
            Left            =   4245
            TabIndex        =   216
            Top             =   2430
            Width           =   2445
         End
         Begin VB.ComboBox cmbCSFAppearance 
            Height          =   315
            Index           =   0
            Left            =   1665
            TabIndex        =   215
            Top             =   2430
            Width           =   2445
         End
         Begin VB.ComboBox cmbCSFGram 
            Height          =   315
            Left            =   4245
            TabIndex        =   214
            Top             =   540
            Width           =   2445
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "% Mononuclear Cells"
            Height          =   195
            Index           =   128
            Left            =   6810
            TabIndex        =   227
            Top             =   1380
            Width           =   1470
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "% Neutrophils"
            Height          =   195
            Index           =   127
            Left            =   6810
            TabIndex        =   226
            Top             =   990
            Width           =   960
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sample 3"
            Height          =   195
            Index           =   126
            Left            =   6885
            TabIndex        =   213
            Top             =   2190
            Width           =   660
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sample 2"
            Height          =   195
            Index           =   125
            Left            =   4275
            TabIndex        =   212
            Top             =   2190
            Width           =   660
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sample 1"
            Height          =   195
            Index           =   124
            Left            =   1665
            TabIndex        =   211
            Top             =   2190
            Width           =   660
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "RCC/cmm"
            Height          =   195
            Index           =   123
            Left            =   795
            TabIndex        =   210
            Top             =   3210
            Width           =   735
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "WCC/cmm"
            Height          =   195
            Index           =   122
            Left            =   750
            TabIndex        =   209
            Top             =   2850
            Width           =   780
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Appearance"
            Height          =   195
            Index           =   121
            Left            =   660
            TabIndex        =   208
            Top             =   2490
            Width           =   870
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "White Cell Differential"
            Height          =   195
            Index           =   120
            Left            =   2610
            TabIndex        =   207
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   119
            Left            =   3345
            TabIndex        =   206
            Top             =   600
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   1035
         Index           =   9
         Left            =   -64980
         Picture         =   "frmEditBacteriology.frx":107E9
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   3060
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3315
         Index           =   4
         Left            =   -66180
         TabIndex        =   198
         Top             =   1650
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator|<Code"
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
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3315
         Index           =   3
         Left            =   -68970
         TabIndex        =   199
         Top             =   1650
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator|<Code"
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
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3315
         Index           =   2
         Left            =   -71820
         TabIndex        =   200
         Top             =   1650
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator|<Code"
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
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3315
         Index           =   1
         Left            =   -74550
         TabIndex        =   201
         Top             =   1650
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator|<Code"
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
      Begin VB.Frame fraMicroResult 
         Caption         =   "RSV"
         Height          =   1275
         Index           =   8
         Left            =   -71580
         TabIndex        =   195
         Top             =   2250
         Width           =   3615
         Begin VB.Label lblRSV 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   780
            TabIndex        =   196
            Top             =   450
            Width           =   2205
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   8
         Left            =   -67770
         Picture         =   "frmEditBacteriology.frx":10C2B
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   2550
         Width           =   1425
      End
      Begin VB.Frame fraMicroResult 
         Caption         =   "Reducing Substances"
         Height          =   2475
         Index           =   7
         Left            =   -70920
         TabIndex        =   186
         Top             =   1620
         Width           =   2175
         Begin VB.CheckBox chkRS 
            Caption         =   "0 %"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   192
            Top             =   450
            Width           =   555
         End
         Begin VB.CheckBox chkRS 
            Caption         =   "0.25 %"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   191
            Top             =   750
            Width           =   795
         End
         Begin VB.CheckBox chkRS 
            Caption         =   "0.5 %"
            Height          =   195
            Index           =   2
            Left            =   720
            TabIndex        =   190
            Top             =   1050
            Width           =   705
         End
         Begin VB.CheckBox chkRS 
            Caption         =   "0.75 %"
            Height          =   195
            Index           =   3
            Left            =   720
            TabIndex        =   189
            Top             =   1350
            Width           =   795
         End
         Begin VB.CheckBox chkRS 
            Caption         =   "1 %"
            Height          =   195
            Index           =   4
            Left            =   720
            TabIndex        =   188
            Top             =   1650
            Width           =   555
         End
         Begin VB.CheckBox chkRS 
            Caption         =   "2 %"
            Height          =   195
            Index           =   5
            Left            =   720
            TabIndex        =   187
            Top             =   1950
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Un&lock Result"
         Height          =   795
         Index           =   7
         Left            =   -68280
         Picture         =   "frmEditBacteriology.frx":1106D
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   3300
         Width           =   1425
      End
      Begin VB.Frame fraMicroResult 
         Caption         =   "Rota/Adeno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Index           =   6
         Left            =   -71820
         TabIndex        =   179
         Top             =   1710
         Width           =   3735
         Begin VB.TextBox txtRota 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1500
            TabIndex        =   181
            Top             =   660
            Width           =   1395
         End
         Begin VB.TextBox txtAdeno 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1500
            TabIndex        =   180
            Top             =   1440
            Width           =   1395
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Rota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   87
            Left            =   630
            TabIndex        =   183
            Top             =   690
            Width           =   525
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Adeno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   88
            Left            =   480
            TabIndex        =   182
            Top             =   1440
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   6
         Left            =   -67920
         Picture         =   "frmEditBacteriology.frx":114AF
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   3210
         Width           =   1425
      End
      Begin VB.Frame fraMicroResult 
         Caption         =   "Occult Blood"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Index           =   5
         Left            =   -71670
         TabIndex        =   170
         Top             =   1680
         Width           =   4245
         Begin VB.CheckBox chkFOB 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   173
            Top             =   510
            Width           =   405
         End
         Begin VB.CheckBox chkFOB 
            Alignment       =   1  'Right Justify
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   172
            Top             =   1020
            Width           =   405
         End
         Begin VB.CheckBox chkFOB 
            Alignment       =   1  'Right Justify
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   660
            TabIndex        =   171
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label lblFOB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1350
            TabIndex        =   176
            Top             =   480
            Width           =   2025
         End
         Begin VB.Label lblFOB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1350
            TabIndex        =   175
            Top             =   1020
            Width           =   2025
         End
         Begin VB.Label lblFOB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   1350
            TabIndex        =   174
            Top             =   1560
            Width           =   2025
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   5
         Left            =   -67230
         Picture         =   "frmEditBacteriology.frx":118F1
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   3150
         Width           =   1425
      End
      Begin VB.Frame fraMicroResult 
         BorderStyle     =   0  'None
         Caption         =   "fraCS"
         Height          =   4995
         Index           =   4
         Left            =   -74940
         TabIndex        =   128
         Top             =   360
         Width           =   11235
         Begin VB.ComboBox cmbOrgGroup 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   3
            Left            =   6270
            TabIndex        =   149
            Text            =   "cmbOrgGroup"
            Top             =   330
            Width           =   2085
         End
         Begin VB.ComboBox cmbOrgGroup 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   2
            Left            =   3420
            TabIndex        =   148
            Text            =   "cmbOrgGroup"
            Top             =   330
            Width           =   2085
         End
         Begin VB.ComboBox cmbOrgGroup 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   1
            Left            =   690
            TabIndex        =   147
            Text            =   "cmbOrgGroup"
            Top             =   330
            Width           =   2085
         End
         Begin VB.ComboBox cmbOrgGroup 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   4
            Left            =   9060
            TabIndex        =   146
            Text            =   "cmbOrgGroup"
            Top             =   330
            Width           =   2085
         End
         Begin VB.ComboBox cmbOrgName 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   1
            Left            =   420
            TabIndex        =   145
            Text            =   "cmbOrgName"
            Top             =   660
            Width           =   2355
         End
         Begin VB.ComboBox cmbOrgName 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   2
            Left            =   3150
            TabIndex        =   144
            Text            =   "cmbOrgName"
            Top             =   660
            Width           =   2355
         End
         Begin VB.ComboBox cmbOrgName 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   3
            Left            =   6000
            TabIndex        =   143
            Text            =   "cmbOrgName"
            Top             =   660
            Width           =   2355
         End
         Begin VB.ComboBox cmbOrgName 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   4
            Left            =   8790
            TabIndex        =   142
            Text            =   "cmbOrgName"
            Top             =   660
            Width           =   2355
         End
         Begin VB.ComboBox cmbABSelect 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   1
            Left            =   420
            TabIndex        =   141
            Text            =   "cmbABSelect"
            Top             =   4620
            Width           =   2355
         End
         Begin VB.ComboBox cmbABSelect 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   2
            Left            =   3120
            TabIndex        =   140
            Text            =   "cmbABSelect"
            Top             =   4620
            Width           =   2385
         End
         Begin VB.ComboBox cmbABSelect 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   3
            Left            =   6000
            TabIndex        =   139
            Text            =   "cmbABSelect"
            Top             =   4620
            Width           =   2355
         End
         Begin VB.ComboBox cmbABSelect 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   4
            Left            =   8790
            TabIndex        =   138
            Text            =   "cmbABSelect"
            Top             =   4620
            Width           =   2355
         End
         Begin VB.CommandButton cmdRemoveSecondary 
            Height          =   525
            Index           =   1
            Left            =   0
            Picture         =   "frmEditBacteriology.frx":11D33
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Remove Secondary Lists"
            Top             =   2580
            Width           =   375
         End
         Begin VB.CommandButton cmdRemoveSecondary 
            Height          =   525
            Index           =   2
            Left            =   2790
            Picture         =   "frmEditBacteriology.frx":1203D
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Remove Secondary Lists"
            Top             =   2580
            Width           =   375
         End
         Begin VB.CommandButton cmdRemoveSecondary 
            Height          =   525
            Index           =   3
            Left            =   5580
            Picture         =   "frmEditBacteriology.frx":12347
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Remove Secondary Lists"
            Top             =   2580
            Width           =   375
         End
         Begin VB.CommandButton cmdRemoveSecondary 
            Height          =   525
            Index           =   4
            Left            =   8370
            Picture         =   "frmEditBacteriology.frx":12651
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Remove Secondary Lists"
            Top             =   2580
            Width           =   375
         End
         Begin VB.ComboBox cmbQualifier 
            Height          =   315
            Index           =   1
            Left            =   420
            TabIndex        =   133
            Text            =   "cmbQualifier"
            Top             =   990
            Width           =   2355
         End
         Begin VB.ComboBox cmbQualifier 
            Height          =   315
            Index           =   2
            Left            =   3150
            TabIndex        =   132
            Text            =   "cmbQualifier"
            Top             =   990
            Width           =   2355
         End
         Begin VB.ComboBox cmbQualifier 
            Height          =   315
            Index           =   3
            Left            =   6000
            TabIndex        =   131
            Text            =   "cmbQualifier"
            Top             =   990
            Width           =   2355
         End
         Begin VB.ComboBox cmbQualifier 
            Height          =   315
            Index           =   4
            Left            =   8790
            TabIndex        =   130
            Text            =   "cmbQualifier"
            Top             =   990
            Width           =   2355
         End
         Begin VB.CommandButton cmdCopySensitivities 
            BackColor       =   &H00FF00FF&
            Caption         =   "Copy from Previous"
            Enabled         =   0   'False
            Height          =   315
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   0
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Label lblSetAllR 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   161
            ToolTipText     =   "Set All Resistant"
            Top             =   3660
            Width           =   270
         End
         Begin VB.Label lblSetAllS 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   160
            ToolTipText     =   "Set All Sensitive"
            Top             =   4050
            Width           =   255
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   79
            Left            =   420
            TabIndex        =   159
            Top             =   330
            Width           =   270
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   80
            Left            =   3150
            TabIndex        =   158
            Top             =   330
            Width           =   270
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   81
            Left            =   6000
            TabIndex        =   157
            Top             =   330
            Width           =   270
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   82
            Left            =   8790
            TabIndex        =   156
            Top             =   330
            Width           =   270
         End
         Begin VB.Label lblSetAllS 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   2910
            TabIndex        =   155
            ToolTipText     =   "Set All Sensitive"
            Top             =   4050
            Width           =   255
         End
         Begin VB.Label lblSetAllR 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   2910
            TabIndex        =   154
            ToolTipText     =   "Set All Resistant"
            Top             =   3660
            Width           =   270
         End
         Begin VB.Label lblSetAllS 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   5700
            TabIndex        =   153
            ToolTipText     =   "Set All Sensitive"
            Top             =   4050
            Width           =   255
         End
         Begin VB.Label lblSetAllR 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   5700
            TabIndex        =   152
            ToolTipText     =   "Set All Resistant"
            Top             =   3660
            Width           =   270
         End
         Begin VB.Label lblSetAllS 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   8490
            TabIndex        =   151
            ToolTipText     =   "Set All Sensitive"
            Top             =   4050
            Width           =   255
         End
         Begin VB.Label lblSetAllR 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   8490
            TabIndex        =   150
            ToolTipText     =   "Set All Resistant"
            Top             =   3660
            Width           =   270
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   1005
         Index           =   4
         Left            =   -63510
         Picture         =   "frmEditBacteriology.frx":1295B
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   3300
         Width           =   705
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   10
         Left            =   -65880
         Picture         =   "frmEditBacteriology.frx":12D9D
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   3840
         Width           =   1425
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   11
         Left            =   -65490
         Picture         =   "frmEditBacteriology.frx":131DF
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   4140
         Width           =   1425
      End
      Begin MSFlexGridLib.MSFlexGrid grdDay 
         Height          =   1845
         Index           =   1
         Left            =   -67890
         TabIndex        =   115
         Top             =   810
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3254
         _Version        =   393216
         Cols            =   5
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "<Date/Time         |<XLD               |<DCA            |<SMAC           |<Technician      "
      End
      Begin ComCtl2.UpDown udHistoricalFaecesView 
         Height          =   285
         Left            =   -64860
         TabIndex        =   114
         Top             =   450
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblViewOrganism"
         BuddyDispid     =   196785
         OrigLeft        =   10290
         OrigTop         =   780
         OrigRight       =   10695
         OrigBottom      =   1020
         Max             =   3
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 4"
         Height          =   5445
         Index           =   3
         Left            =   -67020
         TabIndex        =   109
         Top             =   630
         Width           =   2505
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   111
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   4
            Left            =   60
            MaxLength       =   499
            MultiLine       =   -1  'True
            TabIndex        =   110
            Top             =   600
            Width           =   2355
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 3"
         Height          =   5445
         Index           =   2
         Left            =   -69600
         TabIndex        =   106
         Top             =   630
         Width           =   2505
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   3
            Left            =   60
            MaxLength       =   499
            MultiLine       =   -1  'True
            TabIndex        =   108
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   107
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 2"
         Height          =   5445
         Index           =   1
         Left            =   -72180
         TabIndex        =   103
         Top             =   630
         Width           =   2505
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   2
            Left            =   60
            MaxLength       =   499
            MultiLine       =   -1  'True
            TabIndex        =   105
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   104
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 1"
         Height          =   5445
         Index           =   0
         Left            =   -74790
         TabIndex        =   100
         Top             =   630
         Width           =   2505
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   102
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   1
            Left            =   60
            MaxLength       =   499
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   600
            Width           =   2355
         End
      End
      Begin VB.Frame fraMicroResult 
         Caption         =   "Ova / Parasites"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Index           =   11
         Left            =   -72270
         TabIndex        =   94
         Top             =   1740
         Width           =   6195
         Begin VB.ComboBox cmbOva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   510
            TabIndex        =   99
            Top             =   2640
            Width           =   5055
         End
         Begin VB.ComboBox cmbOva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   510
            TabIndex        =   98
            Top             =   2070
            Width           =   5055
         End
         Begin VB.ComboBox cmbOva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   510
            TabIndex        =   97
            Top             =   1500
            Width           =   5055
         End
         Begin VB.Label lblCrypto 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2730
            TabIndex        =   96
            Top             =   780
            Width           =   2805
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Cryptosporidium"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   43
            Left            =   540
            TabIndex        =   95
            Top             =   810
            Width           =   1710
         End
      End
      Begin VB.Frame fraMicroResult 
         Caption         =   "C.diff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   10
         Left            =   -72990
         TabIndex        =   89
         Top             =   1170
         Width           =   5925
         Begin VB.Label lblcDiffPCR 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2220
            TabIndex        =   281
            Top             =   2340
            Width           =   3225
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "C. difficile PCR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   129
            Left            =   330
            TabIndex        =   280
            Top             =   2370
            Width           =   1560
         End
         Begin VB.Label lblToxinB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2220
            TabIndex        =   93
            Top             =   1560
            Width           =   3225
         End
         Begin VB.Label lblToxinA 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2220
            TabIndex        =   92
            Top             =   780
            Width           =   3225
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Toxin B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   90
            Left            =   1110
            TabIndex        =   91
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Toxin A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   89
            Left            =   1110
            TabIndex        =   90
            Top             =   780
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdCopyFromPrevious 
         BackColor       =   &H00FF80FF&
         Caption         =   "Copy all Details from Sample # 123456789"
         Height          =   405
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   285
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.CheckBox chkPregnant 
         Alignment       =   1  'Right Justify
         Caption         =   "Pregnant"
         Height          =   225
         Left            =   4350
         TabIndex        =   78
         Top             =   900
         Width           =   945
      End
      Begin VB.Frame Frame14 
         Caption         =   "Clinical Details"
         Height          =   1815
         Left            =   6570
         TabIndex        =   70
         Top             =   4440
         Width           =   4515
         Begin VB.TextBox txtClinDetails 
            Height          =   1095
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   600
            Width           =   4245
         End
         Begin VB.ComboBox cmbClinDetails 
            Height          =   315
            Left            =   150
            Sorted          =   -1  'True
            TabIndex        =   71
            Top             =   270
            Width           =   4245
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Patients Current Antibiotics"
         Height          =   1035
         Left            =   6570
         TabIndex        =   65
         Top             =   3180
         Width           =   4335
         Begin VB.CommandButton cmdABsInUse 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3870
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   420
            Width           =   375
         End
         Begin VB.ListBox lstABsInUse 
            Height          =   735
            IntegralHeight  =   0   'False
            ItemData        =   "frmEditBacteriology.frx":13621
            Left            =   150
            List            =   "frmEditBacteriology.frx":13623
            TabIndex        =   67
            ToolTipText     =   "Click to remove entry"
            Top             =   240
            Width           =   3675
         End
         Begin VB.ComboBox cmbABsInUse 
            Height          =   315
            Left            =   150
            TabIndex        =   66
            Text            =   "cmbABsInUse"
            Top             =   420
            Visible         =   0   'False
            Width           =   3675
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Site"
         Height          =   765
         Left            =   6570
         TabIndex        =   61
         Top             =   2250
         Width           =   4515
         Begin VB.TextBox txtSiteDetails 
            Height          =   315
            Left            =   1920
            TabIndex        =   63
            Top             =   270
            Width           =   2505
         End
         Begin VB.ComboBox cmbSite 
            Height          =   315
            Left            =   180
            TabIndex        =   62
            Text            =   "cmbSite"
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Site Details"
            Height          =   195
            Index           =   17
            Left            =   2190
            TabIndex        =   64
            Top             =   30
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdOrderTests 
         Caption         =   "Order Tests"
         Height          =   915
         Left            =   11160
         Picture         =   "frmEditBacteriology.frx":13625
         Style           =   1  'Graphical
         TabIndex        =   56
         Tag             =   "bOrder"
         Top             =   3390
         Width           =   915
      End
      Begin VB.CommandButton cmdSaveInc 
         Caption         =   "&Save"
         Height          =   915
         Left            =   11160
         Picture         =   "frmEditBacteriology.frx":1392F
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5340
         Width           =   915
      End
      Begin VB.Frame Frame4 
         Height          =   5655
         Left            =   330
         TabIndex        =   36
         Top             =   600
         Width           =   5265
         Begin VB.TextBox txtExtSampleID 
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   303
            Top             =   300
            Width           =   1065
         End
         Begin VB.CommandButton cmdCopyTo 
            BackColor       =   &H008080FF&
            Caption         =   "++ cc ++"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Copy To"
            Top             =   2160
            Width           =   375
         End
         Begin VB.CheckBox chkPenicillin 
            Alignment       =   1  'Right Justify
            Caption         =   "Penicillin Allergy"
            Height          =   225
            Left            =   3540
            TabIndex        =   79
            Top             =   540
            Width           =   1425
         End
         Begin VB.ComboBox cmbHospital 
            Height          =   315
            Left            =   900
            TabIndex        =   13
            Text            =   "cmbHospital"
            Top             =   2160
            Width           =   3915
         End
         Begin VB.ComboBox cmbDemogComment 
            Height          =   315
            Left            =   900
            TabIndex        =   60
            Top             =   3570
            Width           =   3915
         End
         Begin VB.ComboBox cmbGP 
            Height          =   315
            Left            =   900
            TabIndex        =   8
            Text            =   "cmbGP"
            Top             =   3210
            Width           =   3915
         End
         Begin VB.ComboBox cmbClinician 
            Height          =   315
            Left            =   900
            TabIndex        =   7
            Text            =   "cmbClinician"
            Top             =   2850
            Width           =   3915
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   1
            Left            =   750
            MaxLength       =   30
            TabIndex        =   12
            Top             =   1770
            Width           =   4215
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   0
            Left            =   750
            MaxLength       =   30
            TabIndex        =   11
            Top             =   1500
            Width           =   4215
         End
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   900
            TabIndex        =   6
            Text            =   "cmbWard"
            Top             =   2490
            Width           =   3915
         End
         Begin VB.TextBox txtDemographicComment 
            Height          =   1515
            Left            =   900
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   3870
            Width           =   3885
         End
         Begin VB.Label lblExtSampleID 
            AutoSize        =   -1  'True
            Caption         =   "Ext. SampleID"
            Height          =   195
            Left            =   2400
            TabIndex        =   304
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Hospital"
            Height          =   195
            Index           =   10
            Left            =   270
            TabIndex        =   75
            Top             =   2220
            Width           =   570
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "GP"
            Height          =   195
            Index           =   13
            Left            =   630
            TabIndex        =   51
            Top             =   3270
            Width           =   225
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Clinician"
            Height          =   195
            Index           =   12
            Left            =   255
            TabIndex        =   50
            Top             =   2880
            Width           =   585
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   49
            Top             =   3630
            Width           =   735
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   48
            Top             =   1530
            Width           =   570
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Ward"
            Height          =   195
            Index           =   11
            Left            =   450
            TabIndex        =   47
            Top             =   2550
            Width           =   390
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            Height          =   195
            Index           =   61
            Left            =   3930
            TabIndex        =   46
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            Height          =   195
            Index           =   60
            Left            =   2760
            TabIndex        =   45
            Top             =   1200
            Width           =   285
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "D.o.B"
            Height          =   195
            Index           =   8
            Left            =   210
            TabIndex        =   44
            Top             =   1230
            Width           =   405
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Index           =   7
            Left            =   210
            TabIndex        =   43
            Top             =   810
            Width           =   420
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Chart #"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   330
            Width           =   525
         End
         Begin VB.Label lblChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   41
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lblName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   750
            TabIndex        =   40
            Top             =   780
            Width           =   4215
         End
         Begin VB.Label lblDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   39
            Top             =   1170
            Width           =   1515
         End
         Begin VB.Label lblAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3180
            TabIndex        =   38
            Top             =   1170
            Width           =   585
         End
         Begin VB.Label lblSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4260
            TabIndex        =   37
            Top             =   1170
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdSaveDemographics 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   915
         Left            =   11160
         Picture         =   "frmEditBacteriology.frx":13F99
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4365
         Width           =   915
      End
      Begin VB.Frame Frame7 
         Caption         =   "Run Date"
         Height          =   1725
         Left            =   6570
         TabIndex        =   23
         Top             =   420
         Width           =   5535
         Begin VB.Frame Frame5 
            Height          =   795
            Left            =   4170
            TabIndex        =   276
            Top             =   0
            Width           =   1365
            Begin VB.OptionButton cRooH 
               Alignment       =   1  'Right Justify
               Caption         =   "Routine"
               Height          =   195
               Index           =   0
               Left            =   420
               TabIndex        =   278
               Top             =   240
               Width           =   885
            End
            Begin VB.OptionButton cRooH 
               Alignment       =   1  'Right Justify
               Caption         =   "Out of Hours"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   277
               Top             =   480
               Width           =   1215
            End
         End
         Begin MSComCtl2.DTPicker dtRecDate 
            Height          =   315
            Left            =   1890
            TabIndex        =   81
            Top             =   1050
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   217186305
            CurrentDate     =   38078
         End
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   150
            TabIndex        =   24
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   217186305
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   1890
            TabIndex        =   25
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   217186305
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   3270
            TabIndex        =   26
            ToolTipText     =   "Time of Sample"
            Top             =   270
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox tRecTime 
            Height          =   315
            Left            =   3270
            TabIndex        =   119
            ToolTipText     =   "Time of Sample"
            Top             =   1050
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDateError 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date Sequence Error"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   675
            Left            =   4410
            TabIndex        =   275
            Top             =   1050
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   2760
            Picture         =   "frmEditBacteriology.frx":143DB
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   1890
            Picture         =   "frmEditBacteriology.frx":1481D
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   2370
            Picture         =   "frmEditBacteriology.frx":14C5F
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1380
            Width           =   360
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Received in Lab"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   80
            Top             =   1110
            Width           =   1770
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   2370
            Picture         =   "frmEditBacteriology.frx":150A1
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   630
            Picture         =   "frmEditBacteriology.frx":154E3
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   2760
            Picture         =   "frmEditBacteriology.frx":15925
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   1890
            Picture         =   "frmEditBacteriology.frx":15D67
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   1020
            Picture         =   "frmEditBacteriology.frx":161A9
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   120
            Picture         =   "frmEditBacteriology.frx":165EB
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Sample Date"
            Height          =   195
            Index           =   15
            Left            =   1920
            TabIndex        =   27
            Top             =   0
            Width           =   1155
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdDay 
         Height          =   1845
         Index           =   2
         Left            =   -67890
         TabIndex        =   116
         Top             =   2640
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3254
         _Version        =   393216
         Cols            =   5
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "<Date/Time         |<XLD SUB      |<DCA SUB     |<PRESTON   |<Technician      "
      End
      Begin MSFlexGridLib.MSFlexGrid grdDay 
         Height          =   1845
         Index           =   3
         Left            =   -67890
         TabIndex        =   117
         Top             =   4470
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3254
         _Version        =   393216
         Cols            =   3
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "<Date/Time         |<CCDA                                                       |<Technician      "
      End
      Begin VB.Label lblRequestID 
         Caption         =   "Request ID : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   330
         TabIndex        =   118
         Top             =   330
         Width           =   3045
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Urine Specimen Comment"
         Height          =   195
         Index           =   1
         Left            =   -73620
         TabIndex        =   291
         Top             =   4320
         Width           =   1830
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   1
         Left            =   -64500
         TabIndex        =   122
         ToolTipText     =   "Sample has been Printed"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   390
         Index           =   9
         Left            =   -64980
         TabIndex        =   204
         ToolTipText     =   "Sample has been Printed"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   8
         Left            =   -66210
         TabIndex        =   197
         ToolTipText     =   "Sample has been Printed"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   7
         Left            =   -68280
         TabIndex        =   193
         ToolTipText     =   "Sample has been Printed"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   6
         Left            =   -67920
         TabIndex        =   184
         ToolTipText     =   "Sample has been Printed"
         Top             =   2340
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   5
         Left            =   -67260
         TabIndex        =   177
         ToolTipText     =   "Sample has been Printed"
         Top             =   2100
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P r i n t e d"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   1440
         Index           =   4
         Left            =   -63240
         TabIndex        =   127
         ToolTipText     =   "Sample has been Printed"
         Top             =   1710
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   10
         Left            =   -65880
         TabIndex        =   121
         ToolTipText     =   "Sample has been Printed"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   11
         Left            =   -65490
         TabIndex        =   120
         ToolTipText     =   "Sample has been Printed"
         Top             =   3270
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblViewOrganism 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65250
         TabIndex        =   113
         Top             =   480
         Width           =   345
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Historical View of Organism"
         Height          =   195
         Index           =   102
         Left            =   -67200
         TabIndex        =   112
         Top             =   480
         Width           =   1920
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   -63870
         Picture         =   "frmEditBacteriology.frx":16A2D
         Top             =   420
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareCross 
         Height          =   225
         Left            =   -63660
         Picture         =   "frmEditBacteriology.frx":16D03
         Top             =   420
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Image imgHRed 
      Height          =   510
      Left            =   16560
      Picture         =   "frmEditBacteriology.frx":16FD9
      Top             =   4410
      Width           =   480
   End
   Begin VB.Image imgHGreen 
      Height          =   510
      Left            =   16560
      Picture         =   "frmEditBacteriology.frx":17CDB
      Top             =   3810
      Width           =   480
   End
   Begin VB.Label lblNOPAS 
      AutoSize        =   -1  'True
      Caption         =   "NOPAS"
      Height          =   195
      Left            =   16290
      TabIndex        =   84
      Top             =   1590
      Width           =   555
   End
   Begin VB.Label lblAandE 
      Caption         =   "A and E"
      Height          =   225
      Left            =   16140
      TabIndex        =   83
      Top             =   690
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Begin VB.Menu mnuBacteria 
         Caption         =   "&Bacteria"
      End
      Begin VB.Menu mnuWCC 
         Caption         =   "&WCC"
      End
      Begin VB.Menu mnuRCC 
         Caption         =   "&RCC"
      End
      Begin VB.Menu mnuOrganisms 
         Caption         =   "&Organisms"
      End
      Begin VB.Menu mnuCSF 
         Caption         =   "&CSF"
         Begin VB.Menu mnuCSFGram 
            Caption         =   "&Gram Stain"
         End
         Begin VB.Menu mnuCSFAppearance 
            Caption         =   "&Appearance"
         End
      End
      Begin VB.Menu mnuMicroLists 
         Caption         =   "&Microbiology"
         Begin VB.Menu mnuDefaultsMicro 
            Caption         =   "&Microbiology"
         End
         Begin VB.Menu mnuMicroIdent 
            Caption         =   "&Identificaton"
            Begin VB.Menu mnuMicroIDGram 
               Caption         =   "&Gram Stans"
            End
            Begin VB.Menu mnuMicroWetPrep 
               Caption         =   "&Wet Prep"
            End
            Begin VB.Menu mnuMicroGWQuantity 
               Caption         =   "&Quantity"
            End
         End
         Begin VB.Menu mnuUrineLists 
            Caption         =   "&Urine"
            Begin VB.Menu mnuUrineCrystals 
               Caption         =   "Cr&ystals"
            End
            Begin VB.Menu mnuUrineCasts 
               Caption         =   "&casts"
            End
            Begin VB.Menu mnuUrineMisc 
               Caption         =   "&Miscellaneous"
            End
         End
         Begin VB.Menu mnuListFaeces 
            Caption         =   "&Faeces"
            Begin VB.Menu mnuListXLDDCA 
               Caption         =   "&XLD/DCA"
            End
            Begin VB.Menu mnuListSMAC 
               Caption         =   "&SMAC"
            End
            Begin VB.Menu mnuListPrestonCCDA 
               Caption         =   "&Preston/CCDA"
            End
         End
      End
   End
End
Attribute VB_Name = "frmEditMicrobiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ComingFromUnlock As Boolean

Private pFromViewReportSID As String

Private mNewRecord As Boolean

Private Activated As Boolean

Private pPrintToPrinter As String

Private IdentLoaded As Boolean
Private FaecesLoaded As Boolean
Private CSLoaded As Boolean
Private FOBLoaded As Boolean
Private RotaAdenoLoaded As Boolean
Private CdiffLoaded As Boolean
Private OPLoaded As Boolean
Private IdentificationLoaded As Boolean
Private CSFLoaded As Boolean

Private SampleIDWithOffset As Long

Private frmOptUrineSpecific As Boolean
Private MatchingDemoLoaded As Boolean

Dim ListBacteria() As String
Dim ListWCC() As String
Dim ListRCC() As String
Dim ListOrganism() As String

'+++Junaid 15-10-2023
Dim m_Time As Integer
Dim m_Flag As Boolean
'---Junaid

'Private ForceSaveability As Boolean

Private Enum Mic
    Demographics = 0
    Urine = 1
    Identification = 2
    Faeces = 3
    CandS = 4
    FOB = 5
    RotaAdeno = 6
    RedSub = 7
    RSV = 8
    CSF = 9
    CDiff = 10
    OP = 11
    IDENTCAVAN = 12
End Enum

Private CurrentSensitivities As Sensitivities

Private LoadingAllDetails As Boolean

Private UrineAutoVal As String

Private ClearingUrine As Boolean
Private FormLoaded As Boolean

Private Const fcsLine_NO = 0
Private Const fcsSr = 1
Private Const fcsQes = 2
Private Const fcsAns = 3





Private Sub btnAdd_Click()
34360     If lstAntibio.Text = "" Then
34370         Exit Sub
34380     End If
34390     If fmeAntibio.Caption = "Antibiotics" Then
34400         If txtAntibiotics.Text = "" Then
34410             txtAntibiotics.Text = Trim(lstAntibio.Text)
34420         Else
34430             txtAntibiotics.Text = Trim(txtAntibiotics.Text) & ", " & Trim(lstAntibio.Text)
34440         End If
34450     ElseIf fmeAntibio.Caption = "Intended Antibiotics" Then
34460         If txtIntAntibiotics.Text = "" Then
34470             txtIntAntibiotics.Text = Trim(lstAntibio.Text)
34480         Else
34490             txtIntAntibiotics.Text = Trim(txtIntAntibiotics.Text) & ", " & Trim(lstAntibio.Text)
34500         End If
34510     End If
End Sub

Private Sub btnAntiBiotics_Click()
34520     fmeAntibio.Visible = True
34530     fmeAntibio.Caption = "Antibiotics"
34540     Call ShowAntiBiotic
End Sub

Private Sub btnHide_Click()
34550     txtDemographicComment.SetFocus
34560     fmeAntibio.Visible = False
34570     txtDemographicComment.SetFocus
End Sub

Private Sub btnHideQuestions_Click()
34580     fmeQuestions.Visible = False
End Sub

Private Sub btnIntAntibiotics_Click()
34590     fmeAntibio.Visible = True
34600     fmeAntibio.Caption = "Intended Antibiotics"
34610     Call ShowAntiBiotic
End Sub

Private Sub btnOtherQuestions_Click()
34620     fmeQuestions.Visible = True
34630     Call FormatGrid
34640     Call GetOtherQuestion(txtSampleID.Text)
End Sub

Private Sub cmbOrgGroup_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
34650     cmbOrgGroup(Index).Text = ""
End Sub

Private Sub cmbOrgName_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
34660     cmbOrgName(Index).Text = ""
End Sub

Private Sub cmbQualifier_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
34670     cmbQualifier(Index) = ""
End Sub

Private Sub cmdOrderBiomnis_Click(Index As Integer)

          Dim sql As String
          Dim n As Integer
          Dim SampleDate As Date
          Dim LabName As String

34680     On Error GoTo cmdOrderBiomnis_Click_Error

34690     Select Case Index
              Case 0
34700             LabName = "Biomnis"
34710         Case 1
34720             LabName = "MAT: Mater Hospital"
34730     End Select

34740     If IsDate(dtSampleDate & " " & tSampleTime) Then
34750         SampleDate = dtSampleDate & " " & tSampleTime
34760     End If

34770     sql = "IF NOT EXISTS(SELECT * FROM BiomnisRequests " & _
              "                  WHERE SampleID = '" & Val(txtSampleID) & "' " & _
              "                  AND TestCode = '" & cmbSite & "' AND SendTo = '" & LabName & "' )" & _
              "      INSERT INTO BiomnisRequests (SampleID, TestCode, TestName, SampleType, SampleDateTime, Department, RequestedBy, SendTo, Status) " & _
              "      VALUES " & _
              "     ('" & Val(txtSampleID) & "', " & _
              "      '" & cmbSite & "', " & _
              "      '" & txtSiteDetails & "', " & _
              "      '" & cmbSite & "', " & _
              "      '" & Format(SampleDate, "dd/MMM/yyyy hh:mm") & "', " & _
              "      'Micro', " & _
              "      '" & UserCode & "^" & UserName & "', " & _
              "      '" & LabName & "', " & _
              "      'OutStanding')"
34780     Cnxn(0).Execute sql


34790     SaveDemographicInc


34800     Exit Sub

cmdOrderBiomnis_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34810     intEL = Erl
34820     strES = Err.Description
34830     LogError "frmEditMicrobiology", "cmdOrderBiomnis_Click", intEL, strES
          
End Sub

Private Sub cmdOrderExt_Click()
          Dim frm As New frmAddToTests

34840     On Error GoTo cmdOrderExt_Click_Error

34850     If UserHasAuthority(UserMemberOf, "MicroExtOrderTest") = False Then
34860         iMsg "You do not have authority to Order External test in Microbiology " & vbCrLf & "Please contact system administrator"
34870         Exit Sub
34880     End If

34890     If Val(txtSampleID) = 0 Then
34900         Exit Sub
34910     End If

34920     If txtSurName = "" And txtDoB = "" Then
34930         iMsg "Please provide Surname and DoB first", vbInformation
34940         Exit Sub
34950     End If

34960     SaveDemographics
34970     frm.SampleID = Format$(Val(txtSampleID))
34980     frm.Sex = txtSex
34990     If IsDate(tSampleTime) Then
35000         frm.SampleDateTime = Format$(dtSampleDate, "dd/MMM/yyyy") & " " & tSampleTime
35010     Else
35020         frm.SampleDateTime = Format$(dtSampleDate, "dd/MMM/yyyy") & " " & "00:01"
35030     End If
35040     frm.ClinicalDetails = txtClinDetails
35050     frm.Show 1

35060     Unload frm
35070     Set frm = Nothing



35080     Exit Sub

cmdOrderExt_Click_Error:

          Dim strES As String
          Dim intEL As Integer

35090     intEL = Erl
35100     strES = Err.Description
35110     LogError "frmEditMicrobiology", "cmdOrderExt_Click", intEL, strES
End Sub

Private Sub cmdScan_Click()
35120     On Error GoTo cmdScan_Click_Error

35130     With frmScan
35140         .txtSampleID = txtSampleID
35150         .Show 1
35160     End With
35170     SetViewScans Val(txtSampleID), cmdViewScan

35180     Exit Sub

cmdScan_Click_Error:

          Dim strES As String
          Dim intEL As Integer

35190     intEL = Erl
35200     strES = Err.Description
35210     LogError "frmEditAll", "cmdScan_Click", intEL, strES
End Sub

Private Sub cmdValidate_Click()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdViewLog_Click
' Author    : Masood
' Date      : 28/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewLog_Click()
35220     On Error GoTo cmdViewLog_Click_Error


          '    LoadConsultantListLog (Val(txtSampleID) + sysOptMicroOffset(0))

35230     With frmConsultantListLog
              '30           .LoadConsultantListLog (Val(txtSampleID) + sysOptMicroOffset(0))
35240         .LoadConsultantListLog (Val(txtSampleID))
35250         .Show 1
35260     End With

35270     Exit Sub


cmdViewLog_Click_Error:

          Dim strES As String
          Dim intEL As Integer

35280     intEL = Erl
35290     strES = Err.Description
35300     LogError "frmEditMicrobiology", "cmdViewLog_Click", intEL, strES
End Sub

Private Sub cmdViewScan_Click()
35310     On Error GoTo cmdViewScan_Click_Error

35320     frmViewScan.CallerDepartment = SSTab1.TabCaption(SSTab1.Tab)
35330     frmViewScan.SampleID = txtSampleID ' + sysOptMicroOffset(0)
35340     frmViewScan.txtSampleID = txtSampleID ' + sysOptMicroOffset(0)
35350     frmViewScan.Show 1
          'DoEvents
          '      DoEvents

35360     Exit Sub

cmdViewScan_Click_Error:

          Dim strES As String
          Dim intEL As Integer

35370     intEL = Erl
35380     strES = Err.Description
35390     LogError "frmEditMicrobiology", "cmdViewScan_Click", intEL, strES
          

End Sub


'---------------------------------------------------------------------------------------
' Procedure : bsearchDob_Click
' Author    : XPMUser
' Date      : 04/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bsearchDob_Click()

35400     On Error GoTo bsearchDob_Click_Error

35410     LoadMatchingDemo


       
35420     Exit Sub

       
bsearchDob_Click_Error:

          Dim strES As String
          Dim intEL As Integer

35430     intEL = Erl
35440     strES = Err.Description
35450     LogError "frmEditMicrobiology", "bsearchDob_Click", intEL, strES
End Sub

Private Sub CheckMicroUrineComment(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim i As Integer
          Dim InsertPositiveComment As Boolean
          Dim PositiveComment As String

35460     On Error GoTo CheckMicroUrineComment_Error


35470     PositiveComment = "Positive cultures must be correlated with signs and symptoms of UTI Particularly with low colony counts"
35480     If UCase$(cmbSite) = "URINE" Then
35490         For i = 1 To 4
35500             If UCase(cmbOrgGroup(i)) <> "NEGATIVE RESULTS" And cmbOrgName(i) <> "" Then
35510                 InsertPositiveComment = True
35520                 Exit For
35530             End If
          
35540         Next i

35550     End If

35560     If InsertPositiveComment Then
35570         If InStr(txtConC, PositiveComment) = 0 Then
35580             If txtConC = "Consultant Comments" Then txtConC = ""
35590             txtConC = Trim$(txtConC) & " " & PositiveComment
35600         End If
35610     End If

35620     Exit Sub

CheckMicroUrineComment_Error:

          Dim strES As String
          Dim intEL As Integer

35630     intEL = Erl
35640     strES = Err.Description
35650     LogError "modRTFMicro", "CheckMicroUrineComment", intEL, strES, sql

End Sub

Private Sub CheckUrineAutoVal()

          Dim OK As Boolean
          Dim sql As String
          Dim tb As Recordset
          Dim S As String
          Dim S1 As String
          Dim S2 As String

35660     On Error GoTo CheckUrineAutoVal_Error

35670     If ClearingUrine Or Trim$(txtWCC) = "" Then Exit Sub
          'DoEvents
          'DoEvents
          'if any isolate found, exit sub
          '+++Junaid 09-05-2024
          '30    sql = "SELECT Count(*) AS Cnt FROM Isolates WHERE SampleID = " & SampleIDWithOffset
35680     sql = "SELECT Count(*) AS Cnt FROM Isolates WHERE SampleID = " & Val(txtSampleID.Text)
          '--- Junaid
35690     Set tb = New Recordset
35700     RecOpenServer 0, tb, sql
35710     If tb!Cnt > 0 Then Exit Sub

35720     If Not IsDate(txtDoB) Or Not IsDate(dtSampleDate) Then
35730         Exit Sub
35740     End If

35750     OK = True

35760     If Trim$(txtBacteria) = "" Then
35770         OK = False
35780     End If
35790     If OK Then
35800         If IsNumeric(txtBacteria) Then
35810             If Val(txtBacteria) > 150 Then
35820                 OK = False
35830             End If
35840         ElseIf Left$(txtBacteria, 1) = ">" Then
35850             OK = False
35860         ElseIf InStr(txtBacteria, "+") > 0 Then
35870             OK = False
35880         End If
35890     End If

35900     If OK Then
35910         If Trim$(txtWCC) = "" Then
35920             OK = False
35930         End If
35940     End If

35950     If OK Then
35960         If IsNumeric(txtWCC) Then
35970             If Val(txtWCC) > 39 Then
35980                 OK = False
35990             End If
36000         ElseIf Left$(txtWCC, 1) = ">" Then
36010             OK = False
36020         End If
36030     End If

36040     If OK Then
36050         If DateDiff("m", txtDoB, dtSampleDate) < 180 Then    '15 years
36060             OK = False
36070         End If
36080     End If

36090     If OK Then
36100         sql = "SELECT Count(*) Tot FROM IncludeAutoValUrine " & _
                  "WHERE SourceType = 'Ward' " & _
                  "AND Hospital = '" & cmbHospital & "' " & _
                  "AND SourceName = '" & AddTicks(cmbWard) & "' " & _
                  "AND Include = 1"
36110         Set tb = New Recordset
36120         RecOpenServer 0, tb, sql
36130         If tb!Tot = 0 Then
36140             OK = False
36150         End If
36160     End If
36170     S1 = "This urine specimen has not met the automated CGH laboratory criteria for culture as the WCC is < 40/cmm and bacterial count is < 150 organisms/cmm (or equivalent). "
36180     S2 = "Please note that pyuria is defined as WCC >/= 10/cmm. If culture of this specimen is still considered clinically indicated, please contact the laboratory within 48 hours of this report and request culture and sensitivity testing. "
36190     S = S1 & vbCrLf & _
              "Consultant Comment: " & vbCrLf & _
              S2
          's = "This urine has been screened using an automated urinalysis instrument and results do not meet our criteria for culture. If UTI is still suspected please resubmit a MSU and indicate on the form that this is a repeat sample and culture will be performed."
36200     If OK Then
36210         If cmdTagRepeat.Caption = "Tag as Repeat" Then
36220             UrineAutoVal = "Pass"
36230             If InStr(txtUrineComment, S) = 0 Then
36240                 txtUrineComment = Trim$(txtUrineComment) & " " & S

36250                 If InStr(txtMSC, S1) = 0 Then
36260                     If txtMSC = "Medical Scientist Comments" Then txtMSC = ""
36270                     txtMSC = Trim$(txtMSC) & " " & S1
36280                 End If
36290                 If InStr(txtConC, S2) = 0 Then
36300                     If txtConC = "Consultant Comments" Then txtConC = ""
36310                     txtConC = Trim$(txtConC) & " " & S2
36320                 End If

                      Dim Iso As New Isolate
                      '+++ Junaid 20-05-2024
                      '680               Iso.SampleID = SampleIDWithOffset
36330                 Iso.SampleID = Trim(txtSampleID.Text)
                      '--- Junaid
36340                 Iso.IsolateNumber = 1
36350                 Iso.OrganismGroup = "Negative Results"
36360                 Iso.OrganismName = "Screened Urine"
36370                 Iso.Qualifier = ""
36380                 Iso.UserName = AddTicks(UserName)
36390                 Iso.RecordDateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
36400                 Iso.Save
36410                 cmbOrgGroup(1) = "Negative Results"
36420                 cmbOrgName(1) = "Screened Urine"
36430                 cmbQualifier(1) = ""

36440             End If
36450         Else
36460             txtUrineComment = Replace(txtUrineComment, S, "")
36470             txtMSC = Replace(txtMSC, S1, "")
36480             txtConC = Replace(txtConC, S2, "")
36490             UrineAutoVal = "Repeat"
36500         End If
36510     Else
36520         txtUrineComment = Replace(txtUrineComment, S, "")
36530         txtMSC = Replace(txtMSC, S1, "")
36540         txtConC = Replace(txtConC, S2, "")
36550         UrineAutoVal = "Fail"
36560     End If

36570     SaveComments False

36580     Exit Sub

CheckUrineAutoVal_Error:

          Dim strES As String
          Dim intEL As Integer

36590     intEL = Erl
36600     strES = Err.Description
36610     LogError "frmEditMicrobiology", "CheckUrineAutoVal", intEL, strES, sql

End Sub

Private Sub ClearCSF()

          Dim n As Integer

36620     cmbCSFGram = ""
36630     txtCSFWCCDiff(0) = ""
36640     txtCSFWCCDiff(1) = ""

36650     For n = 0 To 2
36660         cmbCSFAppearance(n) = ""
36670         txtCSFWCC(n) = ""
36680         txtCSFRCC(n) = ""
36690     Next

End Sub

Private Sub ClearIndividualFaeces()

          Dim n As Integer

36700     For n = 0 To 2
36710         chkFOB(n) = 0
36720         lblFOB(n) = ""
36730         lblFOB(n).BackColor = &H8000000F
36740     Next

36750     txtRota = ""
36760     txtRota.BackColor = &H8000000F
36770     txtAdeno = ""
36780     txtAdeno.BackColor = &H8000000F

36790     lblToxinA = ""
36800     lblToxinA.BackColor = &H8000000F
36810     lblToxinB = ""
36820     lblToxinB.BackColor = &H8000000F

36830     lblCrypto = ""
36840     lblCrypto.BackColor = &H8000000F
36850     For n = 0 To 2
36860         cmbOva(n) = ""
36870     Next

36880     lblcDiffPCR.Caption = ""
36890     lblcDiffPCR.BackColor = &H8000000F

End Sub

Private Sub EnableCopyFrom()

          Dim sql As String
          Dim tb As Recordset
          Dim PrevSID As Long

36900     On Error GoTo EnableCopyFrom_Error

36910     cmdCopyFromPrevious.Visible = False

36920     If sysOptAllowCopyDemographics(0) = False Then
36930         Exit Sub
36940     End If

36950     If Trim$(txtSurName) <> "" Or txtDoB <> "" Then
36960         Exit Sub
36970     End If

36980     PrevSID = Val(txtSampleID) - 1

36990     sql = "Select PatName from Demographics where " & _
              "SampleID = " & PrevSID & " " & _
              "and PatName <> '' " & _
              "and PatName is not null " & _
              "and DoB is not null"
37000     Set tb = New Recordset
37010     RecOpenServer 0, tb, sql
37020     If Not tb.EOF Then
37030         cmdCopyFromPrevious.Caption = "Copy All Details from Sample # " & _
                  Format$(PrevSID) & _
                  " Name " & tb!PatName
37040         cmdCopyFromPrevious.Visible = True
37050     End If

37060     Exit Sub

EnableCopyFrom_Error:

          Dim strES As String
          Dim intEL As Integer

37070     intEL = Erl
37080     strES = Err.Description
37090     LogError "frmEditMicrobiology", "EnableCopyFrom", intEL, strES, sql

End Sub

Private Function CheckIfValid() As Boolean

          Dim sql As String
          Dim tb As Recordset

37100     On Error GoTo CheckIfValid_Error
          '+++ Junaid 08-05-2024
          '20    sql = "Select count(Valid) as tot from PrintValidLog where " & _
           "SampleID = '" & SampleIDWithOffset & "' " & _
           "and Valid = 1"
37110     sql = "Select count(Valid) as tot from PrintValidLog where " & _
              "SampleID = '" & Trim(txtSampleID.Text) & "' " & _
              "and Valid = 1"
          '--- Junaid
37120     Set tb = New Recordset
37130     RecOpenClient 0, tb, sql

37140     CheckIfValid = tb!Tot > 0

37150     Exit Function

CheckIfValid_Error:

          Dim strES As String
          Dim intEL As Integer

37160     intEL = Erl
37170     strES = Err.Description
37180     LogError "frmEditMicrobiology", "CheckIfValid", intEL, strES, sql

End Function


Private Sub ClearUrine()

37190     ClearingUrine = True

37200     txtBacteria = ""
37210     txtPregnancy = ""
37220     txtHCGLevel = ""
37230     txtBenceJones = ""
37240     txtSG = ""
37250     txtFatGlobules = ""

37260     txtpH = ""
37270     txtProtein = ""
37280     txtGlucose = ""
37290     txtKetones = ""
37300     txtUrobilinogen = ""
37310     txtBilirubin = ""
37320     txtBloodHb = ""
37330     txtWCC = ""
37340     txtRCC = ""
37350     cmbCrystals = ""
37360     cmbCasts = ""
37370     cmbMisc(0) = ""
37380     cmbMisc(1) = ""
37390     cmbMisc(2) = ""

37400     ClearingUrine = False

End Sub


Private Sub cmbSiteEffects()

          Dim n As Integer

37410     On Error GoTo cmbSiteEffects_Error

37420     For n = 1 To 12
37430         SSTab1.TabVisible(n) = False
37440     Next
37450     SSTab1.TabVisible(Mic.CandS) = True
37460     SSTab1.TabVisible(Mic.IDENTCAVAN) = True

37470     cmdOrderTests.Enabled = False

37480     Select Case cmbSite
              Case "Faeces":
37490             OrderFaeces
37500             SSTab1.TabVisible(Mic.Faeces) = True

37510         Case "Urine":
37520             OrderUrine
37530             SSTab1.TabVisible(Mic.Urine) = True  'Urine
37540             SSTab1.TabVisible(Mic.CandS) = True
37550             cmdOrderTests.Enabled = True
37560             If LoadUrine() <> 0 Then
37570                 SSTab1.TabCaption(Mic.Urine) = "<<Urine>>"
37580             End If


37590         Case "CSF", "Cerebrospinal Fluid"
37600             SSTab1.TabVisible(Mic.CSF) = True
37610             SSTab1.TabVisible(Mic.CandS) = True

37620         Case Else:
37630             SSTab1.TabVisible(Mic.CandS) = True
37640             SSTab1.TabVisible(Mic.IDENTCAVAN) = False
37650     End Select

37660     lblSiteDetails = cmbSite.Text & " " & Trim(txtSiteDetails.Text)

37670     cmdSaveDemographics.Enabled = True
37680     cmdSaveInc.Enabled = True

37690     Exit Sub

cmbSiteEffects_Error:

          Dim strES As String
          Dim intEL As Integer

37700     intEL = Erl
37710     strES = Err.Description
37720     LogError "frmEditMicrobiology", "cmbSiteEffects", intEL, strES

End Sub

Private Sub EnableTagRepeat()

          Dim NewRecord As Boolean
          Dim sql As String
          Dim tb As Recordset

37730     NewRecord = False
          '+++ Junaid 20-05-2024
          '20    sql = "SELECT COUNT(*) Tot FROM Demographics WHERE SampleID = '" & SampleIDWithOffset & "'"
37740     sql = "SELECT COUNT(*) Tot FROM Demographics WHERE SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
37750     Set tb = New Recordset
37760     RecOpenServer 0, tb, sql
37770     NewRecord = (tb!Tot = 0)

37780     cmdTagRepeat.Visible = False
37790     If cmbSite = "Urine" And NewRecord Then
37800         cmdTagRepeat.Visible = True
37810     End If

End Sub

Private Function Expand(ByVal S As String) As String
          'CaOxm expanded to Calcium Oxalate monohydrate
          'CaOxd expanded to Calcium Oxalate dihydrate

          Dim RetVal As String

37820     If InStr(UCase$(S), "CAOXM") > 0 Then
37830         RetVal = "Calcium Oxalate monohydrate"
37840     ElseIf InStr(UCase$(S), "CAOXD") > 0 Then
37850         RetVal = "Calcium Oxalate dihydrate"
37860     ElseIf InStr(UCase$(S), "TRI") > 0 Then
37870         RetVal = "Triple Phosphate"
37880     ElseIf InStr(UCase$(S), "URI") > 0 Then
37890         RetVal = "Uric Acid"
37900     End If

37910     Expand = RetVal

End Function

Private Sub FillABSelect(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim ExcludeList As String

37920     On Error GoTo FillABSelect_Error

37930     cmbABSelect(Index).Clear

37940     ExcludeList = ""
37950     For n = 1 To grdAB(Index).Rows - 1
37960         ExcludeList = ExcludeList & _
                  "AntibioticName <> '" & grdAB(Index).TextMatrix(n, 0) & "' and "
37970     Next
37980     ExcludeList = Left$(ExcludeList, Len(ExcludeList) - 4)

37990     sql = "Select Distinct AntibioticName, ListOrder from Antibiotics where " & _
              ExcludeList & _
              "order by ListOrder"
38000     Set tb = New Recordset
38010     RecOpenClient 0, tb, sql
38020     Do While Not tb.EOF
38030         cmbABSelect(Index).AddItem Trim$(tb!AntibioticName & "")
38040         tb.MoveNext
38050     Loop
38060     FixComboWidth cmbABSelect(Index)

38070     Exit Sub

FillABSelect_Error:

          Dim strES As String
          Dim intEL As Integer

38080     intEL = Erl
38090     strES = Err.Description
38100     LogError "frmEditMicrobiology", "FillABSelect", intEL, strES, sql

End Sub

Private Sub FillCurrentABs()

          Dim tb As Recordset
          Dim sql As String

38110     On Error GoTo FillCurrentABs_Error

38120     cmbABsInUse.Clear

38130     sql = "Select distinct AntibioticName, ListOrder " & _
              "from Antibiotics " & _
              "order by ListOrder"
38140     Set tb = New Recordset
38150     RecOpenClient 0, tb, sql
38160     Do While Not tb.EOF
38170         cmbABsInUse.AddItem Trim$(tb!AntibioticName & "")
38180         tb.MoveNext
38190     Loop
38200     FixComboWidth cmbABsInUse
38210     Exit Sub

FillCurrentABs_Error:

          Dim strES As String
          Dim intEL As Integer

38220     intEL = Erl
38230     strES = Err.Description
38240     LogError "frmEditMicrobiology", "FillCurrentABs", intEL, strES, sql

End Sub

Private Sub fillFirstMisc(ByVal S As String)

          Dim n As Integer

38250     For n = 0 To 2
38260         If Trim$(cmbMisc(n)) = "" Then
38270             cmbMisc(n) = S
38280             Exit For
38290         ElseIf UCase$(Left$(cmbMisc(n), 3)) = UCase$(Left$(S, 3)) Then
38300             Exit For
38310         End If
38320     Next

End Sub

Private Sub FillForConsultantValidation()

          Dim sql As String
          Dim tb As Recordset
          Dim SID As Long

38330     On Error GoTo FillForConsultantValidation_Error

38340     cmdAddToConsultantList.Caption = "Add to Consultant List"

38350     cmbConsultantVal.Clear

38360     sql = "Select * from ConsultantList " & _
              "Order by SampleID"
38370     Set tb = New Recordset
38380     RecOpenServer 0, tb, sql
38390     Do While Not tb.EOF
38400         SID = Val(tb!SampleID) ' - sysOptMicroOffset(0)
38410         cmbConsultantVal.AddItem Format$(SID)
38420         If SID = Val(txtSampleID) Then
38430             cmdAddToConsultantList.Caption = "Remove from Consultant List"
38440         End If
38450         tb.MoveNext
38460     Loop
38470     FixComboWidth cmbConsultantVal
38480     Exit Sub

FillForConsultantValidation_Error:

          Dim strES As String
          Dim intEL As Integer

38490     intEL = Erl
38500     strES = Err.Description
38510     LogError "frmEditMicrobiology", "FillForConsultantValidation", intEL, strES, sql

End Sub

Private Sub FillHistoricalFaeces()

          Dim n As Integer
          Dim S As String
          Dim strPrevious(1 To 3) As String
          Dim XLD As String
          Dim DCA As String
          Dim SMAC As String
          Dim XLDS As String
          Dim DCAS As String
          Dim Preston As String
          Dim CCDA As String
          Dim OPSMAC As String
          Dim OPPreston As String
          Dim OPCCDA As String
          Dim DTR As String

          Dim WSs As New FaecesWorkSheets
          Dim WS As FaecesWorkSheet

38520     On Error GoTo FillHistoricalFaeces_Error

38530     For n = 1 To 3
38540         grdDay(n).Rows = 2
38550         grdDay(n).AddItem ""
38560         grdDay(n).RemoveItem 1
38570     Next

38580     WSs.Load Val(txtSampleID) '+ sysOptMicroOffset(0)
38590     If WSs.Count > 0 Then
38600         Set WS = WSs("11" & lblViewOrganism): If Not WS Is Nothing Then XLD = WS.Result: OPSMAC = WS.UserName
38610         Set WS = WSs("12" & lblViewOrganism): If Not WS Is Nothing Then DCA = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
38620         Set WS = WSs("13" & lblViewOrganism): If Not WS Is Nothing Then SMAC = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
38630         Set WS = WSs("21" & lblViewOrganism): If Not WS Is Nothing Then XLDS = WS.Result: OPPreston = WS.UserName
38640         Set WS = WSs("22" & lblViewOrganism): If Not WS Is Nothing Then DCAS = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
38650         Set WS = WSs("23" & lblViewOrganism): If Not WS Is Nothing Then Preston = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
38660         Set WS = WSs("3" & lblViewOrganism): If Not WS Is Nothing Then CCDA = WS.Result: OPCCDA = WS.UserName
              '90      Select Case lblViewOrganism.Caption
              '          Case "1":
              '100         Set WS = WSs("111"): If Not WS Is Nothing Then XLD = WS.Result: OPSMAC = WS.UserName
              '110         Set WS = WSs("121"): If Not WS Is Nothing Then DCA = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
              '120         Set WS = WSs("131"): If Not WS Is Nothing Then SMAC = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
              '130         Set WS = WSs("211"): If Not WS Is Nothing Then XLDS = WS.Result: OPPreston = WS.UserName
              '140         Set WS = WSs("221"): If Not WS Is Nothing Then DCAS = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
              '150         Set WS = WSs("231"): If Not WS Is Nothing Then Preston = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
              '160         Set WS = WSs("31"): If Not WS Is Nothing Then CCDA = WS.Result: OPCCDA = WS.UserName
              '170       Case "2":
              '180         Set WS = WSs("112"): If Not WS Is Nothing Then XLD = WS.Result: OPSMAC = WS.UserName
              '190         Set WS = WSs("122"): If Not WS Is Nothing Then DCA = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
              '200         Set WS = WSs("132"): If Not WS Is Nothing Then SMAC = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
              '210         Set WS = WSs("212"): If Not WS Is Nothing Then XLDS = WS.Result: OPPreston = WS.UserName
              '220         Set WS = WSs("222"): If Not WS Is Nothing Then DCAS = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
              '230         Set WS = WSs("232"): If Not WS Is Nothing Then Preston = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
              '240         Set WS = WSs("32"): If Not WS Is Nothing Then CCDA = WS.Result: OPCCDA = WS.UserName
              '250       Case "3":
              '260         Set WS = WSs("113"): If Not WS Is Nothing Then XLD = WS.Result: OPSMAC = WS.UserName
              '270         Set WS = WSs("123"): If Not WS Is Nothing Then DCA = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
              '280         Set WS = WSs("133"): If Not WS Is Nothing Then SMAC = WS.Result: If WS.UserName <> "" Then OPSMAC = WS.UserName
              '290         Set WS = WSs("213"): If Not WS Is Nothing Then XLDS = WS.Result: OPPreston = WS.UserName
              '300         Set WS = WSs("223"): If Not WS Is Nothing Then DCAS = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
              '310         Set WS = WSs("233"): If Not WS Is Nothing Then Preston = WS.Result: If WS.UserName <> "" Then OPPreston = WS.UserName
              '320         Set WS = WSs("33"): If Not WS Is Nothing Then CCDA = WS.Result: OPCCDA = WS.UserName
              '330     End Select

38670         S = Trim$(XLD) & vbTab & Trim$(DCA) & vbTab & Trim$(SMAC)
38680         If Len(S) > 2 Then
38690             strPrevious(1) = S
38700             S = "Current" & vbTab & S & vbTab & OPSMAC
38710             grdDay(1).AddItem S
38720         Else
38730             strPrevious(1) = vbTab & vbTab
38740         End If

38750         S = Trim$(XLDS) & vbTab & Trim$(DCAS) & vbTab & Trim$(Preston)
38760         If Len(S) > 2 Then
38770             strPrevious(2) = S
38780             S = "Current" & vbTab & S & vbTab & OPPreston
38790             grdDay(2).AddItem S
38800         Else
38810             strPrevious(2) = vbTab & vbTab
38820         End If

38830         S = Trim$(CCDA)
38840         strPrevious(3) = S
38850         If Len(S) > 0 Then
38860             S = "Current" & vbTab & S & vbTab & OPCCDA
38870             grdDay(3).AddItem S
38880         End If
38890     End If

38900     Set WSs = New FaecesWorkSheets
38910     WSs.LoadAudit Val(txtSampleID) '+ sysOptMicroOffset(0)
38920     For Each WS In WSs
38930         DTR = Format$(WS.DateTimeOfRecord, "dd/MM/yy HH:mm")
38940         Select Case WS.Dayindex
                  Case "11" & lblViewOrganism: grdDay(1).AddItem DTR & vbTab & WS.Result & vbTab & vbTab & vbTab & WS.UserName
38950             Case "12" & lblViewOrganism: grdDay(1).AddItem DTR & vbTab & vbTab & WS.Result & vbTab & vbTab & WS.UserName
38960             Case "13" & lblViewOrganism: grdDay(1).AddItem DTR & vbTab & vbTab & vbTab & WS.Result & vbTab & WS.UserName
38970             Case "21" & lblViewOrganism: grdDay(2).AddItem DTR & vbTab & WS.Result & vbTab & vbTab & vbTab & WS.UserName
38980             Case "22" & lblViewOrganism: grdDay(2).AddItem DTR & vbTab & vbTab & WS.Result & vbTab & vbTab & WS.UserName
38990             Case "23" & lblViewOrganism: grdDay(2).AddItem DTR & vbTab & vbTab & vbTab & WS.Result & vbTab & WS.UserName
39000             Case "3" & lblViewOrganism: grdDay(3).AddItem DTR & vbTab & WS.Result & vbTab & WS.UserName
39010         End Select
39020     Next

39030     For n = 1 To 3
39040         If grdDay(n).Rows > 2 Then
39050             grdDay(n).RemoveItem 1
39060         End If
39070     Next

39080     Exit Sub

FillHistoricalFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

39090     intEL = Erl
39100     strES = Err.Description
39110     LogError "frmEditMicrobiology", "FillHistoricalFaeces", intEL, strES

End Sub

Private Sub FillMSandConsultantComment()

          Dim tb As Recordset
          Dim sql As String

39120     On Error GoTo FillMSandConsultantComment_Error

39130     cmbConC.Clear
39140     cmbMSC.Clear

39150     sql = "Select Text from Lists where " & _
              "ListType = 'BA' and InUse = 1 " & _
              "ORDER BY ListOrder"
39160     Set tb = New Recordset
39170     RecOpenServer 0, tb, sql
39180     Do While Not tb.EOF
39190         cmbMSC.AddItem tb!Text & ""
39200         cmbConC.AddItem tb!Text & ""
39210         tb.MoveNext
39220     Loop
39230     FixComboWidth cmbMSC
39240     FixComboWidth cmbConC

39250     Exit Sub

FillMSandConsultantComment_Error:

          Dim strES As String
          Dim intEL As Integer

39260     intEL = Erl
39270     strES = Err.Description
39280     LogError "frmEditMicrobiology", "FillMSandConsultantComment", intEL, strES, sql

End Sub

Private Sub FillOrgNames(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String

39290     On Error GoTo FillOrgNames_Error

39300     cmbOrgName(Index).Clear

39310     sql = "Select * from Organisms where " & _
              "GroupName = '" & cmbOrgGroup(Index).Text & "' " & _
              "order by ListOrder"
39320     Set tb = New Recordset
39330     RecOpenClient 0, tb, sql
39340     Do While Not tb.EOF
39350         cmbOrgName(Index).AddItem tb!Name & ""
39360         tb.MoveNext
39370     Loop

39380     FixComboWidth cmbOrgName(Index)

39390     Exit Sub

FillOrgNames_Error:

          Dim strES As String
          Dim intEL As Integer

39400     intEL = Erl
39410     strES = Err.Description
39420     LogError "frmEditMicrobiology", "FillOrgNames", intEL, strES, sql

End Sub


Private Function GetSampleIDWithOffset() As Boolean
          'returns true if valid Sample ID
          'max long data type 2,147,483,647

          Dim RetVal As Boolean

39430     If Val(txtSampleID) > 147483647 Then
39440         txtSampleID = ""
39450         txtSampleID.BackColor = vbRed
39460         SampleIDWithOffset = 0
39470         RetVal = False
39480     Else
39490         txtSampleID.BackColor = vbWhite
39500         SampleIDWithOffset = Val(txtSampleID) '+ sysOptMicroOffset(0)
39510         RetVal = True
39520     End If

39530     GetSampleIDWithOffset = RetVal

End Function


Private Sub FillLists()

39540     FillWards cmbWard, HospName(0)
39550     FillClinicians cmbClinician, HospName(0)
39560     FillGPs cmbGP, HospName(0)
          '      DoEvents
          '      DoEvents
39570     FillCastsCrystalsMiscSite
39580     FillFaecesLists
39590     FillListCSFAppearance
39600     FillListCSFGram

End Sub


Private Function IsAnyRecordPresent(ByVal TableName As String, ByVal SampleID As Long) As Boolean

          Dim sql As String
          Dim tb As Recordset

39610     On Error GoTo IsAnyRecordPresent_Error

39620     sql = "SELECT SampleID FROM " & TableName & " WHERE SampleID = '" & SampleID & "'"
39630     Set tb = New Recordset
39640     RecOpenServer 0, tb, sql

39650     IsAnyRecordPresent = Not tb.EOF

39660     Exit Function

IsAnyRecordPresent_Error:

          Dim strES As String
          Dim intEL As Integer

39670     intEL = Erl
39680     strES = Err.Description
39690     LogError "frmEditMicrobiology", "IsAnyRecordPresent", intEL, strES, sql

End Function

Private Function IsChild() As Boolean

39700     IsChild = False

39710     If Not IsDate(txtDoB) Then Exit Function

39720     If DateDiff("yyyy", txtDoB, Now) < 15 Then
39730         IsChild = True
39740     End If

End Function

Private Function IsPregnant() As Boolean

39750     If chkPregnant = 1 Then
39760         IsPregnant = True
39770     Else
39780         IsPregnant = False
39790     End If

End Function

Private Function IsOutPatient() As Boolean

39800     IsOutPatient = False

End Function

Private Sub LoadComments()

          Dim OB As Observation
          Dim OBs As Observations

39810     On Error GoTo LoadComments_Error

39820     txtUrineComment = ""
39830     txtDemographicComment = ""
39840     txtMSC = "Medical Scientist Comments"
39850     txtConC = "Consultant Comments"

39860     If Val(txtSampleID) = 0 Then Exit Sub

39870     Set OBs = New Observations
          '80    Set OBs = OBs.Load(txtSampleID + sysOptMicroOffset(0), "MicroGeneral", "Demographic", "MicroCS", "MicroConsultant")
39880     Set OBs = OBs.Load(txtSampleID, "MicroGeneral", "Demographic", "MicroCS", "MicroConsultant")
39890     If Not OBs Is Nothing Then
39900         For Each OB In OBs
39910             Select Case UCase$(OB.Discipline)
                      Case "MICROGENERAL": txtUrineComment = Split_Comm(OB.Comment)
39920                 Case "DEMOGRAPHIC": txtDemographicComment = Split_Comm(OB.Comment)
39930                 Case "MICROCS": txtMSC = Split_Comm(OB.Comment)
39940                 Case "MICROCONSULTANT": txtConC = Split_Comm(OB.Comment)
39950             End Select
39960         Next
39970     End If

39980     If txtMSC = "" Then
39990         txtMSC = "Medical Scientist Comments"
40000     End If
40010     If txtConC = "" Then
40020         txtConC = "Consultant Comments"
40030     End If

40040     Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

40050     intEL = Erl
40060     strES = Err.Description
40070     LogError "frmEditMicrobiology", "LoadComments", intEL, strES

End Sub

Private Function LoadFaeces() As Boolean
          'Returns true if Faeces results present

          Dim WSs As New FaecesWorkSheets
          Dim WS As FaecesWorkSheet

40080     On Error GoTo LoadFaeces_Error

40090     ClearFaeces

40100     LoadFaeces = False
          '+++ Junaid 08-05-2024
          '40    WSs.Load SampleIDWithOffset
40110     WSs.Load Val(txtSampleID.Text)
          '--- Junaid
40120     If WSs.Count > 0 Then
40130         LoadFaeces = True
40140         For Each WS In WSs
40150             Select Case WS.Dayindex
                      Case "111": cmbDay1(11) = WS.Result
40160                 Case "112": cmbDay1(12) = WS.Result
40170                 Case "113": cmbDay1(13) = WS.Result
40180                 Case "121": cmbDay1(21) = WS.Result
40190                 Case "122": cmbDay1(22) = WS.Result
40200                 Case "123": cmbDay1(23) = WS.Result
40210                 Case "131": cmbDay1(31) = WS.Result
40220                 Case "132": cmbDay1(32) = WS.Result
40230                 Case "133": cmbDay1(33) = WS.Result

40240                 Case "211": cmbDay2(11) = WS.Result
40250                 Case "212": cmbDay2(12) = WS.Result
40260                 Case "213": cmbDay2(13) = WS.Result
40270                 Case "221": cmbDay2(21) = WS.Result
40280                 Case "222": cmbDay2(22) = WS.Result
40290                 Case "223": cmbDay2(23) = WS.Result
40300                 Case "231": cmbDay2(31) = WS.Result
40310                 Case "232": cmbDay2(32) = WS.Result
40320                 Case "233": cmbDay2(33) = WS.Result

40330                 Case "31": cmbDay3(1) = WS.Result
40340                 Case "32": cmbDay3(2) = WS.Result
40350                 Case "33": cmbDay3(3) = WS.Result
40360             End Select
40370         Next
40380         FillHistoricalFaeces
40390     End If

40400     Exit Function

LoadFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

40410     intEL = Erl
40420     strES = Err.Description
40430     LogError "frmEditMicrobiology", "LoadFaeces", intEL, strES

End Function

Private Sub LoadListBacteria()

          Dim sql As String
          Dim tb As Recordset

40440     On Error GoTo LoadListBacteria_Error

40450     ReDim ListBacteria(0 To 0) As String
40460     ListBacteria(0) = ""

40470     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'BB' and InUse = 1 " & _
              "ORDER BY ListOrder"
40480     Set tb = New Recordset
40490     RecOpenServer 0, tb, sql
40500     Do While Not tb.EOF
40510         ReDim Preserve ListBacteria(0 To UBound(ListBacteria) + 1)
40520         ListBacteria(UBound(ListBacteria)) = tb!Text & ""
40530         tb.MoveNext
40540     Loop

40550     Exit Sub

LoadListBacteria_Error:

          Dim strES As String
          Dim intEL As Integer

40560     intEL = Erl
40570     strES = Err.Description
40580     LogError "frmEditMicrobiology", "LoadListBacteria", intEL, strES, sql

End Sub
Private Sub FillListCSFGram()

          Dim sql As String
          Dim tb As Recordset

40590     On Error GoTo FillListCSFGram_Error

40600     cmbCSFGram.Clear

40610     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'FG' and InUse = 1 " & _
              "ORDER BY ListOrder"
40620     Set tb = New Recordset
40630     RecOpenServer 0, tb, sql
40640     Do While Not tb.EOF
40650         cmbCSFGram.AddItem tb!Text & ""
40660         tb.MoveNext
40670     Loop
40680     FixComboWidth cmbCSFGram

40690     Exit Sub

FillListCSFGram_Error:

          Dim strES As String
          Dim intEL As Integer

40700     intEL = Erl
40710     strES = Err.Description
40720     LogError "frmEditMicrobiology", "FillListCSFGram", intEL, strES, sql

End Sub

Private Sub FillListCSFAppearance()

          Dim sql As String
          Dim tb As Recordset

40730     On Error GoTo FillListCSFAppearance_Error

40740     cmbCSFAppearance(0).Clear
40750     cmbCSFAppearance(1).Clear
40760     cmbCSFAppearance(2).Clear

40770     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'FA' and InUse = 1 " & _
              "ORDER BY ListOrder"
40780     Set tb = New Recordset
40790     RecOpenServer 0, tb, sql
40800     Do While Not tb.EOF
40810         cmbCSFAppearance(0).AddItem tb!Text & ""
40820         cmbCSFAppearance(1).AddItem tb!Text & ""
40830         cmbCSFAppearance(2).AddItem tb!Text & ""
40840         tb.MoveNext
40850     Loop
40860     FixComboWidth cmbCSFAppearance(0)
40870     FixComboWidth cmbCSFAppearance(1)
40880     FixComboWidth cmbCSFAppearance(2)
40890     Exit Sub

FillListCSFAppearance_Error:

          Dim strES As String
          Dim intEL As Integer

40900     intEL = Erl
40910     strES = Err.Description
40920     LogError "frmEditMicrobiology", "FillListCSFAppearance", intEL, strES, sql


End Sub

Private Sub LoadListRCC()

          Dim sql As String
          Dim tb As Recordset

40930     On Error GoTo LoadListRCC_Error

40940     ReDim ListRCC(0 To 0) As String
40950     ListRCC(0) = ""

40960     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'RR' and InUse = 1 " & _
              "ORDER BY ListOrder"
40970     Set tb = New Recordset
40980     RecOpenServer 0, tb, sql
40990     Do While Not tb.EOF
41000         ReDim Preserve ListRCC(0 To UBound(ListRCC) + 1)
41010         ListRCC(UBound(ListRCC)) = tb!Text & ""
41020         tb.MoveNext
41030     Loop

41040     Exit Sub

LoadListRCC_Error:

          Dim strES As String
          Dim intEL As Integer

41050     intEL = Erl
41060     strES = Err.Description
41070     LogError "frmEditMicrobiology", "LoadListRCC", intEL, strES, sql

End Sub
Private Sub LoadListOrganism()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer

41080     On Error GoTo LoadListOrganism_Error

41090     For n = 1 To 4
41100         cmbIdentification(n).Clear
41110     Next

41120     ReDim ListOrganism(0 To 0) As String
41130     ListOrganism(0) = ""

41140     sql = "SELECT * FROM Lists WHERE " & _
              "ListType = 'IN' and InUse = 1 " & _
              "ORDER BY ListOrder"
41150     Set tb = New Recordset
41160     RecOpenServer 0, tb, sql
41170     Do While Not tb.EOF
41180         ReDim Preserve ListOrganism(0 To UBound(ListOrganism) + 1)
41190         ListOrganism(UBound(ListOrganism)) = tb!Text & ""
41200         For n = 1 To 4
41210             cmbIdentification(n).AddItem tb!Text & ""
41220         Next
41230         tb.MoveNext
41240     Loop
41250     For n = 1 To 4
41260         FixComboWidth cmbIdentification(n)
41270     Next n

41280     Exit Sub

LoadListOrganism_Error:

          Dim strES As String
          Dim intEL As Integer

41290     intEL = Erl
41300     strES = Err.Description
41310     LogError "frmEditMicrobiology", "LoadListOrganism", intEL, strES, sql

End Sub

Private Sub LoadListWCC()

          Dim sql As String
          Dim tb As Recordset

41320     On Error GoTo LoadListWCC_Error

41330     ReDim ListWCC(0 To 0) As String
41340     ListWCC(0) = ""

41350     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'WW' and InUse = 1 " & _
              "ORDER BY ListOrder"
41360     Set tb = New Recordset
41370     RecOpenServer 0, tb, sql
41380     Do While Not tb.EOF
41390         ReDim Preserve ListWCC(0 To UBound(ListWCC) + 1)
41400         ListWCC(UBound(ListWCC)) = tb!Text & ""
41410         tb.MoveNext
41420     Loop

41430     Exit Sub

LoadListWCC_Error:

          Dim strES As String
          Dim intEL As Integer

41440     intEL = Erl
41450     strES = Err.Description
41460     LogError "frmEditMicrobiology", "LoadListWCC", intEL, strES, sql

End Sub
Private Function LoadLockStatus(ByVal Index As Integer) As Boolean

          'Returns True if locked

          Dim tb As Recordset
          Dim sql As String
          Dim ColName As String

41470     On Error GoTo LoadLockStatus_Error

41480     Select Case Index
              Case 0: ColName = "Demo"
41490         Case 1: ColName = "Microscopy"
41500         Case 2: ColName = "Ident"
41510         Case 3: ColName = "Faeces"
41520         Case 4: ColName = "CandS"
41530         Case 5: ColName = "FOB"
41540         Case 6: ColName = "RotaAdeno"
41550         Case 7: ColName = "RedSub"
41560         Case 8: ColName = "RSV"
41570         Case 9: ColName = "CSF"
41580         Case 10: ColName = "CDiff"
41590         Case 11: ColName = "OP"
41600         Case 12: ColName = "Identification"
41610     End Select

41620     LoadLockStatus = False
41630     cmdLock(Index).Caption = "&Lock Result"
41640     cmdLock(Index).Picture = frmMain.ImageList1.ListImages("Key").Picture

41650     sql = "SELECT COALESCE(" & ColName & ", 0) AS Status FROM LockStatus WHERE " & _
              "SampleID = '" & Val(txtSampleID) & "'" ' + sysOptMicroOffset(0) & "'"
41660     Set tb = New Recordset
41670     RecOpenServer 0, tb, sql
41680     If Not tb.EOF Then
41690         If tb!Status > 0 Then
41700             cmdLock(Index).Caption = "Un&Lock Result"
41710             cmdLock(Index).Picture = frmMain.ImageList1.ListImages("Locked").Picture
41720             LoadLockStatus = True
41730         End If
41740     End If

41750     Exit Function

LoadLockStatus_Error:

          Dim strES As String
          Dim intEL As Integer

41760     intEL = Erl
41770     strES = Err.Description
41780     LogError "frmEditMicrobiology", "LoadLockStatus", intEL, strES, sql

End Function

Private Sub LockFraCS(ByVal Lockit As Boolean)

41790     cmdLock(1).Visible = Not Lockit
41800     cmdLock(4).Visible = Not Lockit
41810     cmdLock(5).Visible = Not Lockit
41820     cmdLock(6).Visible = Not Lockit
41830     cmdLock(7).Visible = Not Lockit
41840     cmdLock(8).Visible = Not Lockit
41850     cmdLock(10).Visible = Not Lockit
41860     cmdLock(11).Visible = Not Lockit

41870     fraMicroResult(4).Enabled = Not Lockit

End Sub

Private Sub MoveCursorToSaveButton()

          Dim t As Single

41880     t = Timer

41890     SetCursorPos (cmdSaveMicro.Left + (cmdSaveMicro.width / 2)) / Screen.TwipsPerPixelX, _
              (cmdSaveMicro.Top + cmdSaveMicro.height) / Screen.TwipsPerPixelY

41900     cmdSaveMicro.BackColor = vbYellow

41910     Do While Timer - t < 0.5
41920         DoEvents
41930     Loop

41940     cmdSaveMicro.BackColor = vbButtonFace

End Sub

Private Function QueryCEF() As Boolean

          Dim grd As Integer
          Dim y As Integer
          Dim S As String
          Dim FoundSens As Boolean
          Dim FoundCEF As Boolean
          Dim FoundResults As Boolean

41950     QueryCEF = False

41960     If UCase(cmbSite) = "URINE" Then
41970         FoundSens = False
41980         FoundCEF = False
41990         FoundResults = False

42000         For grd = 1 To 4
42010             If grdAB(grd).TextMatrix(1, 0) <> "" Then
42020                 FoundResults = True
42030                 For y = 1 To grdAB(grd).Rows - 1
42040                     grdAB(grd).Col = 0
42050                     grdAB(grd).row = y
42060                     If grdAB(grd).Font.Bold = False Then
42070                         If grdAB(grd).CellBackColor = 0 Then
42080                             If grdAB(grd).TextMatrix(y, 1) = "S" Then
42090                                 FoundSens = True
42100                                 Exit For
42110                             End If
42120                         End If
42130                     End If
42140                 Next
42150                 If Not FoundSens Then
42160                     For y = 1 To grdAB(grd).Rows - 1
42170                         If grdAB(grd).TextMatrix(y, 0) = "Cefuroxime" Then
42180                             grdAB(grd).Col = 2
42190                             grdAB(grd).row = y
42200                             If grdAB(grd).CellPicture = imgSquareTick.Picture Then
42210                                 FoundCEF = True
42220                                 Exit For
42230                             End If
42240                         End If
42250                     Next
42260                     If FoundCEF Then
42270                         Exit For
42280                     End If
42290                 End If
42300             End If
42310         Next
42320         If FoundResults And (Not FoundSens) And (Not FoundCEF) Then

42330             S = "No First line Antibiotics are Sensitive!" & vbCrLf & _
                      "Do you wish to report Cefuroxime?"
42340             If iMsg(S, vbQuestion + vbYesNo) = vbYes Then
42350                 QueryCEF = True
42360             End If
42370         End If

42380     End If

End Function

Private Function QueryGent() As Integer
          'returns 0 if not paediatrics and not scbu
          '        1 if dont report
          '        2 if force report

          Dim grd As Integer
          Dim y As Integer
          Dim S As String
          Dim Reported As Boolean
          Dim FoundResults As Boolean

42390     On Error GoTo QueryGent_Error

42400     QueryGent = 0
42410     FoundResults = False
42420     If UCase(cmbWard) = "PAEDIATRICS" Or UCase$(cmbWard) = "SPECIAL CARE BABY UNIT" Then
42430         Reported = False
42440         For grd = 1 To 4
42450             If grdAB(grd).TextMatrix(1, 0) <> "" Then
42460                 FoundResults = True
42470                 For y = 1 To grdAB(grd).Rows - 1
42480                     If grdAB(grd).TextMatrix(y, 0) = "Gentamicin" Then
42490                         grdAB(grd).row = y
42500                         grdAB(grd).Col = 2
42510                         If grdAB(grd).CellPicture = imgSquareTick.Picture Then
42520                             Reported = True
42530                             Exit For
42540                         End If
42550                     End If
42560                 Next
42570                 If Reported Then
42580                     Exit For
42590                 End If
42600             End If
42610         Next
42620         If FoundResults And Reported Then
42630             S = "This Isolate is from a Patient" & vbCrLf & _
                      "in " & cmbWard & "." & vbCrLf & _
                      "Do you wish to report Gentamicin?"
42640             If iMsg(S, vbQuestion + vbYesNo) = vbYes Then
42650                 QueryGent = 2
42660             Else
42670                 QueryGent = 1
42680             End If
42690         End If
42700     End If

42710     Exit Function

QueryGent_Error:

          Dim strES As String
          Dim intEL As Integer

42720     intEL = Erl
42730     strES = Err.Description
42740     LogError "frmEditMicrobiology", "QueryGent", intEL, strES

End Function

Private Sub SaveCurrentAntibiotics()

          Dim CURS As New CurrentAntibiotics
          Dim Cur As CurrentAntibiotic
          Dim n As Integer

42750     On Error GoTo SaveCurrentAntibiotics_Error

42760     For n = 0 To lstABsInUse.ListCount - 1
42770         If Trim$(lstABsInUse.List(n)) <> "" Then
42780             Set Cur = New CurrentAntibiotic
                  '+++ Junaid 20-05-2024
                  '50            Cur.SampleID = SampleIDWithOffset
42790             Cur.SampleID = Trim(txtSampleID.Text)
                  '--- Junaid
42800             Cur.Entry = n
42810             Cur.Antibiotic = lstABsInUse.List(n)
42820             Cur.UserName = UserName
42830             CURS.Add Cur
42840         End If
42850     Next
42860     CURS.Save

42870     Set CurrentSensitivities = New Sensitivities
          '+++ Junaid 20-05-2024
          '140   CurrentSensitivities.Load SampleIDWithOffset
42880     CurrentSensitivities.Load Trim(txtSampleID.Text)
          '--- JUnaid
42890     LoadSensitivities

42900     Exit Sub

SaveCurrentAntibiotics_Error:

          Dim strES As String
          Dim intEL As Integer

42910     intEL = Erl
42920     strES = Err.Description
42930     LogError "frmEditMicrobiology", "SaveCurrentAntibiotics", intEL, strES

End Sub

Private Sub SaveMicro(ByVal IncrementSID As Boolean, ByVal Validate As Integer)

42940     pBar = 0

          '20    If Not GetSampleIDWithOffset Then Exit Sub
42950     m_Flag = True
42960     SaveComments
42970     DoEvents
          '40    SaveIdentification
42980     If SSTab1.Tab = Mic.Faeces Then
42990         SaveFaeces
43000         FillHistoricalFaeces
43010     End If

          '60    If Not CheckIfValid() Then

43020     Select Case SSTab1.Tab
              Case Mic.Urine: SaveUrine Validate
43030         Case Mic.Identification: 'SaveIdentification
43040         Case Mic.CandS: SaveIsolates
43050             SaveSensitivities Validate
43060         Case Mic.FOB: SaveFOB Validate
43070         Case Mic.RotaAdeno: SaveRotaAdeno Validate
43080         Case Mic.CDiff: SaveCdiff Validate
43090         Case Mic.OP: SaveOP Validate
43100         Case Mic.IDENTCAVAN: SaveIdentification
43110         Case Mic.RSV: SaveRSV Validate
43120         Case Mic.CSF: SaveCSF
43130         Case Mic.RedSub: SaveRedSub Validate
43140     End Select
43150     UpdateMRU Me

          'Call LabNoUpdatePrvData(txtChart, Trim$(UCase$(AddTicks(txtSurName & " " & txtForeName))), txtDoB, left$(txtSex, 1), txtLabNo)
          
          'Abubaker +++ 13|09|2023 (I commented down this piece of code responsible for increamenting Sample ID)
          'Abubaker +++ 05|10|2023 (Uncommented)
          'Abubaker +++ 09|10|2023 (re-commented)

          '230   If IncrementSID Then
          '240       txtSampleID = Format$(Val(txtSampleID) + 1)
          '250       GetSampleIDWithOffset
          '260   End If


          '250   End If
          '      DoEvents
43160     cmdSaveMicro.Enabled = False
43170     cmdSaveHold.Enabled = False

End Sub

Private Sub SelectBloodCulture()

          Dim sql As String
          Dim tb As Recordset
          Dim SuggestedAssID As String
          Dim ConfirmedAssID As Long
          Dim RecordsEffected As Long

43180     On Error GoTo SelectBloodCulture_Error

43190     SaveDemographics

          '30    sql = "SELECT AssID FROM Demographics WHERE " & _
          '            "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
          '            "AND AssID IS NOT NULL"
43200     sql = "SELECT AssID FROM Demographics WHERE " & _
              "SampleID = '" & Val(txtSampleID) & "' " & _
              "AND AssID IS NOT NULL"
43210     Set tb = New Recordset
43220     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          ''
          '--- Junaid
43230     If tb.EOF Then    'AssID unknown - look in prev and next for possibilities
43240         sql = "SELECT SampleID FROM Demographics WHERE " & _
                  "( SampleID = '" & Val(txtSampleID) + 1 & "' " & _
                  "  OR SampleID = '" & Val(txtSampleID) - 1 & "') " & _
                  "AND Chart = '" & AddTicks(txtChart) & "' " & _
                  "AND PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                  "AND AssID IS NULL"
43250         Set tb = New Recordset
43260         RecOpenServer 0, tb, sql
              '+++ Junaid 12-12-2023
              ''
              '--- Junaid
43270         If tb.EOF Then    'not in previous or next
43280             SuggestedAssID = ""
43290         Else
                  '130           SuggestedAssID = CStr(Val(tb!SampleID) - sysOptMicroOffset(0))
43300             SuggestedAssID = CStr(Val(tb!SampleID))
43310         End If
43320     Else    'AssID already known
43330         SuggestedAssID = CStr(Val(tb!AssID))
              '160       SuggestedAssID = CStr(Val(tb!AssID) - sysOptMicroOffset(0))
43340         If Val(SuggestedAssID) < 1 Then
43350             SuggestedAssID = ""
43360         End If
43370     End If
43380     ConfirmedAssID = Val(iBOX("Confirm associated Sample ID", "Blood Culture", SuggestedAssID))
43390     If ConfirmedAssID <> 0 Then
              '230       sql = "Update Demographics " & _
              '                "Set AssID = '" & ConfirmedAssID + sysOptMicroOffset(0) & "' " & _
              '                "WHERE SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
43400         sql = "Update Demographics " & _
                  "Set AssID = '" & ConfirmedAssID + sysOptMicroOffset(0) & "' " & _
                  "WHERE SampleID = '" & Val(txtSampleID) & "'"
43410         Set tb = Cnxn(0).Execute(sql, RecordsEffected)
              '+++ Junaid 12-12-2023
              '
              '--- Junaid
43420         If RecordsEffected > 0 Then
                  '260           sql = "Update Demographics " & _
                  '                    "Set AssID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
                  '                    "WHERE SampleID = '" & ConfirmedAssID + sysOptMicroOffset(0) & "'"
43430             sql = "Update Demographics " & _
                      "Set AssID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
                      "WHERE SampleID = '" & ConfirmedAssID & "'"
43440             Set tb = Cnxn(0).Execute(sql, RecordsEffected)
                  '+++ Junaid 12-12-2023
                  '
                  '--- Junaid
43450             If RecordsEffected = 0 Then
43460                 iMsg "Record " & ConfirmedAssID & " does not exist!" & vbCrLf & "Cannot update record", vbExclamation
43470             End If
43480         Else
43490             iMsg "Record " & txtSampleID & " does not exist!" & vbCrLf & "Cannot update record", vbExclamation
43500         End If
43510     End If

43520     Exit Sub

SelectBloodCulture_Error:

          Dim strES As String
          Dim intEL As Integer

43530     intEL = Erl
43540     strES = Err.Description
43550     LogError "frmEditMicrobiology", "SelectBloodCulture", intEL, strES, sql

End Sub

Private Sub ShowUnlock(ByVal Index As Integer)

43560     cmdSaveMicro.Enabled = True
43570     cmdSaveHold.Enabled = True
43580     cmdLock(Index).Visible = True
43590     cmdLock(Index).Caption = "&Lock Result"
43600     cmdLock(Index).Picture = frmMain.ImageList1.ListImages("Key").Picture

End Sub

Private Sub UpdateLockStatus(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim ColName As String
          Dim SetIt As Integer

43610     On Error GoTo UpdateLockStatus_Error

43620     SetIt = IIf(cmdLock(Index).Caption = "&Lock Result", 1, 0)

43630     Select Case Index
              Case 0: ColName = "Demo"
43640         Case 1: ColName = "Microscopy"
43650         Case 2: ColName = "Ident"
43660         Case 3: ColName = "Faeces"
43670         Case 4: ColName = "CandS"
43680         Case 5: ColName = "FOB"
43690         Case 6: ColName = "RotaAdeno"
43700         Case 7: ColName = "RedSub"
43710         Case 8: ColName = "RSV"
43720         Case 9: ColName = "CSF"
43730         Case 10: ColName = "CDiff"
43740         Case 11: ColName = "OP"
43750         Case 12: ColName = "Identification"
43760     End Select

43770     sql = "IF EXISTS (SELECT * FROM LockStatus WHERE " & _
              "           SampleID = '" & Val(txtSampleID) & "') " & _
              "  UPDATE LockStatus " & _
              "  SET " & ColName & " = " & SetIt & " " & _
              "  WHERE SampleID = '" & Val(txtSampleID) & "' " & _
              "ELSE " & _
              "  INSERT INTO LockStatus " & _
              " (SampleID, " & ColName & ") " & _
              "  VALUES " & _
              " ('" & Val(txtSampleID) & "', " & SetIt & ")"
43780     Cnxn(0).Execute sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid

          '170   Set tb = New Recordset
          '180   RecOpenServer 0, tb, sql
          '190   If tb.EOF Then
          '200     tb.AddNew
          '210     tb!SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
          '220   End If

43790     If SetIt = 1 Then
              '240     tb(ColName) = 1
43800         cmdLock(Index).Caption = "Un&Lock Result"
43810         cmdLock(Index).Picture = frmMain.ImageList1.ListImages("Locked").Picture
43820     Else
              '280     tb(ColName) = 0
43830         cmdLock(Index).Caption = "&Lock Result"
43840         cmdLock(Index).Picture = frmMain.ImageList1.ListImages("Key").Picture
43850     End If

          '320   tb.Update

43860     fraMicroResult(Index).Enabled = cmdLock(Index).Caption = "&Lock Result"

43870     Exit Sub

UpdateLockStatus_Error:

          Dim strES As String
          Dim intEL As Integer

43880     intEL = Erl
43890     strES = Err.Description
43900     LogError "frmEditMicrobiology", "UpdateLockStatus", intEL, strES, sql

End Sub
Private Function LoadOP(ByVal Fxs As FaecesResults) As Boolean
          'Returns true if OP results present

          Dim n As Integer
          Dim Found As Boolean
          Dim Fx As FaecesResult

43910     On Error GoTo LoadOP_Error

43920     cmdLock(Mic.OP).Visible = False
43930     fraMicroResult(Mic.OP).Enabled = True
43940     lblCrypto = ""
43950     lblCrypto.BackColor = &H8000000F
43960     For n = 0 To 2
43970         cmbOva(n) = ""
43980     Next

43990     Found = False

44000     For Each Fx In Fxs
44010         Select Case UCase$(Fx.TestName)
                  Case "AUS"
44020                 Found = True
44030                 If Fx.Result = "N" Then
44040                     lblCrypto = "Negative"
44050                     lblCrypto.BackColor = vbGreen
44060                 ElseIf Fx.Result = "P" Then
44070                     lblCrypto = "Positive"
44080                     lblCrypto.BackColor = vbRed
44090                 End If
44100             Case "OP0": Found = True: cmbOva(0) = Fx.Result
44110             Case "OP1": Found = True: cmbOva(1) = Fx.Result
44120             Case "OP2": Found = True: cmbOva(2) = Fx.Result
44130         End Select
44140     Next

44150     If Found Then
44160         cmdLock(Mic.OP).Visible = True
44170         If LoadLockStatus(Mic.OP) Then
44180             fraMicroResult(Mic.OP).Enabled = False
44190         End If
44200         LoadOP = True
44210     Else
44220         LoadOP = False
44230     End If

44240     Exit Function

LoadOP_Error:

          Dim strES As String
          Dim intEL As Integer

44250     intEL = Erl
44260     strES = Err.Description
44270     LogError "frmEditMicrobiology", "LoadOP", intEL, strES

End Function


Private Function LoadCDiff(ByVal GenResults As GenericResults, ByVal Fxs As FaecesResults) As Boolean
          'Returns true if Cdiff results present

          Dim Found As Boolean
          Dim GenResult As GenericResult
          Dim Fx As FaecesResult

44280     On Error GoTo LoadCDiff_Error

44290     Found = False

44300     cmdLock(Mic.CDiff).Visible = False
44310     fraMicroResult(Mic.CDiff).Enabled = True
44320     lblToxinA = ""
44330     lblToxinA.BackColor = &H8000000F
44340     lblToxinB = ""
44350     lblToxinB.BackColor = &H8000000F

44360     LoadCDiff = False

44370     For Each Fx In Fxs

44380         If UCase$(Fx.TestName) = "TOXINAL" Then
44390             Found = True
44400             If Fx.Result = "N" Then
44410                 lblToxinA = "Not Detected"
44420                 lblToxinA.BackColor = vbGreen
44430             ElseIf Fx.Result = "P" Then
44440                 lblToxinA = "Positive"
44450                 lblToxinA.BackColor = vbRed
44460             ElseIf Fx.Result = "R" Then
44470                 lblToxinA = "Rejected"
44480             ElseIf Fx.Result = "I" Then
44490                 lblToxinA = "Inconclusive"
44500                 lblToxinA.BackColor = vbYellow
44510             End If

44520         ElseIf UCase$(Fx.TestName) = "TOXINATA" Then
44530             Found = True
44540             If Fx.Result = "N" Then
44550                 lblToxinB = "Not Detected"
44560                 lblToxinB.BackColor = vbGreen
44570             ElseIf Fx.Result = "P" Then
44580                 lblToxinB = "Positive"
44590                 lblToxinB.BackColor = vbRed
44600             ElseIf Fx.Result = "R" Then
44610                 lblToxinB = "Rejected"
44620             ElseIf Fx.Result = "I" Then
44630                 lblToxinB = "Inconclusive"
44640                 lblToxinB.BackColor = vbYellow
44650             End If
44660         End If
44670     Next

44680     lblcDiffPCR.Caption = ""
44690     lblcDiffPCR.BackColor = &H8000000F

44700     For Each GenResult In GenResults
44710         If UCase(GenResult.TestName) = "CDIFFPCR" Then

44720             Found = True
44730             lblcDiffPCR.Caption = GenResult.Result
44740             If lblcDiffPCR.Caption = "Toxigenic C.diff: NEGATIVE" Then
44750                 lblcDiffPCR.BackColor = vbGreen
44760             ElseIf lblcDiffPCR.Caption = "Toxigenic C.diff: POSITIVE" Then
44770                 lblcDiffPCR.BackColor = vbRed
44780             Else
44790                 lblcDiffPCR.BackColor = &H8000000F
44800             End If
44810         End If
44820     Next

44830     If Found Then
44840         cmdLock(Mic.CDiff).Visible = True
44850         If LoadLockStatus(Mic.CDiff) Then
44860             fraMicroResult(Mic.CDiff).Enabled = False
44870         End If
44880         LoadCDiff = True
44890     Else
44900         LoadCDiff = False
44910     End If

44920     Exit Function

LoadCDiff_Error:

          Dim strES As String
          Dim intEL As Integer

44930     intEL = Erl
44940     strES = Err.Description
44950     LogError "frmEditMicrobiology", "LoadCDiff", intEL, strES

End Function

Private Function LoadCSF() As Boolean
          'Returns true if CSF results present

          Dim tb As Recordset
          Dim sql As String

44960     On Error GoTo LoadCSF_Error

44970     ClearCSF

44980     cmdLock(Mic.CSF).Visible = False
44990     fraMicroResult(Mic.CSF).Enabled = True

45000     LoadCSF = False
          '+++ Junaid 08-05-2024
          '60    sql = "SELECT * FROM CSFResults WHERE " & _
           "SampleID = '" & SampleIDWithOffset & "'"
45010     sql = "SELECT * FROM CSFResults WHERE " & _
              "SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
45020     Set tb = New Recordset
45030     RecOpenServer 0, tb, sql

45040     If Not tb.EOF Then
45050         cmbCSFGram = tb!Gram & ""
45060         txtCSFWCCDiff(0) = tb!WCCDiff0 & ""
45070         txtCSFWCCDiff(1) = tb!WCCDiff1 & ""
45080         cmbCSFAppearance(0) = tb!Appearance0 & ""
45090         cmbCSFAppearance(1) = tb!Appearance1 & ""
45100         cmbCSFAppearance(2) = tb!Appearance2 & ""
45110         txtCSFWCC(0) = tb!WCC0 & ""
45120         txtCSFWCC(1) = tb!WCC1 & ""
45130         txtCSFWCC(2) = tb!WCC2 & ""
45140         txtCSFRCC(0) = tb!RCC0 & ""
45150         txtCSFRCC(1) = tb!RCC1 & ""
45160         txtCSFRCC(2) = tb!RCC2 & ""

45170         cmdLock(Mic.CSF).Visible = True
45180         If LoadLockStatus(Mic.CSF) Then
45190             fraMicroResult(Mic.CSF).Enabled = False
45200         End If
45210         LoadCSF = True
45220     Else
45230         LoadCSF = False
45240     End If

45250     Exit Function

LoadCSF_Error:

          Dim strES As String
          Dim intEL As Integer

45260     intEL = Erl
45270     strES = Err.Description
45280     LogError "frmEditMicrobiology", "LoadCSF", intEL, strES, sql

End Function

Private Sub LoadPrintStatus(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim ColName As String

45290     On Error GoTo LoadPrintStatus_Error

45300     Select Case Index
              Case 0: ColName = "Demo"
45310         Case 1: ColName = "Microscopy"
45320         Case 2: ColName = "Ident"
45330         Case 3: ColName = "Faeces"
45340         Case 4: ColName = "CandS"
45350         Case 5: ColName = "FOB"
45360         Case 6: ColName = "RotaAdeno"
45370         Case 7: ColName = "RedSub"
45380         Case 8: ColName = "RSV"
45390         Case 9: ColName = "CSF"
45400         Case 10: ColName = "CDiff"
45410         Case 11: ColName = "OP"
45420         Case 12: ColName = "Identification"
45430     End Select

45440     lblPrinted(Index).Visible = False

          '170   sql = "SELECT COALESCE(" & ColName & ", 0) AS Status FROM PrintedStatus WHERE " & _
          '            "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
45450     sql = "SELECT COALESCE(" & ColName & ", 0) AS Status FROM PrintedStatus WHERE " & _
              "SampleID = '" & Val(txtSampleID) & "'"
45460     Set tb = New Recordset
45470     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
45480     If Not tb.EOF Then
45490         If tb!Status > 0 Then
45500             lblPrinted(Index).Visible = True
45510         End If
45520     End If

45530     Exit Sub

LoadPrintStatus_Error:

          Dim strES As String
          Dim intEL As Integer

45540     intEL = Erl
45550     strES = Err.Description
45560     LogError "frmEditMicrobiology", "LoadPrintStatus", intEL, strES, sql

End Sub

Private Function LoadRotaAdeno(ByVal Fxs As FaecesResults) As Boolean
          'Returns true if Rota/Adeno results present

          Dim Found As Boolean
          Dim Fx As FaecesResult

45570     On Error GoTo LoadRotaAdeno_Error

45580     Found = False

45590     cmdLock(Mic.RotaAdeno).Visible = False
45600     fraMicroResult(Mic.RotaAdeno).Enabled = True
45610     txtRota = ""
45620     txtRota.BackColor = &H8000000F
45630     txtAdeno = ""
45640     txtAdeno.BackColor = &H8000000F

45650     LoadRotaAdeno = False

45660     For Each Fx In Fxs
45670         If UCase$(Fx.TestName) = "ROTA" Then
45680             Found = True
45690             If Fx.Result = "N" Then
45700                 txtRota = "Negative"
45710                 txtRota.BackColor = vbGreen
45720             ElseIf Fx.Result = "P" Then
45730                 txtRota = "Positive"
45740                 txtRota.BackColor = vbRed
45750             End If
45760         ElseIf UCase$(Fx.TestName) = "ADENO" Then
45770             Found = True
45780             If Fx.Result = "N" Then
45790                 txtAdeno = "Negative"
45800                 txtAdeno.BackColor = vbGreen
45810             ElseIf Fx.Result = "P" Then
45820                 txtAdeno = "Positive"
45830                 txtAdeno.BackColor = vbRed
45840             End If
45850         End If
45860     Next

45870     If Found Then
45880         cmdLock(Mic.RotaAdeno).Visible = True
45890         If LoadLockStatus(Mic.RotaAdeno) Then
45900             fraMicroResult(Mic.RotaAdeno).Enabled = False
45910         End If
45920         LoadRotaAdeno = True
45930     Else
45940         LoadRotaAdeno = False
45950     End If

45960     Exit Function

LoadRotaAdeno_Error:

          Dim strES As String
          Dim intEL As Integer

45970     intEL = Erl
45980     strES = Err.Description
45990     LogError "frmEditMicrobiology", "LoadRotaAdeno", intEL, strES

End Function

Private Function LoadRedSub(ByVal GenResults As GenericResults) As Boolean
          'Returns true if Reducing Substances results present

          Dim n As Integer
          Dim GenResult As GenericResult

46000     On Error GoTo LoadRedSub_Error

46010     cmdLock(Mic.RedSub).Visible = False
46020     fraMicroResult(Mic.RedSub).Enabled = True
46030     For n = 0 To 5
46040         chkRS(n).Value = 0
46050     Next

46060     LoadRedSub = False
46070     For Each GenResult In GenResults
46080         If UCase$(GenResult.TestName) = "REDSUB" Then
46090             For n = 0 To 5
46100                 If chkRS(n).Caption = GenResult.Result Then
46110                     chkRS(n).Value = 1
46120                     Exit For
46130                 End If
46140             Next

46150             cmdLock(Mic.RedSub).Visible = True
46160             If LoadLockStatus(Mic.RedSub) Then
46170                 fraMicroResult(Mic.RedSub).Enabled = False
46180             End If
46190             SSTab1.TabVisible(Mic.RedSub) = True
46200             LoadRedSub = True

46210         End If
46220     Next

46230     Exit Function

LoadRedSub_Error:

          Dim strES As String
          Dim intEL As Integer

46240     intEL = Erl
46250     strES = Err.Description
46260     LogError "frmEditMicrobiology", "LoadRedSub", intEL, strES

End Function

Private Function LoadRSV(ByVal GenResults As GenericResults) As Boolean
          'Returns true if RSV results present

          Dim GenResult As GenericResult

46270     On Error GoTo LoadRSV_Error

46280     cmdLock(Mic.RSV).Visible = False
46290     fraMicroResult(Mic.RSV).Enabled = True
46300     lblRSV.Caption = ""
46310     lblRSV.BackColor = &H8000000F

46320     LoadRSV = False
46330     For Each GenResult In GenResults
46340         If GenResult.TestName = "RSV" Then

46350             Select Case UCase$(GenResult.Result)
                      Case "NEGATIVE"
46360                     lblRSV = "Negative"
46370                     lblRSV.BackColor = vbGreen
46380                 Case "POSITIVE"
46390                     lblRSV = "Positive"
46400                     lblRSV.BackColor = vbRed
46410                 Case "INCONCLUSIVE"
46420                     lblRSV = "Inconclusive"
46430                     lblRSV.BackColor = vbYellow
46440             End Select

46450             cmdLock(Mic.RSV).Visible = True
46460             If LoadLockStatus(Mic.RSV) Then
46470                 fraMicroResult(Mic.RSV).Enabled = False
46480             End If

46490             LoadRSV = True

46500         End If
46510     Next

46520     Exit Function

LoadRSV_Error:

          Dim strES As String
          Dim intEL As Integer

46530     intEL = Erl
46540     strES = Err.Description
46550     LogError "frmEditMicrobiology", "LoadRSV", intEL, strES

End Function

Private Function LoadFOB(ByVal Fxs As FaecesResults) As Boolean
          'Returns true if FOB results present

          Dim n As Integer
          Dim Found As Boolean
          Dim Fx As FaecesResult

46560     On Error GoTo LoadFOB_Error

46570     Found = False
46580     cmdLock(Mic.FOB).Visible = False
46590     fraMicroResult(Mic.FOB).Enabled = True
46600     For n = 0 To 2
46610         chkFOB(n) = 0
46620         lblFOB(n) = ""
46630         lblFOB(n).BackColor = &H8000000F
46640     Next

46650     LoadFOB = False

46660     For Each Fx In Fxs
46670         Select Case UCase$(Fx.TestName)
                  Case "OB0"
46680                 Found = True
46690                 chkFOB(0) = 1
46700                 If Fx.Result = "N" Then
46710                     lblFOB(0) = "Negative": lblFOB(0).BackColor = vbGreen
46720                 ElseIf Fx.Result = "P" Then
46730                     lblFOB(0) = "Positive": lblFOB(0).BackColor = vbRed
46740                 End If
46750             Case "OB1"
46760                 Found = True
46770                 chkFOB(1) = 1
46780                 If Fx.Result = "N" Then
46790                     lblFOB(1) = "Negative": lblFOB(1).BackColor = vbGreen
46800                 ElseIf Fx.Result = "P" Then
46810                     lblFOB(1) = "Positive": lblFOB(1).BackColor = vbRed
46820                 End If
46830             Case "OB2"
46840                 Found = True
46850                 chkFOB(2) = 1
46860                 If Fx.Result = "N" Then
46870                     lblFOB(2) = "Negative": lblFOB(2).BackColor = vbGreen
46880                 ElseIf Fx.Result = "P" Then
46890                     lblFOB(2) = "Positive": lblFOB(2).BackColor = vbRed
46900                 End If
46910         End Select
46920     Next

46930     If Found Then
46940         cmdLock(Mic.FOB).Visible = True
46950         If LoadLockStatus(Mic.FOB) Then
46960             fraMicroResult(Mic.FOB).Enabled = False
46970         End If
46980         LoadFOB = True
46990     Else
47000         LoadFOB = False
47010     End If

47020     Exit Function

LoadFOB_Error:

          Dim strES As String
          Dim intEL As Integer

47030     intEL = Erl
47040     strES = Err.Description
47050     LogError "frmEditMicrobiology", "LoadFOB", intEL, strES

End Function

Private Function LoadIdentification() As Integer
          'Returns number of Isolates Loaded

          '+++ Junaid 22-12-2023
          Dim sql As String
          Dim tb As Recordset
          '--- Junaid
          Dim n As Integer
          Dim intMax As Integer
          Dim Ix As UrineIdent
          Dim Ixs As New UrineIdents

47060     On Error GoTo LoadIdentification_Error

47070     intMax = 0
47080     txtIdentification(1) = ""
47090     txtIdentification(2) = ""
47100     txtIdentification(3) = ""
47110     txtIdentification(4) = ""
47120     For n = 1 To 4
47130         cmbIdentification(n) = ""
47140         txtIdentification(n) = ""
47150     Next
          '      DoEvents
          '      DoEvents
          '+++ Junaid 22-12-2023
47160     sql = "Select * from CavanLog Where SampleID = '" & txtSampleID.Text & "'"
47170     Set tb = New Recordset
47180     RecOpenClient 0, tb, sql
47190     If Not tb.EOF Then
47200         txtIdentification(1) = "(" & ConvertNull(tb!SampleID, "") & "): "
47210     End If
          '--- Junaid
          '+++ Junaid 26-12-2023
47220     Ixs.Clear
          '--- Junaid
          '+++ JUnaid 20-05-2024
          '70    Ixs.Load SampleIDWithOffset
47230     Ixs.Load Trim(txtSampleID.Text)
          '--- Junaid

47240     For n = 1 To 4
47250         Set Ix = Ixs.Item("Notes", Format(n))
47260         If Not Ix Is Nothing Then
47270             intMax = n
47280             txtIdentification(n) = Trim$(Ix.Result)
47290         End If
47300     Next

47310     LoadIdentification = intMax

47320     cmdSaveMicro.Enabled = False
47330     cmdSaveHold.Enabled = False

47340     Exit Function

LoadIdentification_Error:

          Dim strES As String
          Dim intEL As Integer

47350     intEL = Erl
47360     strES = Err.Description
47370     LogError "frmEditMicrobiology", "LoadIdentification", intEL, strES

End Function

Private Sub LoadIsolates()

          Dim Isos As New Isolates
          Dim Iso As Isolate
          Dim n As Integer

47380     On Error GoTo LoadIsolates_Error

47390     For n = 1 To 4
47400         cmbOrgGroup(n) = ""
47410         cmbOrgName(n) = ""
47420         cmbQualifier(n) = ""
47430     Next
          '+++Junaid 14-05-2024
          '70    Isos.Load SampleIDWithOffset
47440     Isos.Load Trim(txtSampleID.Text)
          '--- Junaid
47450     For Each Iso In Isos
47460         n = Iso.IsolateNumber
47470         cmbOrgGroup(n) = Iso.OrganismGroup
47480         cmbOrgName(n) = Iso.OrganismName
47490         cmbQualifier(n) = Iso.Qualifier
47500     Next

47510     Exit Sub

LoadIsolates_Error:

          Dim strES As String
          Dim intEL As Integer

47520     intEL = Erl
47530     strES = Err.Description
47540     LogError "frmEditMicrobiology", "LoadIsolates", intEL, strES

End Sub


Private Function LoadUrine() As Integer
          'Returns 0 if no Urine Results Present
          '        1 if valid
          '        2 if not valid
          Dim ForPreg As Boolean
          Dim UReqs As New UrineRequests
          Dim UReq As UrineRequest
          Dim UResults As New UrineResults
          Dim UResult As UrineResult
          Dim RetVal As Integer
          Dim Result As String

47550     On Error GoTo LoadUrine_Error

47560     RetVal = 0

47570     UrineAutoVal = ""

47580     ForPreg = False
          '+++ Junaid 08-05-2024
          '50    UReqs.Load SampleIDWithOffset
47590     UReqs.Load Val(txtSampleID.Text)
          '--- Junaid
47600     Set UReq = UReqs.Item("Pregnancy")
47610     If Not UReq Is Nothing Then
47620         ForPreg = True
47630     End If

47640     ClearUrine

47650     cmdLock(Mic.Urine).Visible = False
47660     fraMicroResult(Mic.Urine).Enabled = True
          '+++ Junaid 08-05-2024
          '130   UResults.Load SampleIDWithOffset
47670     UResults.Load Val(txtSampleID.Text)
          '--- Junaid
47680     cmdTagRepeat.Caption = "Tag as Repeat"

47690     For Each UResult In UResults

47700         If UResult.Valid Then
47710             RetVal = 1
47720         Else
47730             RetVal = 2
47740         End If

47750         Select Case UCase$(UResult.TestName)
                  Case "PREGNANCY": txtPregnancy = UResult.Result
47760             Case "BACTERIA": txtBacteria = UResult.Result
47770             Case "WCC": txtWCC = UResult.Result
47780             Case "RCC": txtRCC = UResult.Result
47790             Case "HCGLEVEL": txtHCGLevel = UResult.Result
47800             Case "BENCEJONES": txtBenceJones = UResult.Result
47810             Case "SG": txtSG = UResult.Result
47820             Case "FATGLOBULES": txtFatGlobules = UResult.Result
47830             Case "PH": txtpH = UResult.Result
47840             Case "PROTEIN": txtProtein = UResult.Result
47850             Case "GLUCOSE": txtGlucose = UResult.Result
47860             Case "KETONES": txtKetones = UResult.Result
47870             Case "UROBILINOGEN": txtUrobilinogen = UResult.Result
47880             Case "BILIRUBIN": txtBilirubin = UResult.Result
47890             Case "BLOODHB": txtBloodHb = UResult.Result
47900             Case "CRYSTALS": cmbCrystals = UResult.Result
47910             Case "CASTS": cmbCasts = UResult.Result
47920             Case "MISC0": cmbMisc(0) = UResult.Result
47930             Case "MISC1": cmbMisc(1) = UResult.Result
47940             Case "MISC2": cmbMisc(2) = UResult.Result
47950             Case "TAGREPEAT": cmdTagRepeat.Caption = UResult.Result
47960         End Select
47970     Next

47980     Set UResults = New UrineResults
47990     UResults.LoadSedimax txtSampleID
48000     UResults.CheckForAllResults
48010     If UResults.Count > 0 Then
48020         ShowUnlock Mic.Urine
48030         For Each UResult In UResults

48040             UResults.Save UResult

48050             Select Case UCase$(UResult.TestName)

                      Case "RBC"
48060                     If Trim$(txtRCC) = "" Then
48070                         txtRCC = GetPlussesOrNil(UResult.Result)
48080                         RetVal = 1
48090                     End If

48100                 Case "WBC"
48110                     If Trim$(txtWCC) = "" Then
48120                         txtWCC = GetWBCValue(UResult.Result)
48130                         RetVal = 1
48140                     End If

48150                 Case "BAC"
48160                     If Trim$(txtBacteria) = "" Then
48170                         txtBacteria = GetPlussesOrNil(UResult.Result)
48180                         RetVal = 1
48190                     End If

48200                 Case "CRY"
48210                     If cmbCrystals = "" Then
48220                         Result = GetPlusses(UResult.Result)
48230                         If Result <> "" Then
48240                             cmbCrystals = Expand(UResult.Result) & " " & Result
48250                             RetVal = 1
48260                         End If
48270                     End If

48280                 Case "HYA":
48290                     If cmbCasts = "" Then
48300                         Result = GetPlusses(UResult.Result)
48310                         If Result <> "" Then
48320                             cmbCasts = "Casts - Hyalin " & Result
48330                             RetVal = 1
48340                         End If
48350                     End If

48360                 Case "PAT":
48370                     If cmbCasts = "" Then
48380                         Result = GetPlusses(UResult.Result)
48390                         If Result <> "" Then
48400                             cmbCasts = "Casts - Pathological " & Result
48410                             RetVal = 1
48420                         End If
48430                     End If

48440                 Case "EPI":
48450                     Result = GetPlusses(UResult.Result)
48460                     If Result <> "" Then
48470                         Result = "Epithelial Cells " & Result
48480                         fillFirstMisc Result
48490                         RetVal = 1
48500                     End If

48510                 Case "YEA":    'Yeasts
48520                     Result = GetPlusses(UResult.Result)
48530                     If Result <> "" Then
48540                         Result = "Yeasts " & Result
48550                         fillFirstMisc Result
48560                         RetVal = 1
48570                     End If

48580             End Select
48590         Next
48600     End If

48610     If RetVal = 1 Then    'valid
48620         fraMicroResult(Mic.Urine).Enabled = False
48630     Else
48640         cmdLock(Mic.Urine).Visible = True
48650         If LoadLockStatus(Mic.Urine) Then
48660             fraMicroResult(Mic.Urine).Enabled = False
48670         End If
48680     End If

          '1100  If Not ForPreg And Trim$(txtBacteria & txtWCC & txtRCC) = "" Then
          '1110    cmdNADMicro_Click
          '1120  End If
48690     If UReqs.Count > 0 Then
48700         RetVal = 1
48710     End If

48720     LoadUrine = RetVal

48730     Exit Function

LoadUrine_Error:

          Dim strES As String
          Dim intEL As Integer

48740     intEL = Erl
48750     strES = Err.Description
48760     LogError "frmEditMicrobiology", "LoadUrine", intEL, strES

End Function


Private Function UrineResultsPresent() As Boolean

          Dim UResults As New UrineResults
          Dim UResult As UrineResult
          Dim RetVal As Boolean

48770     On Error GoTo UrineResultsPresent_Error

48780     RetVal = False
          '+++ Junaid 08-05-2024
          '30    UResults.Load SampleIDWithOffset
48790     UResults.Load Val(txtSampleID.Text)
          '--- Junaid

48800     For Each UResult In UResults

48810         Select Case UCase$(UResult.TestName)
                  Case "PREGNANCY": RetVal = True
48820             Case "BACTERIA": RetVal = True
48830             Case "WCC": RetVal = True
48840             Case "RCC": RetVal = True
48850             Case "HCGLEVEL": RetVal = True
48860             Case "BENCEJONES": RetVal = True
48870             Case "SG": RetVal = True
48880             Case "FATGLOBULES": RetVal = True
48890             Case "PH": RetVal = True
48900             Case "PROTEIN": RetVal = True
48910             Case "GLUCOSE": RetVal = True
48920             Case "KETONES": RetVal = True
48930             Case "UROBILINOGEN": RetVal = True
48940             Case "BILIRUBIN": RetVal = True
48950             Case "BLOODHB": RetVal = True
48960             Case "CRYSTALS": RetVal = True
48970             Case "CASTS": RetVal = True
48980             Case "MISC0", "MISC1", "MISC2": RetVal = True
48990         End Select
49000     Next

49010     UrineResultsPresent = RetVal

49020     Exit Function

UrineResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

49030     intEL = Erl
49040     strES = Err.Description
49050     LogError "frmEditMicrobiology", "UrineResultsPresent", intEL, strES

End Function


Private Sub OrderFaeces()

          Dim f As Form
          Dim n As Integer
          Dim Fxs As FaecesRequests
          Dim Fx As FaecesRequest

49060     For n = 1 To 12
49070         SSTab1.TabVisible(n) = False
49080     Next
49090     SSTab1.TabVisible(Mic.CandS) = True
49100     SSTab1.TabVisible(Mic.IDENTCAVAN) = True

49110     Set f = New frmMicroOrderFaeces
49120     With f
49130         .txtSampleID = txtSampleID
49140         .Show 1
49150         Set Fxs = .FaecalOrders
49160     End With
49170     Unload f
49180     Set f = Nothing
49190     If Not Fxs Is Nothing Then
49200         For Each Fx In Fxs
49210             Select Case UCase$(Fx.Request)
                      Case "ROTAADENO": SSTab1.TabVisible(Mic.RotaAdeno) = True
49220                 Case "CDIFF": SSTab1.TabVisible(Mic.CDiff) = True
49230                 Case "OB0", "OB1", "OB2": SSTab1.TabVisible(Mic.FOB) = True
49240                 Case "REDSUB": SSTab1.TabVisible(Mic.RedSub) = True
49250             End Select
49260         Next
49270     End If
49280     SSTab1.TabVisible(Mic.OP) = True

49290     cmdOrderTests.Enabled = True

End Sub

Private Sub OrderUrine()

          Dim f As Form
          Dim n As Integer
          Dim Uxs As UrineRequests
          Dim Ux As UrineRequest

49300     On Error GoTo OrderUrine_Error

49310     For n = 1 To 12
49320         SSTab1.TabVisible(n) = False
49330     Next
49340     SSTab1.TabVisible(Mic.CandS) = True
49350     SSTab1.TabVisible(Mic.IDENTCAVAN) = True
49360     SSTab1.TabVisible(Mic.Urine) = True

49370     Set f = New frmMicroOrderUrine
49380     With f
49390         .txtSampleID = txtSampleID
49400         .Show 1
49410         Set Uxs = .UrineOrders
49420         txtSiteDetails.Text = Trim(.SiteDetails)
49430         DoEvents
49440     End With
49450     Unload f
49460     Set f = Nothing
49470     fraMicroscopy.Visible = True

49480     fraPregnancy.Visible = False
49490     If Not Uxs Is Nothing Then
49500         For Each Ux In Uxs
49510             Select Case UCase$(Ux.Request)
                      Case "PREGNANCY":
49520                     fraPregnancy.Visible = True
49530                     fraMicroscopy.Visible = False
                          '160       Case "CDIFF": SSTab1.TabVisible(Mic.CDiff) = True
                          '170       Case "OB0", "OB1", "OB2": SSTab1.TabVisible(Mic.FOB) = True
                          '180       Case "REDSUB": SSTab1.TabVisible(Mic.RedSub) = True
49540             End Select
49550         Next
49560     End If
49570     If LoadUrine() <> 0 Then
49580         SSTab1.TabCaption(Mic.Urine) = "<<Urine>>"
49590     End If

49600     cmdOrderTests.Enabled = True

49610     Exit Sub

OrderUrine_Error:

          Dim strES As String
          Dim intEL As Integer

49620     intEL = Erl
49630     strES = Err.Description
49640     LogError "frmEditMicrobiology", "OrderUrine", intEL, strES

End Sub
Private Sub PrintThis(PrintAction As String)

          Dim tb As Recordset
          Dim sql As String

49650     On Error GoTo PrintThis_Error

49660     pBar = 0

49670     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
49680         Exit Sub
49690     End If

49700     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

          '70    SaveDemographics

49710     sql = "Select * from PrintPending where " & _
              "Department = 'M' " & _
              "and SampleID = '" & txtSampleID & "'"
49720     Set tb = New Recordset
49730     RecOpenClient 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
49740     If tb.EOF Then
49750         tb.AddNew
49760     End If
49770     tb!SampleID = Val(txtSampleID.Text)
49780     tb!Ward = cmbWard
49790     tb!Clinician = cmbClinician
49800     tb!GP = cmbGP
49810     tb!Department = "M"
49820     tb!Initiator = UserName
49830     tb!PrintAction = PrintAction
49840     tb!UsePrinter = pPrintToPrinter
49850     tb.Update

49860     Exit Sub

PrintThis_Error:

          Dim strES As String
          Dim intEL As Integer

49870     intEL = Erl
49880     strES = Err.Description
49890     LogError "frmEditMicrobiology", "PrintThis", intEL, strES, sql

End Sub

Private Sub SaveComments(Optional CheckAutoVal As Boolean = True)

          Dim sql As String
          Dim tb As Recordset
          Dim OBs As Observations
          Dim MSCSave As String
          Dim ConCSave As String

49900     On Error GoTo SaveComments_Error

49910     txtSampleID = Format(Val(txtSampleID))
49920     If Val(txtSampleID) = 0 Then Exit Sub

49930     If chkPregnant.Value = 0 And CheckAutoVal Then CheckUrineAutoVal
49940     CheckMicroUrineComment txtSampleID


49950     If txtMSC = "Medical Scientist Comments" Then
49960         MSCSave = ""
49970     Else
49980         MSCSave = txtMSC
49990     End If
50000     If txtConC = "Consultant Comments" Then
50010         ConCSave = ""
50020     Else
50030         ConCSave = txtConC
50040     End If

50050     Set OBs = New Observations
          '+++ Junaid 09-05-2024
          '170   OBs.Save SampleIDWithOffset, True, _
          '               "Demographic", Trim$(txtDemographicComment), _
          '               "MicroGeneral", Trim$(txtUrineComment), _
          '               "MicroCS", Trim$(MSCSave), _
          '               "MicroConsultant", Trim$(ConCSave), _
          '               "MicroCSAutoComment", Trim$(txtCommentMicro)

50060     OBs.Save Trim(txtSampleID.Text), True, _
              "Demographic", Trim$(txtDemographicComment), _
              "MicroGeneral", Trim$(txtUrineComment), _
              "MicroCS", Trim$(MSCSave), _
              "MicroConsultant", Trim$(ConCSave), _
              "MicroCSAutoComment", Trim$(txtCommentMicro)
          '--- Junaid

50070     Exit Sub

SaveComments_Error:

          Dim strES As String
          Dim intEL As Integer

50080     intEL = Erl
50090     strES = Err.Description
50100     LogError "frmEditMicrobiology", "SaveComments", intEL, strES

End Sub

Private Sub SaveFaeces()

          Dim WS As FaecesWorkSheet

50110     On Error GoTo SaveFaeces_Error

50120     Set WS = New FaecesWorkSheet
50130     WS.Dayindex = "111": WS.Result = cmbDay1(11)
          '40    WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50140     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50150     WS.Save

50160     Set WS = New FaecesWorkSheet
50170     WS.Dayindex = "112": WS.Result = cmbDay1(12)
          '80    WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50180     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50190     WS.Save

50200     Set WS = New FaecesWorkSheet
50210     WS.Dayindex = "113": WS.Result = cmbDay1(13)
          '120   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50220     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50230     WS.Save

50240     Set WS = New FaecesWorkSheet
50250     WS.Dayindex = "121": WS.Result = cmbDay1(21)
          '160   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50260     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50270     WS.Save

50280     Set WS = New FaecesWorkSheet
50290     WS.Dayindex = "122": WS.Result = cmbDay1(22)
          '200   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50300     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50310     WS.Save

50320     Set WS = New FaecesWorkSheet
50330     WS.Dayindex = "123": WS.Result = cmbDay1(23)
          '240   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50340     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50350     WS.Save

50360     Set WS = New FaecesWorkSheet
50370     WS.Dayindex = "131": WS.Result = cmbDay1(31)
          '280   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50380     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50390     WS.Save

50400     Set WS = New FaecesWorkSheet
50410     WS.Dayindex = "132": WS.Result = cmbDay1(32)
          '320   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50420     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50430     WS.Save

50440     Set WS = New FaecesWorkSheet
50450     WS.Dayindex = "133": WS.Result = cmbDay1(33)
          '360   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50460     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50470     WS.Save




50480     Set WS = New FaecesWorkSheet
50490     WS.Dayindex = "211": WS.Result = cmbDay2(11)
          '400   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50500     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50510     WS.Save

50520     Set WS = New FaecesWorkSheet
50530     WS.Dayindex = "212": WS.Result = cmbDay2(12)
          '440   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50540     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50550     WS.Save

50560     Set WS = New FaecesWorkSheet
50570     WS.Dayindex = "213": WS.Result = cmbDay2(13)
          '480   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50580     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50590     WS.Save

50600     Set WS = New FaecesWorkSheet
50610     WS.Dayindex = "221": WS.Result = cmbDay2(21)
          '520   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50620     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50630     WS.Save

50640     Set WS = New FaecesWorkSheet
50650     WS.Dayindex = "222": WS.Result = cmbDay2(22)
          '560   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50660     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50670     WS.Save

50680     Set WS = New FaecesWorkSheet
50690     WS.Dayindex = "223": WS.Result = cmbDay2(23)
          '600   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50700     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50710     WS.Save

50720     Set WS = New FaecesWorkSheet
50730     WS.Dayindex = "231": WS.Result = cmbDay2(31)
          '640   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50740     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50750     WS.Save

50760     Set WS = New FaecesWorkSheet
50770     WS.Dayindex = "232": WS.Result = cmbDay2(32)
          '680   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50780     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50790     WS.Save

50800     Set WS = New FaecesWorkSheet
50810     WS.Dayindex = "233": WS.Result = cmbDay2(33)
          '720   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50820     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50830     WS.Save


50840     Set WS = New FaecesWorkSheet
50850     WS.Dayindex = "31": WS.Result = cmbDay3(1)
          '760   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50860     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50870     WS.Save

50880     Set WS = New FaecesWorkSheet
50890     WS.Dayindex = "32": WS.Result = cmbDay3(2)
          '800   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50900     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50910     WS.Save

50920     Set WS = New FaecesWorkSheet
50930     WS.Dayindex = "33": WS.Result = cmbDay3(3)
          '840   WS.SampleID = SampleIDWithOffset: WS.UserName = UserName
50940     WS.SampleID = Trim(txtSampleID.Text): WS.UserName = UserName
50950     WS.Save

50960     Exit Sub

SaveFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

50970     intEL = Erl
50980     strES = Err.Description
50990     LogError "frmEditMicrobiology", "SaveFaeces", intEL, strES

End Sub

Private Sub SaveOP(ByVal Valid As Integer)

          Dim n As Integer

51000     On Error GoTo SaveOP_Error

51010     SaveFaecesResult Valid, "AUS", Left$(lblCrypto, 1)

51020     For n = 0 To 2
51030         SaveFaecesResult Valid, "OP" & Format$(n), cmbOva(n)
51040     Next

51050     Exit Sub

SaveOP_Error:

          Dim strES As String
          Dim intEL As Integer

51060     intEL = Erl
51070     strES = Err.Description
51080     LogError "frmEditMicrobiology", "SaveOP", intEL, strES

End Sub


Private Sub SaveCdiff(ByVal Valid As Integer)

          'Dim Fx As New FaecesResult
          'Dim Fxs As New FaecesResults

51090     On Error GoTo SaveCdiff_Error

51100     SaveGenericResult Valid, "cDiffPCR", lblcDiffPCR.Caption

51110     SaveFaecesResult Valid, "ToxinAL", Left$(lblToxinA, 1)
          '80    Fx.SampleID = SampleIDWithOffset
          '90    Fx.TestName = "ToxinAL"
          '100   Fx.Result = Left$(lblToxinA, 1)
          '110   Fx.UserName = UserName
          '120   Fx.Valid = Valid
          '130   Fxs.Save Fx
51120     SaveFaecesResult Valid, "ToxinATA", Left$(lblToxinB, 1)

          '140   Fx.TestName = "ToxinATA"
          '150   Fx.Result = Left$(lblToxinB, 1)
          '160   Fxs.Save Fx

51130     Exit Sub

SaveCdiff_Error:

          Dim strES As String
          Dim intEL As Integer

51140     intEL = Erl
51150     strES = Err.Description
51160     LogError "frmEditMicrobiology", "SaveCdiff", intEL, strES

End Sub


Private Sub SaveRotaAdeno(ByVal Valid As Integer)

          'Dim Fx As New FaecesResult
          'Dim Fxs As New FaecesResults

51170     On Error GoTo SaveRotaAdeno_Error

51180     SaveFaecesResult Valid, "Rota", Left$(txtRota, 1)
          '20    Fx.SampleID = SampleIDWithOffset
          '30    Fx.TestName = "Rota"
          '40    Fx.Result = Left$(txtRota, 1)
          '50    Fx.UserName = UserName
          '60    Fxs.Save Fx
51190     SaveFaecesResult Valid, "Adeno", Left$(txtAdeno, 1)
          '
          '70    Fx.TestName = "Adeno"
          '80    Fx.Result = Left$(txtAdeno, 1)
          '90    Fxs.Save Fx

51200     Exit Sub

SaveRotaAdeno_Error:

          Dim strES As String
          Dim intEL As Integer

51210     intEL = Erl
51220     strES = Err.Description
51230     LogError "frmEditMicrobiology", "SaveRotaAdeno", intEL, strES

End Sub


Private Sub SaveFOB(ByVal Valid As Integer)

          Dim n As Integer
          '      Dim Fx As FaecesResult
          '      Dim Fxs As New FaecesResults

51240     On Error GoTo SaveFOB_Error

51250     For n = 0 To 2
51260         SaveFaecesResult Valid, "OB" & Format$(n), Left$(lblFOB(n), 1)
              'If chkFOB(n) Then
              '40        Set Fx = New FaecesResult
              '50        Fx.SampleID = SampleIDWithOffset
              '60        Fx.TestName = "OB" & Format$(n)
              '70        Fx.Result = Left$(lblFOB(n), 1)
              '80        Fx.UserName = UserName
              '90        Fxs.Save Fx
              '100     End If
51270     Next

51280     Exit Sub

SaveFOB_Error:

          Dim strES As String
          Dim intEL As Integer

51290     intEL = Erl
51300     strES = Err.Description
51310     LogError "frmEditMicrobiology", "SaveFOB", intEL, strES

End Sub

Private Sub SaveCSF()

          Dim tb As Recordset
          Dim sql As String
          Dim Found As Boolean
          Dim n As Integer

51320     On Error GoTo SaveCSF_Error

          '+++ Junaid 08-05-2024
51330     If txtSampleID.Text = "" Then
51340         Exit Sub
51350     End If
          '--- Junaid
51360     Found = False
51370     If Trim$(cmbCSFGram & txtCSFWCCDiff(0) & txtCSFWCCDiff(1)) <> "" Then
51380         Found = True
51390     End If
51400     For n = 0 To 2
51410         If Trim$(cmbCSFAppearance(n) & txtCSFWCC(n) & txtCSFRCC(n)) <> "" Then
51420             Found = True
51430         End If
51440     Next

51450     If Found Then

              '+++ Junaid 20-05-2024
              '120       sql = "SELECT * FROM CSFResults WHERE " & _
              '                "SampleID = '" & SampleIDWithOffset & "'"
51460         sql = "SELECT * FROM CSFResults WHERE " & _
                  "SampleID = '" & Trim(txtSampleID.Text) & "'"
              '--- Junaid
51470         Set tb = New Recordset
51480         RecOpenServer 0, tb, sql

51490         If tb.EOF Then tb.AddNew

              '+++ Junaid 08-05-2024
              '160       tb!SampleID = SampleIDWithOffset
51500         tb!SampleID = Val(txtSampleID.Text)
              '--- Junaid
51510         tb!Gram = cmbCSFGram
51520         tb!WCCDiff0 = txtCSFWCCDiff(0)
51530         tb!WCCDiff1 = txtCSFWCCDiff(1)
51540         tb!Appearance0 = cmbCSFAppearance(0)
51550         tb!Appearance1 = cmbCSFAppearance(1)
51560         tb!Appearance2 = cmbCSFAppearance(2)
51570         tb!WCC0 = txtCSFWCC(0)
51580         tb!WCC1 = txtCSFWCC(1)
51590         tb!WCC2 = txtCSFWCC(2)
51600         tb!RCC0 = txtCSFRCC(0)
51610         tb!RCC1 = txtCSFRCC(1)
51620         tb!RCC2 = txtCSFRCC(2)
51630         tb!UserName = UserName

51640         tb.Update

51650     End If

51660     Exit Sub

SaveCSF_Error:

          Dim strES As String
          Dim intEL As Integer

51670     intEL = Erl
51680     strES = Err.Description
51690     LogError "frmEditMicrobiology", "SaveCSF", intEL, strES, sql

End Sub


Private Sub SaveRSV(ByVal Valid As Integer)

51700     SaveGenericResult Valid, "RSV", lblRSV.Caption

End Sub


Private Sub SaveGenericResult(ByVal Valid As Integer, ByVal Analyte As String, ByVal Result As String)

          Dim Gx As GenericResult
          Dim GXs As New GenericResults
          Dim v As Integer
          Dim VBy As String
          Dim VDT As String

51710     On Error GoTo SaveGenericResult_Error

51720     v = 0
51730     VBy = ""
51740     VDT = ""
          '+++ Junaid 08-05-2024
51750     If txtSampleID.Text = "" Then
51760         Exit Sub
51770     End If
          '--- Junaid
51780     If Valid = gDONTCARE Then
              '+++ Junaid 08-05-2024
              '60        GXs.Load SampleIDWithOffset
51790         GXs.Load Val(txtSampleID.Text)
              '--- Junaid
51800         If GXs.Count > 0 Then
51810             Set Gx = GXs(Analyte)
51820             If Not Gx Is Nothing Then
51830                 v = Gx.Valid
51840                 VBy = Gx.ValidatedBy
51850                 VDT = Gx.ValidatedDateTime
51860             End If
51870         End If
51880     Else
51890         If Valid = gVALID Then
51900             v = 1
51910             VBy = UserName
51920             VDT = Now
51930         End If
51940     End If

51950     Set Gx = New GenericResult
          '+++ Junaid 08-05-2024
          '230   Gx.SampleID = SampleIDWithOffset
51960     Gx.SampleID = Val(txtSampleID.Text)
          '--- Junaid
51970     Gx.TestName = Analyte
51980     Gx.Result = Result
51990     Gx.UserName = UserName
52000     Gx.Valid = v
52010     Gx.ValidatedBy = VBy
52020     Gx.ValidatedDateTime = VDT
52030     GXs.Save Gx

52040     Exit Sub

SaveGenericResult_Error:

          Dim strES As String
          Dim intEL As Integer

52050     intEL = Erl
52060     strES = Err.Description
52070     LogError "frmEditMicrobiology", "SaveGenericResult", intEL, strES

End Sub

Private Sub SaveFaecesResult(ByVal Valid As Integer, ByVal Analyte As String, ByVal Result As String)

          Dim Fx As FaecesResult
          Dim Fxs As New FaecesResults
          Dim v As Integer
          Dim VBy As String
          Dim VDT As String

52080     On Error GoTo SaveFaecesResult_Error

52090     v = 0
52100     VBy = ""
52110     VDT = ""
          '+++ Junaid 08-05-2024
52120     If txtSampleID.Text = "" Then
52130         Exit Sub
52140     End If
          '--- Junaid
52150     If Valid = gDONTCARE Then
              '+++ Junaid 08-05-2024
              '60        Fxs.Load SampleIDWithOffset
52160         Fxs.Load Val(txtSampleID.Text)
              '--- Junaid
52170         If Fxs.Count > 0 Then
52180             Set Fx = Fxs(Analyte)
52190             If Not Fx Is Nothing Then
52200                 v = Fx.Valid
52210                 VBy = Fx.ValidatedBy
52220                 VDT = Fx.ValidatedDateTime
52230             End If
52240         End If
52250     Else
52260         If Valid = gVALID Then
52270             v = 1
52280             VBy = UserName
52290             VDT = Now
52300         End If
52310     End If

52320     Set Fx = New FaecesResult
          '+++ Junaid 08-05-2024
          '230   Fx.SampleID = SampleIDWithOffset
52330     Fx.SampleID = Val(txtSampleID.Text)
          '--- Junaid
52340     Fx.TestName = Analyte
52350     Fx.Result = Result
52360     Fx.UserName = UserName
52370     Fx.Valid = v
52380     Fx.ValidatedBy = VBy
52390     Fx.ValidatedDateTime = VDT
52400     Fxs.Save Fx

52410     Exit Sub

SaveFaecesResult_Error:

          Dim strES As String
          Dim intEL As Integer

52420     intEL = Erl
52430     strES = Err.Description
52440     LogError "frmEditMicrobiology", "SaveFaecesResult", intEL, strES

End Sub


Private Sub SaveRedSub(ByVal Valid As Integer)

          Dim n As Integer
          Dim Result As String

52450     On Error GoTo SaveRedSub_Error

52460     Result = ""
52470     For n = 0 To 5
52480         If chkRS(n).Value = 1 Then
52490             Result = chkRS(n).Caption
52500             Exit For
52510         End If
52520     Next

52530     SaveGenericResult Valid, "RedSub", Result

52540     Exit Sub

SaveRedSub_Error:

          Dim strES As String
          Dim intEL As Integer

52550     intEL = Erl
52560     strES = Err.Description
52570     LogError "frmEditMicrobiology", "SaveRedSub", intEL, strES

End Sub

Private Sub SaveIdentification()

          Dim Ix As UrineIdent
          Dim Ixs As New UrineIdents
          Dim n As Integer
          Dim sql As String

52580     On Error GoTo SaveIdentification_Error


          '+++ Junaid 26-12-2023
52590     Ixs.Clear
52600     sql = "Insert Into UrineIdent50Audit(SampleID, TestName, Result, Isolate, UserName, DateTimeOfRecord, ArchivedBy, ArchiveDateTime) "
52610     sql = sql & "(Select SampleID, TestName, Result, Isolate, UserName, DateTimeOfRecord, '" & UserName & "', GetDate() From UrineIdent50 Where SampleID = '" & txtSampleID.Text & "')"
52620     Cnxn(0).Execute sql: Call WriteToFile_Execution(sql)
          '      --- Junaid
52630     For n = 1 To 4
52640         Set Ix = New UrineIdent
52650         Ix.SampleID = txtSampleID.Text
52660         Ix.Isolate = n
52670         Ix.TestName = "Notes"
52680         Ix.Result = Trim$(txtIdentification(n))
52690         Ix.UserName = UserName
52700         Ixs.Save Ix
52710     Next

52720     Exit Sub

SaveIdentification_Error:

          Dim strES As String
          Dim intEL As Integer

52730     intEL = Erl
52740     strES = Err.Description
52750     LogError "frmEditMicrobiology", "SaveIdentification", intEL, strES

End Sub

Private Sub SaveIsolates()

          Dim intIsolate As Integer
          Dim Isos As New Isolates
          Dim Iso As Isolate
          Dim Sxs As New Sensitivities
          Dim sx As Sensitivity

52760     On Error GoTo SaveIsolates_Error

          '+++ Junaid 08-05-2024
          '20    Sxs.Load SampleIDWithOffset
52770     If txtSampleID.Text = "" Then
52780         Exit Sub
52790     End If
52800     Sxs.Load Val(txtSampleID.Text)
          '--- Junaid

52810     For intIsolate = 1 To 4
52820         If cmbOrgGroup(intIsolate) = "" Then
                  '+++ Junaid 09-05-2024
                  '50            Isos.Delete SampleIDWithOffset, intIsolate
52830             Isos.Delete Trim(txtSampleID.Text), intIsolate
                  '--- Junaid
52840             For Each sx In Sxs
52850                 If sx.IsolateNumber = intIsolate Then
52860                     Sxs.Delete sx
52870                 End If
52880             Next
52890         Else
52900             Set Iso = New Isolate
52910             Iso.IsolateNumber = intIsolate
52920             Iso.OrganismGroup = cmbOrgGroup(intIsolate)
52930             Iso.OrganismName = cmbOrgName(intIsolate)
52940             Iso.Qualifier = cmbQualifier(intIsolate)
                  '+++ Junaid 09-05-2024
                  '170           Iso.SampleID = SampleIDWithOffset
52950             Iso.SampleID = Trim(txtSampleID.Text)
                  '--- Junaid
52960             Iso.UserName = UserName
52970             Isos.Add Iso
52980         End If
52990     Next
53000     Isos.Save

53010     Exit Sub

SaveIsolates_Error:

          Dim strES As String
          Dim intEL As Integer

53020     intEL = Erl
53030     strES = Err.Description
53040     LogError "frmEditMicrobiology", "SaveIsolates", intEL, strES

End Sub

Private Sub SaveUrine(ByVal Validate As Boolean)

          Dim U As New UrineResult
          Dim Us As New UrineResults
          Dim n As Integer

          Dim TestNames(1 To 15) As String
          Dim txtRs(1 To 15) As TextBox
          Dim cmbRs(1 To 5) As ComboBox

53050     On Error GoTo SaveUrine_Error
          '+++ Junaid 08-05-2024
          '20    U.SampleID = SampleIDWithOffset
53060     If txtSampleID.Text = "" Then
53070         Exit Sub
53080     End If
53090     U.SampleID = Val(txtSampleID.Text)
          '--- Junaid
53100     U.UserName = UserName
53110     If Validate Then
53120         U.Valid = 1
53130         U.ValidatedBy = UserName
53140         U.ValidatedDateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
53150     Else
53160         U.Valid = 0
53170         U.ValidatedBy = ""
53180         U.ValidatedDateTime = ""
53190     End If

53200     For n = 1 To 15
53210         TestNames(n) = Choose(n, "Pregnancy", "Bacteria", "HCGLevel", "BenceJones", _
                  "SG", "FatGlobules", "pH", "Protein", "Glucose", "Ketones", _
                  "Urobilinogen", "Bilirubin", "BloodHb", "WCC", "RCC")
53220         Set txtRs(n) = Choose(n, txtPregnancy, txtBacteria, txtHCGLevel, txtBenceJones, _
                  txtSG, txtFatGlobules, txtpH, txtProtein, txtGlucose, txtKetones, _
                  txtUrobilinogen, txtBilirubin, txtBloodHb, txtWCC, txtRCC)

53230         U.TestName = TestNames(n)
53240         U.Result = txtRs(n).Text
53250         Us.Save U
53260     Next

53270     For n = 1 To 5
53280         TestNames(n) = Choose(n, "Crystals", "Casts", "Misc0", "Misc1", "Misc2")
53290         Set cmbRs(n) = Choose(n, cmbCrystals, cmbCasts, cmbMisc(0), cmbMisc(1), cmbMisc(2))
53300         U.TestName = TestNames(n)
53310         U.Result = cmbRs(n).Text
53320         Us.Save U
53330     Next

53340     U.TestName = "TagRepeat"
53350     U.Result = cmdTagRepeat.Caption
53360     Us.Save U

53370     U.TestName = "AutoVal"
53380     U.Result = UrineAutoVal
53390     Us.Save U

53400     Exit Sub

SaveUrine_Error:

          Dim strES As String
          Dim intEL As Integer

53410     intEL = Erl
53420     strES = Err.Description
53430     LogError "frmEditMicrobiology", "SaveUrine", intEL, strES

End Sub

Private Sub chkFOB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

53440     ShowUnlock Mic.FOB

End Sub


Private Sub chkPregnant_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

53450     cmdSaveDemographics.Enabled = True

End Sub

Private Sub chkRS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

          Dim n As Integer
          Dim intOriginal As Integer

53460     intOriginal = chkRS(Index).Value

53470     For n = 0 To 5
53480         chkRS(n).Value = 0
53490     Next

53500     chkRS(Index).Value = intOriginal

53510     ShowUnlock Mic.RedSub

End Sub


Private Sub cmbABSelect_Click(Index As Integer)

          Dim sx As New Sensitivity
          Dim SxTrial As Sensitivity
          Dim RSI As String
          Dim Result As String
          Dim Rundate As String
          Dim RunDateTime As String

53520     On Error GoTo cmbABSelect_Click_Error

53530     SaveSensitivities 0

53540     RSI = ""
53550     Result = ""
53560     Rundate = Format$(Now, "dd/MMM/yyyy")
53570     RunDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")

53580     For Each SxTrial In CurrentSensitivities
53590         If SxTrial.AntibioticName = cmbABSelect(Index).Text And SxTrial.IsolateNumber = Index Then
53600             RSI = SxTrial.RSI
53610             Result = SxTrial.Result
53620             Rundate = SxTrial.Rundate
53630             RunDateTime = SxTrial.RunDateTime
53640             Exit For
53650         End If
53660     Next

53670     sx.AntibioticName = cmbABSelect(Index).Text
53680     sx.AntibioticCode = AntibioticCodeFor(sx.AntibioticName)
53690     sx.Forced = True
53700     sx.RSI = RSI
53710     sx.Result = Result
53720     sx.Report = True
53730     sx.IsolateNumber = Index
53740     sx.Rundate = Rundate
53750     sx.RunDateTime = RunDateTime
          '+++ Junaid 20-05-2024
          '250   sx.SampleID = SampleIDWithOffset
53760     sx.SampleID = Trim(txtSampleID.Text)
          '--- Junaid
53770     sx.UserCode = UserCode
53780     sx.Valid = False
53790     sx.Save

53800     Set CurrentSensitivities = New Sensitivities
          '+++ Junaid 20-05-2024
          '300   CurrentSensitivities.Load SampleIDWithOffset
53810     CurrentSensitivities.Load Trim(txtSampleID.Text)
          '--- Junaid

53820     LoadSensitivities

53830     Exit Sub

cmbABSelect_Click_Error:

          Dim strES As String
          Dim intEL As Integer

53840     intEL = Erl
53850     strES = Err.Description
53860     LogError "frmEditMicrobiology", "cmbABSelect_Click", intEL, strES

End Sub

Private Sub cmbABSelect_KeyPress(Index As Integer, KeyAscii As Integer)

53870     KeyAscii = 0

End Sub

Private Sub cmbABsInUse_Click()

          Dim n As Integer

53880     lstABsInUse.AddItem cmbABsInUse
53890     cmbABsInUse.Visible = False
53900     lstABsInUse.Visible = True

53910     lblABsInUse = ""
53920     For n = 0 To lstABsInUse.ListCount - 1
53930         lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
53940     Next

End Sub


Private Sub cmbCasts_Click()

53950     ShowUnlock Mic.Urine

End Sub


Private Sub cmbCasts_LostFocus()

          Dim tb As Recordset
          Dim sql As String

53960     On Error GoTo cmbCasts_LostFocus_Error

53970     sql = "Select * from Lists where " & _
              "ListType = 'CA' and InUse = 1 " & _
              "and Code = '" & UCase(cmbCasts) & "'"
53980     Set tb = New Recordset
53990     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
54000     If Not tb.EOF Then
54010         cmbCasts = tb!Text & ""
54020     End If

54030     Exit Sub

cmbCasts_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

54040     intEL = Erl
54050     strES = Err.Description
54060     LogError "frmEditMicrobiology", "cmbCasts_LostFocus", intEL, strES, sql

End Sub

Private Sub cmbConC_Click()

54070     If txtConC = "Consultant Comments" Then
54080         txtConC = ""
54090     End If

54100     txtConC = txtConC & cmbConC
54110     txtConC.SetFocus
54120     txtConC.SelStart = Len(txtConC)
54130     cmbConC.Visible = False

54140     cmdSaveMicro.Enabled = True
54150     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbConC_LostFocus()

          Dim sql As String
          Dim tb As Recordset

54160     On Error GoTo cmbConC_LostFocus_Error

54170     sql = "Select * from Lists where " & _
              "ListType = 'BA' and InUse = 1 " & _
              "AND Code = '" & AddTicks(cmbConC) & "'"
54180     Set tb = New Recordset
54190     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
54200     If Not tb.EOF Then
54210         txtConC.Text = Trim$(txtConC.Text & " " & tb!Text & "")
54220     End If

54230     cmbConC.Visible = False

54240     Exit Sub

cmbConC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

54250     intEL = Erl
54260     strES = Err.Description
54270     LogError "frmEditMicrobiology", "cmbConC_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbConsultantVal_Click()

54280     txtSampleID = cmbConsultantVal
54290     txtSampleID = Format$(Val(txtSampleID))
54300     If txtSampleID = 0 Then Exit Sub

          '40    If Not GetSampleIDWithOffset Then Exit Sub

54310     LoadAllDetails

54320     cmdSaveDemographics.Enabled = False
54330     cmdSaveInc.Enabled = False
54340     cmdSaveMicro.Enabled = False
54350     cmdSaveHold.Enabled = False

End Sub

Private Sub cmbCrystals_Click()

54360     ShowUnlock Mic.Urine

End Sub


Private Sub cmbCrystals_LostFocus()

          Dim tb As Recordset
          Dim sql As String

54370     On Error GoTo cmbCrystals_LostFocus_Error

54380     sql = "Select * from Lists where " & _
              "ListType = 'CR' " & _
              "and Code = '" & UCase(cmbCrystals) & "' and InUse = 1"
54390     Set tb = New Recordset
54400     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
54410     If Not tb.EOF Then
54420         cmbCrystals = tb!Text & ""
54430     End If

54440     Exit Sub

cmbCrystals_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

54450     intEL = Erl
54460     strES = Err.Description
54470     LogError "frmEditMicrobiology", "cmbCrystals_LostFocus", intEL, strES, sql

End Sub

Private Sub cmbCSFAppearance_Click(Index As Integer)

54480     cmdSaveMicro.Enabled = True
54490     cmdSaveHold.Enabled = True

54500     ShowUnlock Mic.CSF

End Sub

Private Sub cmbCSFGram_Click()

54510     cmdSaveMicro.Enabled = True
54520     cmdSaveHold.Enabled = True

54530     ShowUnlock Mic.CSF

End Sub

Private Sub cmbDay1_Click(Index As Integer)

54540     cmdSaveMicro.Enabled = True
54550     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDay1_KeyPress(Index As Integer, KeyAscii As Integer)

54560     cmdSaveMicro.Enabled = True
54570     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDay2_Click(Index As Integer)

54580     cmdSaveMicro.Enabled = True
54590     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDay2_KeyPress(Index As Integer, KeyAscii As Integer)

54600     cmdSaveMicro.Enabled = True
54610     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDay3_Click(Index As Integer)

54620     cmdSaveMicro.Enabled = True
54630     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDay3_KeyPress(Index As Integer, KeyAscii As Integer)

54640     cmdSaveMicro.Enabled = True
54650     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDemogComment_Click()

54660     txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
54670     cmbDemogComment = ""

54680     cmdSaveDemographics.Enabled = True
54690     cmdSaveInc.Enabled = True

End Sub


Private Sub cmbDemogComment_KeyPress(KeyAscii As Integer)

54700     KeyAscii = 0

End Sub


Private Sub cmbDemogComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

54710     On Error GoTo cmbDemogComment_LostFocus_Error

54720     sql = "Select * from Lists where " & _
              "ListType = 'DE' " & _
              "and Code = '" & cmbDemogComment & "' and InUse = 1"
54730     Set tb = New Recordset
54740     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
54750     If Not tb.EOF Then
54760         txtDemographicComment = Trim$(txtDemographicComment & " " & tb!Text & "")
54770     Else
54780         txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
54790     End If
54800     cmbDemogComment = ""

54810     Exit Sub

cmbDemogComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

54820     intEL = Erl
54830     strES = Err.Description
54840     LogError "frmEditMicrobiology", "cmbDemogComment_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbHospital_Click()

54850     FillWards cmbWard, cmbHospital
54860     FillClinicians cmbClinician, cmbHospital
54870     FillGPs cmbGP, cmbHospital

54880     cmdSaveDemographics.Enabled = True
54890     cmdSaveInc.Enabled = True

End Sub


Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

54900     KeyAscii = 0

End Sub


Private Sub cmbIdentification_Click(Index As Integer)

54910     txtIdentification(Index) = txtIdentification(Index) & cmbIdentification(Index)

54920     cmdSaveMicro.Enabled = True
54930     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbIdentification_KeyPress(Index As Integer, KeyAscii As Integer)
54940     cmdSaveMicro.Enabled = True
54950     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbIdentification_LostFocus(Index As Integer)

          Dim sql As String
          Dim tb As Recordset

54960     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'IN' " & _
              "AND Code = '" & AddTicks(cmbIdentification(Index)) & "'"
54970     Set tb = New Recordset
54980     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
54990     If Not tb.EOF Then
55000         cmbIdentification(Index) = tb!Text & ""
55010     End If

End Sub

Private Sub cmbMisc_Click(Index As Integer)

55020     ShowUnlock Mic.Urine

End Sub


Private Sub cmbMisc_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

55030     On Error GoTo cmbMisc_LostFocus_Error

55040     sql = "Select * from Lists where " & _
              "ListType = 'MI' " & _
              "and Code = '" & cmbMisc(Index) & "' and InUse = 1"
55050     Set tb = New Recordset
55060     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
55070     If Not tb.EOF Then
55080         cmbMisc(Index) = tb!Text & ""
55090     End If

55100     Exit Sub

cmbMisc_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

55110     intEL = Erl
55120     strES = Err.Description
55130     LogError "frmEditMicrobiology", "cmbMisc_LostFocus", intEL, strES, sql

End Sub

Private Sub cmbMSC_Click()

55140     If txtMSC() = "Medical Scientist Comments" Then
55150         txtMSC = ""
55160     End If

55170     txtMSC = txtMSC & cmbMSC
55180     txtMSC.SetFocus
55190     txtMSC.SelStart = Len(txtMSC)
55200     cmbMSC.Visible = False

55210     cmdSaveMicro.Enabled = True
55220     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbMSC_LostFocus()

          Dim sql As String
          Dim tb As Recordset

55230     On Error GoTo cmbMSC_LostFocus_Error

55240     sql = "Select * from Lists where " & _
              "ListType = 'BA' and InUse = 1 " & _
              "AND Code = '" & AddTicks(cmbMSC) & "'"
55250     Set tb = New Recordset
55260     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
55270     If Not tb.EOF Then
55280         txtMSC.Text = Trim$(txtMSC.Text & " " & tb!Text & "")
55290     End If

55300     cmbMSC.Visible = False

55310     Exit Sub

cmbMSC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

55320     intEL = Erl
55330     strES = Err.Description
55340     LogError "frmEditMicrobiology", "cmbMSC_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbOrgGroup_Click(Index As Integer)

          Dim sx As Sensitivity

55350     If Trim$(cmbOrgGroup(Index)) <> "" Then
55360         For Each sx In CurrentSensitivities
55370             If sx.IsolateNumber = Index Then
55380                 CurrentSensitivities.Delete sx
55390             End If
55400         Next
55410         cmbOrgName(Index) = ""
55420         cmbQualifier(Index) = ""
55430         grdAB(Index).Rows = 2
55440         grdAB(Index).AddItem ""
55450         grdAB(Index).RemoveItem 1
55460         FillAbGrid Index
55470     End If

55480     FillABSelect Index
55490     FillOrgNames Index

55500     cmdSaveMicro.Enabled = True
55510     cmdSaveHold.Enabled = True
55520     grdAB(Index).Visible = True
55530     Call ABExistsInCurrent(Index)
End Sub

Private Sub cmbOrgGroup_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim sx As Sensitivity

55540     On Error GoTo cmbOrgGroup_LostFocus_Error

55550     If Trim$(cmbOrgGroup(Index)) = "" Then
55560         For Each sx In CurrentSensitivities
55570             If sx.IsolateNumber = Index Then
55580                 CurrentSensitivities.Delete sx
55590             End If
55600         Next
55610         cmbOrgName(Index) = ""
55620         cmbQualifier(Index) = ""
55630         grdAB(Index).Rows = 2
55640         grdAB(Index).AddItem ""
55650         grdAB(Index).RemoveItem 1
55660         FillAbGrid Index
55670     End If

55680     sql = "Select * from Lists where " & _
              "ListType = 'OR' " & _
              "and Code = '" & cmbOrgGroup(Index) & "' and InUse = 1"
55690     Set tb = New Recordset
55700     RecOpenServer 0, tb, sql
          '+++ Junaid 12-12-2023
          '
          '--- Junaid
55710     If Not tb.EOF Then
55720         cmbOrgGroup(Index) = tb!Text & ""
55730     End If

55740     Exit Sub

cmbOrgGroup_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

55750     intEL = Erl
55760     strES = Err.Description
55770     LogError "frmEditMicrobiology", "cmbOrgGroup_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbOrgName_Click(Index As Integer)

55780     cmdSaveMicro.Enabled = True
55790     cmdSaveHold.Enabled = True
          
End Sub


Private Sub cmbOva_Click(Index As Integer)

55800     ShowUnlock Mic.OP

End Sub

Private Sub cmbOva_KeyPress(Index As Integer, KeyAscii As Integer)

55810     KeyAscii = 0

End Sub


Private Sub cmbQualifier_Click(Index As Integer)

55820     cmdSaveMicro.Enabled = True
55830     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbSite_Change()

55840     lblSiteDetails = cmbSite & " " & Trim(txtSiteDetails.Text)

55850     cmdOrderTests.Enabled = False
55860     If cmbSite = "Faeces" Or cmbSite = "Urine" Then
55870         cmdOrderTests.Enabled = True
55880     End If

55890     EnableTagRepeat

End Sub

Private Sub cmbSite_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim Found As Boolean
          Dim tb As Recordset
          Dim sql As String

55900     On Error GoTo cmbSite_KeyUp_Error

55910     If Trim$(cmbSite) = "" Then
55920         Exit Sub
55930     End If

55940     Found = False
55950     sql = "Select * from Lists where " & _
              "ListType = 'SI' " & _
              "and ( Text = '" & cmbSite & "' " & _
              "or Code = '" & cmbSite & "') and InUse = 1"
55960     Set tb = New Recordset
55970     RecOpenServer 0, tb, sql
55980     If tb.EOF Then
55990         cmbSite = ""
56000     End If

56010     cmbSiteEffects

56020     EnableTagRepeat

56030     Exit Sub

cmbSite_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

56040     intEL = Erl
56050     strES = Err.Description
56060     LogError "frmEditMicrobiology", "cmbSite_KeyUp", intEL, strES, sql

End Sub


Private Sub cmdABsInUse_Click()

56070     lstABsInUse.Visible = False
56080     cmbABsInUse.Visible = True
56090     cmbABsInUse.SetFocus

56100     cmdSaveDemographics.Enabled = True
56110     cmdSaveInc.Enabled = True

End Sub

Private Sub cmdConC_Click()

56120     cmbConC.Visible = True
56130     cmbConC.SetFocus

End Sub

Private Sub cmdCopyFromPrevious_Click()
    'Zyam 02-5-24

    '      Dim tb As Recordset
    '      Dim sql As String
    '      Dim PrevSID As Long
    '      Dim OBs As Observations
    '
    '10    On Error GoTo cmdCopyFromPrevious_Click_Error
    '
    '20    PrevSID = sysOptMicroOffset(0) + Val(txtSampleID) - 1
    '    '+++Junaid 15-10-2023
    '    If MsgBox("Are you sure you want to copy all details from " & PrevSID & "?", vbInformation + vbYesNo) = vbNo Then
    '        Exit Sub
    '    End If
    '    '---Junaid
    '30    sql = "Select * from Demographics where " & _
    '            "SampleID = " & PrevSID
    '40    Set tb = New Recordset
    '50    RecOpenServer 0, tb, sql
    '
    '60    If Trim$(tb!Hospital & "") <> "" Then
    '70        cmbHospital = Trim$(tb!Hospital)
    '80        lblChartNumber = Trim$(tb!Hospital) & " Chart #"
    '90        If UCase$(tb!Hospital) = UCase$(HospName(0)) Then
    '100           lblChartNumber.BackColor = &H8000000F
    '110           lblChartNumber.ForeColor = vbBlack
    '120       Else
    '130           lblChartNumber.BackColor = vbRed
    '140           lblChartNumber.ForeColor = vbYellow
    '150       End If
    '160   Else
    '170       cmbHospital = HospName(0)
    '180       lblChartNumber.Caption = HospName(0) & " Chart #"
    '190       lblChartNumber.BackColor = &H8000000F
    '200       lblChartNumber.ForeColor = vbBlack
    '210   End If
    '220   If IsDate(tb!SampleDate) Then
    '230       dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
    '240   Else
    '250       dtSampleDate = Format$(Now, "dd/mm/yyyy")
    '260   End If
    '270   If IsDate(tb!Rundate) Then
    '280       dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
    '290   Else
    '300       dtRunDate = Format$(Now, "dd/mm/yyyy")
    '310   End If
    '320   mNewRecord = False
    '330   If Not IsNull(tb!RooH) Then
    '340       cRooH(0) = IIf(tb!RooH, True, False)
    '350       cRooH(1) = Not tb!RooH
    '360   Else
    '370       cRooH(0) = True
    '380   End If
    '390   txtChart = tb!Chart & ""
    '400   txtLabNo = tb!LabNo & ""
    '410   txtSurName = SurName(tb!PatName & "")
    '420   txtForeName = ForeName(tb!PatName & "")
    '430   txtAddress(0) = tb!Addr0 & ""
    '440   txtAddress(1) = tb!Addr1 & ""
    '450   Select Case Left$(Trim$(UCase$(tb!Sex & "")), 1)
    '      Case "M": txtSex = "Male"
    '460   Case "F": txtSex = "Female"
    '470   Case Else: txtSex = ""
    '480   End Select
    '490   txtDoB = Format$(tb!DoB, "dd/mm/yyyy")
    '500   txtAge = tb!Age & ""
    '510   cmbWard = tb!Ward & ""
    '520   cmbClinician = tb!Clinician & ""
    '530   cmbGP = tb!GP & ""
    '540   txtClinDetails = tb!ClDetails & ""
    '550   If IsDate(tb!SampleDate) Then
    '560       dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
    '570       If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
    '580           tSampleTime = Format$(tb!SampleDate, "hh:mm")
    '590       Else
    '600           tSampleTime.Mask = ""
    '610           tSampleTime.Text = ""
    '620           tSampleTime.Mask = "##:##"
    '630       End If
    '640   Else
    '650       dtSampleDate = Format$(Now, "dd/mm/yyyy")
    '660       tSampleTime.Mask = ""
    '670       tSampleTime.Text = ""
    '680       tSampleTime.Mask = "##:##"
    '690   End If
    '700   If IsDate(tb!RecDate & "") Then
    '710       dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
    '720       If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
    '730           tRecTime = Format$(tb!RecDate, "hh:mm")
    '740       Else
    '750           tRecTime.Mask = ""
    '760           tRecTime.Text = ""
    '770           tRecTime.Mask = "##:##"
    '780       End If
    '790   Else
    '800       dtRecDate = Format$(Now, "dd/mm/yyyy")
    '810       tRecTime.Mask = ""
    '820       tRecTime.Text = ""
    '830       tRecTime.Mask = "##:##"
    '840   End If
    '
    '850   cmdSaveDemographics.Enabled = True
    '860   cmdSaveInc.Enabled = True
    '
    '870   If sysOptBloodBank(0) Then
    '880       If Trim$(txtChart) <> "" Then
    '890           sql = "Select  * from PatientDetails where " & _
    '                    "PatNum = '" & txtChart & "'"
    '900           Set tb = New Recordset
    '910           RecOpenClientBB 0, tb, sql
    '920           bViewBB.Enabled = Not tb.EOF
    '930       End If
    '940   End If
    '
    '950   Set OBs = New Observations
    '960   Set OBs = OBs.Load(PrevSID, "Demographic")
    '970   If Not OBs Is Nothing Then
    '980       txtDemographicComment = OBs.Item(1).Comment
    '990   End If
    '
    '      'Dim SDS As New SiteDetails
    '      '1000  SDS.Load PrevSID
    '      '1010  If SDS.Count > 0 Then
    '      '1020    cmbSite = SDS(1).Site
    '      '1030    txtSiteDetails = SDS(1).SiteDetails
    '      '1040  End If
    '1000  cmbSite = ""
    '1010  txtSiteDetails = ""
    '
    '      Dim CURS As New CurrentAntibiotics
    '      Dim Cur As CurrentAntibiotic
    '1020  CURS.Load PrevSID
    '1030  lblABsInUse = ""
    '1040  For Each Cur In CURS
    '1050      lstABsInUse.AddItem Cur.Antibiotic
    '1060      lblABsInUse = lblABsInUse & Cur.Antibiotic & " "
    '1070  Next
    '
    '1080  CopyCC (PrevSID)
    '1090  CheckCC
    '
    '1100  cmdCopyFromPrevious.Visible = False
    '
    '1110  Exit Sub
    '
    'cmdCopyFromPrevious_Click_Error:
    '
    '      Dim strES As String
    '      Dim intEL As Integer
    '
    '1120  intEL = Erl
    '1130  strES = Err.Description
    '1140  LogError "frmEditMicrobiology", "cmdCopyFromPrevious_Click", intEL, strES, sql

    'Zyam 02-5-2024

End Sub

Private Sub cmdCopySensitivities_Click()
56140     Exit Sub
          'Abubaker+++ 06/09/2023 (Added Confirmation Propmt To Copy sensitivities and isolates from the previous record)
          '    Dim response As VbMsgBoxResult
          '
          '    response = MsgBox("Do you want to copy sensitivities and isolates from the previous record?", vbYesNo + vbQuestion, "Confirmation")
          '
          '    If response = vbYes Then
          '
          '      'copy sensitivities and isolates from previous record
          '      Dim sx As Sensitivity
          '
          '10    On Error GoTo cmdCopySensitivities_Click_Error
          '
          '20    GetSampleIDWithOffset
          '
          '30    CurrentSensitivities.Load SampleIDWithOffset - 1
          '40    For Each sx In CurrentSensitivities
          '50        sx.SampleID = SampleIDWithOffset
          '60    Next
          '70    CurrentSensitivities.Save
          '
          '80    LoadIsolates
          '90    LoadSensitivities
          '
          '100   cmdCopySensitivities.Visible = False
          '
          '     Else
          '        ' User clicked "No," do nothing or provide an alternative action
          '        ' You can add code here to handle the case when the user chooses not to proceed.
          '     End If
          'Abubaker--- 06/09/2023

56150     Exit Sub



cmdCopySensitivities_Click_Error:

          Dim strES As String
          Dim intEL As Integer

56160     intEL = Erl
56170     strES = Err.Description
56180     LogError "frmEditMicrobiology", "cmdCopySensitivities_Click", intEL, strES

End Sub

Private Sub cmdCopyTo_Click()

          Dim S As String

56190     S = cmbWard & " " & cmbClinician
56200     S = Trim$(S) & " " & cmbGP
56210     S = Trim$(S)

56220     frmCopyTo.lblOriginal = S
56230     frmCopyTo.lblSampleID = txtSampleID
56240     frmCopyTo.Dept = "M"
56250     frmCopyTo.Show 1

56260     CheckCC

End Sub

Private Sub CopyCC(ByVal strPrevID As String)

          Dim sql As String
          Dim tb As Recordset
          Dim sn As Recordset

56270     On Error GoTo CopyCC_Error

56280     If Trim$(txtSampleID) = "" Then Exit Sub

56290     sql = "Select * from SendCopyTo where " & _
              "SampleID = '" & Val(strPrevID) & "'"
56300     Set tb = New Recordset
56310     RecOpenServer 0, tb, sql
56320     If Not tb.EOF Then
              '70        sql = "Select * from SendCopyTo where " & _
              '                "SampleID = '" & sysOptMicroOffset(0) + Val(txtSampleID) & "'"
56330         sql = "Select * from SendCopyTo where " & _
                  "SampleID = '" & Val(txtSampleID) & "'"
56340         Set sn = New Recordset
56350         RecOpenServer 0, sn, sql
56360         If sn.EOF Then
56370             sn.AddNew
56380             sn!SampleID = Val(txtSampleID)
56390             sn!Ward = tb!Ward & ""
56400             sn!Clinician = tb!Clinician & ""
56410             sn!GP = tb!GP & ""
56420             sn!Device = tb!Device & ""
56430             sn!Destination = tb!Destination & ""
56440             sn.Update
56450         End If
56460     End If

56470     Exit Sub

CopyCC_Error:

          Dim strES As String
          Dim intEL As Integer

56480     intEL = Erl
56490     strES = Err.Description
56500     LogError "frmEditAll", "CopyCC", intEL, strES, sql
End Sub

Private Sub CheckCC()

          Dim sql As String
          Dim tb As Recordset

56510     On Error GoTo CheckCC_Error

56520     cmdCopyTo.Caption = "cc"
56530     cmdCopyTo.Font.Bold = False
56540     cmdCopyTo.BackColor = &H8000000F

56550     If Trim$(txtSampleID) = "" Then Exit Sub

56560     sql = "Select * from SendCopyTo where " & _
              "SampleID = '" & Val(txtSampleID) & "'"
56570     Set tb = New Recordset
56580     RecOpenServer 0, tb, sql
56590     If Not tb.EOF Then
56600         cmdCopyTo.Caption = "++ cc ++"
56610         cmdCopyTo.Font.Bold = True
56620         cmdCopyTo.BackColor = &H8080FF
56630     End If

56640     Exit Sub

CheckCC_Error:

          Dim strES As String
          Dim intEL As Integer

56650     intEL = Erl
56660     strES = Err.Description
56670     LogError "frmEditMicrobiology", "CheckCC", intEL, strES, sql

End Sub


Private Sub cmdGramPrep_Click()
56680     frmGramStains.SampleID = txtSampleID
56690     frmGramStains.Show 1
End Sub

Private Sub cmdHealthLink_Click()

          Dim SID As String

56700     SID = Format$(Val(txtSampleID))

56710     With cmdHealthLink
56720         If .Picture = imgHGreen.Picture Then
56730             Set .Picture = imgHRed.Picture
56740             ReleaseMicro SID, False
56750         Else
56760             Set .Picture = imgHGreen.Picture
56770             ReleaseMicro SID, True
56780         End If
56790     End With

End Sub

Private Sub cmdHealthLink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

56800     With cmdHealthLink
56810         If .Picture = imgHGreen.Picture Then
56820             .ToolTipText = "Released to HealthLink"
56830         Else
56840             .ToolTipText = "Click to Release to HealthLink"
56850         End If
56860     End With

End Sub


Private Sub cmdIsoArchive_Click()

56870     With frmManageIsolates
56880         .lblSampleID = txtSampleID
56890         .Show 1
56900     End With

End Sub

Private Sub cmdIsoRepeat_Click()

56910     With frmManageIsolates
56920         .lblSampleID = txtSampleID
56930         .Show 1
56940     End With

End Sub

Private Sub cmdLock_Click(Index As Integer)

56950     UpdateLockStatus Index

End Sub

Private Sub cmdMSC_Click()

56960     cmbMSC.Visible = True
56970     cmbMSC.SetFocus

End Sub

Private Sub cmdNADMicro_Click()

56980     txtBacteria = "Nil"
56990     txtWCC = "Nil"
57000     txtRCC = "Nil"
57010     cmbCrystals = "Nil"
57020     cmbCasts = "Nil"
57030     cmbMisc(0) = "Nil"

57040     ShowUnlock Mic.Urine

End Sub


Private Sub cmdPhone_Click()

57050     With frmPhoneLog
57060         .SampleID = Val(txtSampleID) ' + sysOptMicroOffset(0)
57070         If cmbGP <> "" Then
57080             .GP = cmbGP
57090             .WardOrGP = "GP"
57100         Else
57110             .GP = cmbWard
57120             .WardOrGP = "Ward"
57130         End If
57140         .Show 1
57150     End With

57160     CheckIfPhoned
57170     LoadMicroComments

End Sub

Private Sub CheckIfPhoned()

          Dim S As String
          Dim PhLog As PhoneLog
          Dim sql As String
          Dim OBs As Observations

57180     On Error GoTo CheckIfPhoned_Error

57190     PhLog = CheckPhoneLog(Val(txtSampleID))
57200     If PhLog.SampleID <> 0 Then
57210         cmdPhone.BackColor = vbYellow
57220         cmdPhone.Caption = "Results Phoned"
57230         cmdPhone.ToolTipText = "Results Phoned"
              '70        If InStr(txtDemographicComment.Text, "Results Phoned") = 0 Then
              '80            s = "Results Phoned to " & PhLog.PhonedTo & " at " & _
              '                  Format$(PhLog.DateTime, "hh:mm") & " on " & Format$(PhLog.DateTime, "dd/MM/yyyy") & _
              '                  " by " & PhLog.PhonedBy & "."
              '90            If Trim$(txtDemographicComment.Text) = "" Then
              '100               txtDemographicComment.Text = s
              '110           Else
              '120               txtDemographicComment.Text = txtDemographicComment.Text & ". " & s
              '130           End If
              '140           Set OBs = New Observations
              '150           OBs.Save PhLog.SampleID, True, "Demographic", txtDemographicComment.Text
              '
              '160       End If
57240     Else
57250         cmdPhone.BackColor = &H8000000F
57260         cmdPhone.Caption = "Phone Results"
57270         cmdPhone.ToolTipText = "Phone Results"
57280     End If

57290     Exit Sub

CheckIfPhoned_Error:

          Dim strES As String
          Dim intEL As Integer

57300     intEL = Erl
57310     strES = Err.Description
57320     LogError "frmEditMicrobiology", "CheckIfPhoned", intEL, strES, sql

End Sub
Private Sub LoadMicroComments()

          Dim OB As Observation
          Dim OBs As Observations



57330     On Error GoTo LoadMicroComments_Error

57340     txtDemographicComment = ""

57350     If Val(txtSampleID) = 0 Then Exit Sub

57360     Set OBs = New Observations
57370     Set OBs = OBs.Load(Val(txtSampleID), "Demographic")
57380     If Not OBs Is Nothing Then
57390         For Each OB In OBs
57400             Select Case UCase$(OB.Discipline)
                      Case UCase("Demographic"): txtDemographicComment = Split_Comm(OB.Comment)
57410             End Select
57420         Next
57430     End If



57440     Exit Sub

LoadMicroComments_Error:
          Dim strES As String
          Dim intEL As Integer

57450     intEL = Erl
57460     strES = Err.Description
57470     LogError "frmEditMicrobiology", "LoadMicroComments", intEL, strES

End Sub
Private Sub cmdRemoveSecondary_Click(Index As Integer)

          Dim y As Integer
          Dim sx As Sensitivity

57480     On Error GoTo cmdRemoveSecondary_Click_Error

57490     With grdAB(Index)
57500         .Col = 0
57510         For y = 1 To .Rows - 1
57520             .row = y
57530             If .CellFontBold = True Then
57540                 Set sx = CurrentSensitivities.Item(CStr(Index), .TextMatrix(y, 6))
57550                 sx.Forced = 0
57560                 sx.Secondary = 0
57570                 sx.Report = 0
                      '80          CurrentSensitivities.Delete Sx
57580             End If
57590         Next
57600     End With

57610     LoadSensitivities

57620     Exit Sub

cmdRemoveSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

57630     intEL = Erl
57640     strES = Err.Description
57650     LogError "frmEditMicrobiology", "cmdRemoveSecondary_Click", intEL, strES

End Sub

Private Sub cmdSaveHold_Click()
57660     bPrint.Enabled = False
57670     cmdIntrim.Enabled = False
57680     UpDown1.Enabled = False
57690     If UserHasAuthority(UserMemberOf, "MicroOtherSave") = False Then
57700         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
57710         bPrint.Enabled = True
57720         cmdIntrim.Enabled = True
57730         UpDown1.Enabled = True
57740         Exit Sub
57750     End If

57760     SaveMicro False, 0
57770     MsgBox "Record has been saved.", vbInformation
57780     LoadAllDetails
57790     bPrint.Enabled = True
57800     cmdIntrim.Enabled = True
57810     UpDown1.Enabled = True

End Sub

Private Sub cmdSensArchive_Click()

57820     With frmManageSensitivities
57830         .lblSampleID = txtSampleID
57840         .Show 1
57850     End With

End Sub

Private Sub cmdSensRepeat_Click()

57860     With frmManageSensitivities
57870         .lblSampleID = txtSampleID
57880         .Show 1
57890     End With

End Sub

Private Sub cmdSetValid_Click()

57900     With frmSetValid
57910         .lblSampleID = txtSampleID
57920         .Show 1
57930     End With

End Sub

Private Sub cmdTagRepeat_Click()

          Dim Us As New UrineResults
          Dim U As New UrineResult

57940     If cmdTagRepeat.Caption = "Tag as Repeat" Then
57950         cmdTagRepeat.Caption = "Un-Tag Repeat"
57960     Else
57970         cmdTagRepeat.Caption = "Tag as Repeat"
57980     End If
          '+++ Junaid 20-05-2024
          '60    U.SampleID = SampleIDWithOffset
57990     U.SampleID = Trim(txtSampleID.Text)
          '--- Junaid
58000     U.UserName = UserName
58010     U.Valid = 0
58020     U.ValidatedBy = ""
58030     U.ValidatedDateTime = ""

58040     U.TestName = "TagRepeat"
58050     U.Result = cmdTagRepeat.Caption
58060     Us.Save U

58070     If chkPregnant.Value = 0 Then CheckUrineAutoVal

End Sub

Private Sub cmdUnLock_Click()

58080     If UCase(iBOX("Enter password to unValidate ?", , , True)) = UCase$(TechnicianPassFor(UserName)) Then

58090         ComingFromUnlock = True
58100         LockDemographics Me, False
58110         txtSurName.SetFocus

58120     End If

End Sub

Private Sub cmdUseSecondary_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim ABName As String
          Dim RunDateTime As String
          Dim sx As Sensitivity

58130     On Error GoTo cmdUseSecondary_Click_Error

58140     sql = "Select Distinct LTRIM(RTRIM(A.AntibioticName)) ABName, A.Code, D.ListOrder, " & _
              "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
              "from ABDefinitions as D, Antibiotics as A where " & _
              "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
              "and D.Site = '" & cmbSite & "' " & _
              "and D.PriSec = 'S' " & _
              "and D.AntibioticName = A.AntibioticName " & _
              "order by D.ListOrder"
58150     Set tb = New Recordset
58160     RecOpenServer 0, tb, sql
58170     If tb.EOF Then
58180         sql = "Select Distinct LTRIM(RTRIM(A.AntibioticName)) ABName, A.Code, D.ListOrder, " & _
                  "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                  "from ABDefinitions as D, Antibiotics as A where " & _
                  "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                  "and (D.Site = 'Generic' or D.Site is Null ) and D.PriSec = 'S' " & _
                  "and D.AntibioticName = A.AntibioticName " & _
                  "order by D.ListOrder"
58190         Set tb = New Recordset
58200         RecOpenServer 0, tb, sql
58210         If tb.EOF Then
58220             Exit Sub
58230         End If
58240     End If
58250     Do While Not tb.EOF
58260         Found = False
58270         For n = 1 To grdAB(Index).Rows - 1
58280             If Trim$(grdAB(Index).TextMatrix(n, 0)) = tb!ABName Then
58290                 Found = True
58300                 Exit For
58310             End If
58320         Next

58330         If Not Found Then
58340             Set sx = CurrentSensitivities.Item(Index, tb!Code)
58350             If Not sx Is Nothing Then
58360                 sx.Secondary = True
58370                 sx.Forced = True
58380                 sx.UserCode = UserCode
58390             Else


58400                 Set sx = New Sensitivity
58410                 sx.AntibioticName = tb!ABName
58420                 sx.AntibioticCode = tb!Code
58430                 sx.Forced = True
58440                 sx.Secondary = True
58450                 sx.IsolateNumber = Index
58460                 sx.Rundate = Format$(Now, "dd/MMM/yyyy")
58470                 sx.RunDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
                      '+++ Junaid 20-05-2024
                      '360               sx.SampleID = SampleIDWithOffset
58480                 sx.SampleID = Trim(txtSampleID.Text)
                      '--- Junaid
58490                 sx.UserCode = UserCode
58500                 sx.Valid = False
58510                 If IsChild() And Not tb!AllowIfChild <> 0 Then
58520                     sx.CPOFlag = "C"
58530                 ElseIf IsPregnant() And Not tb!AllowIfPregnant <> 0 Then
58540                     sx.CPOFlag = "P"
58550                 ElseIf IsOutPatient() And Not tb!AllowIfOutPatient <> 0 Then
58560                     sx.CPOFlag = "O"
58570                 End If
58580                 If UCase(cmdValidateMicro.Caption) = UCase("Validate") Or UCase(cmdValidateMicro.Caption) = UCase("&Validate") Then sx.Save
58590                 CurrentSensitivities.Add sx
58600             End If
58610         End If
58620         tb.MoveNext
58630     Loop
58640     LoadSensitivities

58650     Exit Sub

cmdUseSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

58660     intEL = Erl
58670     strES = Err.Description
58680     LogError "frmEditMicrobiology", "cmdUseSecondary_Click", intEL, strES, sql

End Sub

Private Sub cmdViewReports_Click()


          Dim f As Form

58690     Set f = New frmReportViewer

58700     f.Dept = "Microbiology"
58710     f.SampleID = txtSampleID
58720     f.Show 1

58730     Set f = Nothing

End Sub

Private Sub SetViewReports(ByVal SampleID As String)

          Dim sql As String
          Dim tb As New Recordset

58740     On Error GoTo SetViewReports_Error

58750     cmdViewReports.Visible = False

58760     If SampleID <> "" Then
58770         sql = "SELECT COUNT(*) Tot FROM Reports " & _
                  "WHERE SampleID = '" & SampleID & "' " & _
                  "AND Dept = 'Microbiology'"
              '+++ Junaid 12-01-2024
              '40    Set tb = Cnxn(0).Execute(Sql)
58780         Set tb = New Recordset
58790         RecOpenServer 0, tb, sql
              '--- Junaid
58800         cmdViewReports.Visible = tb!Tot > 0
58810     Else
58820         cmdViewReports.Visible = False
58830     End If
58840     Exit Sub

SetViewReports_Error:

          Dim strES As String
          Dim intEL As Integer

58850     intEL = Erl
58860     strES = Err.Description
58870     LogError "frmEditMicrobiology", "SetViewReports", intEL, strES, sql

End Sub


Private Sub cMRU_GotFocus()

58880     If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
58890         If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
                  '30            If Not GetSampleIDWithOffset Then Exit Sub
58900             SaveDemographics
58910             cmdSaveDemographics.Enabled = False
58920             cmdSaveInc.Enabled = False
58930         End If
58940     End If

End Sub

Private Sub cmdAddToConsultantList_Click()

          Dim sql As String
          Dim tb As Recordset

58950     On Error GoTo cmdAddToConsultantList_Click_Error

58960     Select Case Left$(cmdAddToConsultantList.Caption, 3)
              Case "Add":
58970             sql = "Select * from ConsultantList where " & _
                      "SampleID = " & Val(txtSampleID)
58980             Set tb = New Recordset
58990             RecOpenServer 0, tb, sql
59000             If tb.EOF Then tb.AddNew
59010             tb!SampleID = Val(txtSampleID)
59020             tb.Update
59030             Call ConsultantListLog(Val(txtSampleID), "Added to consultant list")
59040         Case "Rem":
59050             sql = "Delete from ConsultantList " & _
                      "where SampleID = '" & Val(txtSampleID) & "'"
59060             Cnxn(0).Execute sql
59070             Call ConsultantListLog(Val(txtSampleID), "Removed from consultant list")
59080     End Select

59090     FillForConsultantValidation

59100     Exit Sub

cmdAddToConsultantList_Click_Error:

          Dim strES As String
          Dim intEL As Integer

59110     intEL = Erl
59120     strES = Err.Description
59130     LogError "frmEditMicrobiology", "cmdAddToConsultantList_Click", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Consultant
' Author    : Masood
' Date      : 28/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : ConsultantListLog
' Author    : Masood
' Date      : 28/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ConsultantListLog(ByVal SampleID As String, ByVal Status As String)
59140     On Error GoTo ConsultantListLog_Error


          Dim sql As String
          Dim tb As New ADODB.Recordset
59150     sql = "Select * from ConsultantListLog where 1 = 2 "
59160     Set tb = New Recordset
59170     RecOpenServer 0, tb, sql
59180     If tb.EOF Then tb.AddNew
59190     tb!SampleID = SampleID
59200     tb!UserName = UserName
59210     tb!Status = Status
59220     tb.Update

       
59230     Exit Sub

       
ConsultantListLog_Error:

          Dim strES As String
          Dim intEL As Integer

59240     intEL = Erl
59250     strES = Err.Description
59260     LogError "frmEditMicrobiology", "Consultant", intEL, strES, sql

       
End Sub


Private Sub cmdPatientNotePad_Click()
59270     On Error GoTo cmdPatientNotePad_Click_Error

59280     frmPatientNotePad.SampleID = txtSampleID
59290     frmPatientNotePad.Caller = "Microbiology"
59300     frmPatientNotePad.Show 1

59310     Exit Sub

cmdPatientNotePad_Click_Error:

          Dim strES As String
          Dim intEL As Integer

59320     intEL = Erl
59330     strES = Err.Description
59340     LogError "frmEditMicrobiology", "cmdPatientNotePad_Click", intEL, strES
          

End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtRecDate_CloseUp()

59350     pBar = 0

59360     cmdSaveDemographics.Enabled = True
59370     cmdSaveInc.Enabled = True

End Sub


Private Sub dtRecDate_LostFocus()
59380     SetDatesColour
End Sub

Private Sub dtRunDate_LostFocus()
59390     SetDatesColour
End Sub

Private Sub dtSampleDate_LostFocus()
59400     SetDatesColour
End Sub

Private Sub fraMicroResult_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
59410     pBar = 0
End Sub

Private Sub grdAB_Click(Index As Integer)

          Dim S As String
          Dim sx As New Sensitivity

59420     On Error GoTo grdAB_Click_Error

59430     If Not fraMicroResult(Mic.CandS).Enabled Then Exit Sub

59440     cmdSaveMicro.Enabled = True
59450     cmdSaveHold.Enabled = True

59460     With grdAB(Index)
59470         If .MouseRow = 0 Then Exit Sub

59480         If .CellBackColor = &HFFFFC0 Then
59490             .Enabled = False
59500             If iMsg("Remove " & Trim$(.Text) & " from List?", vbQuestion + vbYesNo) = vbYes Then
59510                 Set sx = CurrentSensitivities.Item(CStr(Index), .TextMatrix(.row, 6))
59520                 CurrentSensitivities.ForceUnForce sx, 0
59530                 CurrentSensitivities.SetSecondary sx, 0
59540                 Set CurrentSensitivities = New Sensitivities
                      '+++ Junaid 20-05-2024
                      '140               CurrentSensitivities.Load SampleIDWithOffset
59550                 CurrentSensitivities.Load Trim(txtSampleID.Text)
                      '--- Junaid
59560                 LoadSensitivities
59570             End If
59580             .Enabled = True
59590         ElseIf .Col = 1 Then
59600             S = Trim$(.TextMatrix(.row, 1))
59610             Select Case S
                      Case "": S = "R"
59620                 Case "R": S = "S"
59630                 Case "S": S = "I"
59640                 Case "I": S = ""
59650                 Case Else: S = ""
59660             End Select
59670             .TextMatrix(.row, 1) = S
59680         ElseIf .Col = 2 Then
59690             .TextMatrix(.row, 5) = UserCode
59700             If .CellPicture = imgSquareTick.Picture Then
59710                 Set .CellPicture = imgSquareCross.Picture
59720             Else
59730                 If .TextMatrix(.row, 2) = "C" Then
59740                     If MsgBox("Report " & .TextMatrix(.row, 0) & " on a Child?", vbQuestion + vbYesNo) = vbNo Then
59750                         Exit Sub
59760                     End If
59770                 ElseIf .TextMatrix(.row, 2) = "P" Then
59780                     If MsgBox("Report " & .TextMatrix(.row, 0) & " for Pregnant Patient?", vbQuestion + vbYesNo) = vbNo Then
59790                         Exit Sub
59800                     End If
59810                 ElseIf .TextMatrix(.row, 2) = "O" Then
59820                     If MsgBox("Report " & .TextMatrix(.row, 0) & " for an Out-Patient?", vbQuestion + vbYesNo) = vbNo Then
59830                         Exit Sub
59840                     End If
59850                 End If
59860                 Set .CellPicture = imgSquareTick.Picture
59870             End If
59880             .LeftCol = 0
59890         ElseIf .Col = 3 Then
59900             .Enabled = False
59910             S = iBOX("Enter Result", , .TextMatrix(.row, 3))
59920             .TextMatrix(.row, 3) = S
59930             .TextMatrix(.row, 4) = Format$(Now, "dd/MM/yy HH:mm")
59940             .TextMatrix(.row, 5) = TechnicianCodeFor(UserName)
59950             .Enabled = True
59960         End If
59970     End With

59980     Exit Sub

grdAB_Click_Error:

          Dim strES As String
          Dim intEL As Integer

59990     intEL = Erl
60000     strES = Err.Description
60010     LogError "frmEditMicrobiology", "grdAB_Click", intEL, strES

End Sub

Private Sub grdAB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
60020     pBar = 0
End Sub

Private Sub iRecDate_Click(Index As Integer)

60030     If Index = 0 Then
60040         dtRecDate = DateAdd("d", -1, dtRecDate)
60050     Else
60060         If DateDiff("d", dtRecDate, Now) > 0 Then
60070             dtRecDate = DateAdd("d", 1, dtRecDate)
60080         End If
60090     End If

60100     cmdSaveInc.Enabled = True
60110     cmdSaveDemographics.Enabled = True

End Sub

Private Sub iRelevant_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

60120     If cmdSaveMicro.Enabled Then
60130         MoveCursorToSaveButton
60140     End If

End Sub

Private Sub lblcDiffPCR_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

60150     With lblcDiffPCR
60160         Select Case .Caption
                  Case ""
60170                 .Caption = "Toxigenic C.diff: NEGATIVE"
60180                 .BackColor = vbGreen
60190             Case "Toxigenic C.diff: NEGATIVE"
60200                 .Caption = "Toxigenic C.diff: POSITIVE"
60210                 .BackColor = vbRed
60220             Case "Toxigenic C.diff: POSITIVE"
60230                 .Caption = ""
60240                 .BackColor = &H8000000F
60250         End Select
60260     End With

60270     ShowUnlock Mic.CDiff

End Sub


Private Sub lblCrypto_Click()

60280     With lblCrypto
60290         Select Case .Caption
                  Case ""
60300                 .Caption = "Negative"
60310                 .BackColor = vbGreen
60320             Case "Negative"
60330                 .Caption = "Positive"
60340                 .BackColor = vbRed
60350             Case "Positive"
60360                 .Caption = ""
60370                 .BackColor = &H8000000F
60380         End Select
60390     End With

60400     ShowUnlock Mic.OP

End Sub

Private Sub lblFOB_Click(Index As Integer)

60410     With lblFOB(Index)
60420         Select Case .Caption
                  Case ""
60430                 .Caption = "Negative"
60440                 .BackColor = vbGreen
60450             Case "Negative"
60460                 .Caption = "Positive"
60470                 .BackColor = vbRed
60480             Case "Positive"
60490                 .Caption = ""
60500                 .BackColor = &H8000000F
60510         End Select
60520     End With

60530     ShowUnlock Mic.FOB

End Sub

Private Sub lblRSV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

60540     With lblRSV
60550         Select Case .Caption
                  Case ""
60560                 .Caption = "Negative"
60570                 .BackColor = vbGreen
60580             Case "Negative"
60590                 .Caption = "Positive"
60600                 .BackColor = vbRed
60610             Case "Positive"
60620                 .Caption = "Inconclusive"
60630                 .BackColor = vbYellow
60640             Case "Inconclusive"
60650                 .Caption = ""
60660                 .BackColor = &H8000000F
60670         End Select
60680     End With

60690     ShowUnlock Mic.RSV

End Sub


Private Sub lblSetAllR_Click(Index As Integer)

          Dim y As Integer

60700     With grdAB(Index)
60710         For y = 1 To .Rows - 1
60720             If .TextMatrix(y, 0) <> "" Then
60730                 .TextMatrix(y, 1) = "R"
60740             End If
60750         Next
60760     End With

60770     cmdSaveMicro.Enabled = True
60780     cmdSaveHold.Enabled = True

End Sub

Private Sub lblSetAllS_Click(Index As Integer)

          Dim y As Integer

60790     With grdAB(Index)
60800         For y = 1 To .Rows - 1
60810             If .TextMatrix(y, 0) <> "" Then
60820                 .TextMatrix(y, 1) = "S"
60830             End If
60840         Next
60850     End With

60860     cmdSaveMicro.Enabled = True
60870     cmdSaveHold.Enabled = True

End Sub


Private Sub lblToxinA_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

60880     With lblToxinA
60890         Select Case .Caption
                  Case ""
60900                 .Caption = "Not Detected"
60910                 .BackColor = vbGreen
60920             Case "Not Detected"
60930                 .Caption = "Positive"
60940                 .BackColor = vbRed
60950             Case "Positive"
60960                 .Caption = "Inconclusive"
60970                 .BackColor = vbYellow
60980             Case "Inconclusive"
60990                 .Caption = "Rejected"
61000                 .BackColor = &H8000000F
61010             Case "Rejected"
61020                 .Caption = ""
61030                 .BackColor = &H8000000F
61040         End Select
61050         lblToxinB.Caption = .Caption
61060         lblToxinB.BackColor = .BackColor
61070     End With

61080     ShowUnlock Mic.CDiff

End Sub


Private Sub bDoB_Click()

61090     pBar = 0

61100     LogEvent "DoB Search Click", "frmEditAll", "bDoB_Click"

61110     With frmPatHistory
61120         If SSTab1.Tab = 0 Then
61130             .chkShort.Value = 1
61140         Else
61150             .chkShort.Value = 0
61160         End If
61170         If HospName(0) = "Monaghan" And SSTab1.Tab = 0 Then
61180             .optBoth = True
61190         Else
61200             .optHistoric = True
61210         End If
61220         .lblDept = "M"
61230         .oFor(2) = True
61240         .txtName = txtDoB
61250         .FromEdit = True
61260         .EditScreen = Me
61270         .bsearch = True
61280         If Not .NoPreviousDetails Then
61290             .Show 1
61300             If txtLabNo = "" Then
61310                 If FormLoaded Then txtLabNo = ""      'Val(FndMaxID("demographics", "LabNo", ""))
61320             End If
61330         Else
61340             FlashNoPrevious Me
61350         End If
61360     End With
61370     CheckCC

End Sub

Private Sub bFAX_Click(Index As Integer)

61380     If UserHasAuthority(UserMemberOf, "MicroOtherFax") = False Then
61390         iMsg "You do not have authority to fax" & vbCrLf & "Please contact system administrator"
61400         Exit Sub
61410     End If

61420     pBar = 0

End Sub

Private Sub bHistory_Click()

          Dim f As Form
          Dim SID As String

61430     pBar = 0

61440     If pFromViewReportSID <> "" Then
61450         Me.Hide
61460     Else
61470         If Trim$(txtSurName) = "" Then
61480             Exit Sub
61490         End If
61500         Set f = New frmMicroReport
61510         With f
61520             .lblChart = txtChart
61530             .lblName = Trim$(txtSurName & " " & txtForeName)
61540             .lblDoB = txtDoB
61550             .lblSex = Trim$(Left$(txtSex & " ", 1))
61560             .Show 1
61570             If .ReturnedSampleID <> "" Then
61580                 SID = .ReturnedSampleID
61590             Else
61600                 SID = txtSampleID
61610             End If
61620         End With
61630         Unload f
61640         Set f = Nothing

61650         txtSampleID = SID
61660         LoadAllDetails
61670     End If

End Sub



Private Sub cmbSite_Click()

61680     If cmbSite = "Blood Culture" Then
61690         SelectBloodCulture
61700     ElseIf cmbSite = "MRSA Screen" Or cmbSite = "VRE Screen" Then
              '40        If Not GetSampleIDWithOffset Then Exit Sub
61710         SaveDemographics
61720         With frmAssociate
61730             .lblSID = txtSampleID
61740             .lblChart = txtChart
61750             .lblName = Trim$(txtSurName & " " & txtForeName)
61760             .Show 1
61770         End With
61780     Else
61790         cmbSiteEffects
61800     End If

61810     EnableTagRepeat

End Sub

Private Sub cmdNAD_Click()

61820     If txtProtein = "" Then txtProtein = "Nil"
61830     If txtGlucose = "" Then txtGlucose = "Nil"
61840     If txtKetones = "" Then txtKetones = "Nil"
61850     If txtWCC = "" Then txtWCC = "Nil"
61860     If txtRCC = "" Then txtRCC = "Nil"
61870     If cmbCasts = "" Then cmbCasts = "Nil"
61880     If cmbCrystals = "" Then cmbCrystals = "Nil"
61890     If txtBilirubin = "" Then txtBilirubin = "Nil"
61900     If txtUrobilinogen = "" Then txtUrobilinogen = "Nil"
61910     If txtBloodHb = "" Then txtBloodHb = "Nil"

61920     ShowUnlock Mic.Urine

End Sub

Private Sub cmdOrderTests_Click()

61930     pBar = 0

61940     If UserHasAuthority(UserMemberOf, "MicroOrderTest") = False Then
61950         iMsg "You do not have authority to Order test in microbiology" & vbCrLf & "Please contact system administrator"
61960         Exit Sub
61970     End If

61980     If cmbSite = "Urine" Then
61990         OrderUrine
62000         EnableTagRepeat
62010     ElseIf cmbSite = "Faeces" Then
62020         OrderFaeces
62030     Else
62040         OrderFaeces
62050     End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdIntrim_Click
' Author    : Masood
' Date      : 12/Jan/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdIntrim_Click()

          Dim tb As Recordset
          Dim sql As String
          'Dim Validate As Integer
          Dim x As Integer
          Dim y As Integer
          Dim SID As String



62060     On Error GoTo cmdIntrim_Click_Error
62070     UpDown1.Enabled = False
62080     DoEvents

          '20    If Not GetSampleIDWithOffset Then Exit Sub
          '+++ Junaid 20-05-2024
          '30    SID = Format$(Val(txtSampleID) + sysOptMicroOffset(0))
62090     SID = Format$(Val(txtSampleID))
          '--- Junaid


62100     SaveMicro False, False
          '50    SaveDemographics

62110     PrintThis ("Save")

62120     With cmdHealthLink
62130         Set .Picture = imgHGreen.Picture
62140         ReleaseMicro SID, True
62150     End With


          'Select Case SSTab1.Tab
          '    Case Mic.Urine: SaveUrine Validate
          '    Case Mic.Identification:
          '    Case Mic.Faeces: SaveFaeces: FillHistoricalFaeces
          '    Case Mic.CandS: SaveIsolates: SaveSensitivities Validate
          '    Case Mic.FOB: SaveFOB Validate
          '    Case Mic.RotaAdeno: SaveRotaAdeno Validate
          '    Case Mic.CDiff: SaveCdiff Validate
          '    Case Mic.OP: SaveOP Validate
          '    Case Mic.IDENTCAVAN: SaveIdentification
          '    Case Mic.RSV: SaveRSV Validate
          '    Case Mic.CSF: SaveCSF
          '    Case Mic.RedSub: SaveRedSub Validate
          'End Select
          '
          '
          'SaveComments
          'UpdateMRU Me

          'Zyam commented this so the it doesn't go the next sampleid when the interim report is printed 21-05-24
          'txtSampleID = Format$(Val(txtSampleID) + 1)
          'Zyam 21-05-24
          '120   If Not GetSampleIDWithOffset Then Exit Sub
62160     LoadAllDetails
62170     UpDown1.Enabled = True
62180     Exit Sub


cmdIntrim_Click_Error:

          Dim strES As String
          Dim intEL As Integer
62190     UpDown1.Enabled = True
62200     intEL = Erl
62210     strES = Err.Description
62220     LogError "frmEditMicrobiology", "cmdIntrim_Click", intEL, strES, sql

End Sub

Private Sub bPrint_Click()

62230     On Error GoTo bPrint_Click_Error
          Dim SID As String
62240     UpDown1.Enabled = False
62250     DoEvents
62260     If UserHasAuthority(UserMemberOf, "MicroOtherPrint") = False Then
62270         iMsg "You do not have authority to print" & vbCrLf & "Please contact system administrator"
62280         UpDown1.Enabled = True
62290         Exit Sub
62300     End If


62310     SaveMicro False, True
          '70    SaveDemographics

62320     SID = Format$(Val(txtSampleID))

62330     With cmdHealthLink
62340         Set .Picture = imgHGreen.Picture
62350         ReleaseMicro SID, True
62360     End With

62370     If SSTab1.TabVisible(1) Then
62380         SaveUrine gVALID
62390     End If
62400     DoEvents
          'Select Case SSTab1.Tab
          '    Case Mic.Urine:
          '    Case Mic.Identification:
          '    Case Mic.Faeces: SaveFaeces: FillHistoricalFaeces
          '    Case Mic.CandS: SaveIsolates: SaveSensitivities gYES
          '    Case Mic.FOB: SaveFOB gVALID
          '    Case Mic.RotaAdeno: SaveRotaAdeno gVALID
          '    Case Mic.CDiff: SaveCdiff gVALID
          '    Case Mic.OP: SaveOP gVALID
          '    Case Mic.IDENTCAVAN: SaveIdentification
          '    Case Mic.RSV: SaveRSV gVALID
          '    Case Mic.CSF: SaveCSF
          '    Case Mic.RedSub: SaveRedSub gVALID
          'End Select

62410     PrintThis ("")
          '+++ Junaid 20-05-2024
          '170   UpdatePrintValidLog SampleIDWithOffset, "MICRO", 1, 0
62420     UpdatePrintValidLog Trim(txtSampleID.Text), "MICRO", 1, 0
          '--- Junaid

          'Criteria for Auto SignOff of Micro Results
          'When Criteria is met the Micro Results are set to be Sign Off
          'This can be seen in Ward Enquiry
          '+++ Junaid 20-05-2024
          '175   UpdatePrintValidLog_AutoSignOff SampleIDWithOffset, "MICRO"
62430     UpdatePrintValidLog_AutoSignOff Trim(txtSampleID.Text), "MICRO"
          '--- Junaid
          
          '180   txtSampleID = Format$(Val(txtSampleID) + 1)
          '190   If Not GetSampleIDWithOffset Then Exit Sub
62440     LoadAllDetails
62450     UpDown1.Enabled = True
62460     Exit Sub

bPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer
62470     UpDown1.Enabled = True
62480     intEL = Erl
62490     strES = Err.Description
62500     LogError "frmEditMicrobiology", "bPrint_Click", intEL, strES


End Sub

Private Sub SaveDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim tbL As Recordset
          Dim Hosp As String
          Dim NewLabNumber As String

62510     On Error GoTo SaveDemographics_Error

62520     txtSampleID = Format(Val(txtSampleID))
62530     If Val(txtSampleID) = 0 Then Exit Sub
          '+++ Junaid 21-05-2024
62540     Call SaveAuditDemo(Trim(txtSampleID.Text))
          '--- Junaid
62550     NewLabNumber = DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
          'DoEvents
          'DoEvents
          '50    SaveComments

62560     If Trim$(tSampleTime) <> "__:__" Then
62570         If Not IsDate(tSampleTime) Then
62580             iMsg "Invalid Sample Time", vbExclamation
62590             Exit Sub
62600         End If
62610     End If

62620     If Not IsDate(tRecTime) Then
62630         iMsg "Invalid Received Time", vbExclamation
62640         Exit Sub
62650     End If

62660     If InStr(UCase$(lblChartNumber), "CAVAN") Then
62670         Hosp = "Cavan"
62680     ElseIf InStr(UCase$(lblChartNumber), "MONAGHAN") Then
62690         Hosp = "Monaghan"
62700     Else
62710         Hosp = ""
62720     End If

62730     If Trim$(cmbSite) = "" Then
62740         iMsg "Invalid site", vbExclamation
62750         Exit Sub
62760     End If

          Dim SDS As New SiteDetails
          Dim SD As New SiteDetail
          '+++ Junaid 20-05-2024
          '270   SD.SampleID = SampleIDWithOffset
62770     SD.SampleID = Trim(txtSampleID.Text)
          '--- Junaid
62780     SD.Site = cmbSite
62790     SD.SiteDetails = Trim(txtSiteDetails.Text)
62800     SD.UserName = UserName
62810     SDS.Save SD
          '+++ Junaid 22-12-2023
62820     sql = "Select * from CavanLog Where SampleID = '" & txtSampleID.Text & "'"
62830     Set tbL = New Recordset
62840     RecOpenClient 0, tbL, sql
62850     If tbL.EOF Then
62860         tbL.AddNew
62870         tbL!SampleID = txtSampleID.Text
62880         tbL!AddDate = Format$(Now, "dd/mmm/yyyy")
62890         tbL!AddUser = UserName
62900         tbL.Update
62910     End If
          '--- Junaid
62920     SaveCurrentAntibiotics
          '+++ Junaid 20-05-2024
          '330   sql = "Select * from Demographics where " & _
          '            "SampleID = '" & SampleIDWithOffset & "'"
62930     sql = "Select * from Demographics where " & _
              "SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
62940     Set tb = New Recordset
62950     RecOpenClient 0, tb, sql
62960     If tb.EOF Then
62970         tb.AddNew
62980         tb!Fasting = 0
62990         tb!FAXed = 0
63000     Else

63010         If Trim$(tb!PatName & "") <> "" And _
                  Trim$(UCase$(tb!PatName & "")) <> Trim$(UCase$(txtSurName & " " & txtForeName)) Then
                  '+++ Junaid 20-05-2024
                  '420           If FlagMessage("Name", tb!PatName, txtSurName & " " & txtForeName, SampleIDWithOffset) Then
63020             If FlagMessage("Name", tb!PatName, txtSurName & " " & txtForeName, Trim(txtSampleID.Text)) Then
                      '--- Junaid
63030                 txtSurName = SurName(tb!PatName & "")
63040                 txtForeName = ForeName(tb!PatName & "")
63050             End If
63060         End If
63070         If Not IsNull(tb!DoB) Then
63080             If Format(tb!DoB, "dd/mm/yyyy") <> Format(txtDoB, "dd/mm/yyyy") Then
                      '+++ Junaid 20-05-2024
                      '490               If FlagMessage("DoB", tb!DoB, txtDoB, SampleIDWithOffset) Then
63090                 If FlagMessage("DoB", tb!DoB, txtDoB, Trim(txtSampleID.Text)) Then
                          '--- Junaid
63100                     txtDoB = Format(tb!DoB, "dd/mm/yyyy")
63110                 End If
63120             End If
63130         End If
63140         If Trim$(tb!Chart & "") <> "" And Trim$(UCase$(tb!Chart & "")) <> Trim$(UCase$(txtChart)) Then
                  '+++ Junaid 20-05-2024
                  '550           If FlagMessage("Chart", tb!Chart, txtChart, SampleIDWithOffset) Then
63150             If FlagMessage("Chart", tb!Chart, txtChart, Trim(txtSampleID.Text)) Then
                      '--- Junaid
63160                 txtChart = tb!Chart & ""
63170             End If
63180         End If
63190         If Trim$(tb!Ward & "") <> "" And Trim$(UCase$(tb!Ward & "")) <> Trim$(UCase$(cmbWard)) Then
                  '+++ Junaid 20-05-2024
                  '600           If FlagMessage("Ward", tb!Ward, cmbWard, SampleIDWithOffset) Then
63200             If FlagMessage("Ward", tb!Ward, cmbWard, Trim(txtSampleID.Text)) Then
                      '--- Junaid
63210                 cmbWard = tb!Ward & ""
63220             End If
63230         End If
63240         If Trim$(tb!Clinician & "") <> "" And Trim$(UCase$(tb!Clinician & "")) <> Trim$(UCase$(cmbClinician)) Then
                  '+++ Junaid 20-05-2024
                  '650           If FlagMessage("Clinician", tb!Clinician, cmbClinician, SampleIDWithOffset) Then
63250             If FlagMessage("Clinician", tb!Clinician, cmbClinician, Trim(txtSampleID.Text)) Then
                      '--- Junaid
63260                 cmbClinician = tb!Clinician & ""
63270             End If
63280         End If

63290     End If

63300     tb!RooH = cRooH(0)
63310     tb!LabNo = NewLabNumber
63320     If IsDate(tRecTime) Then
63330         tb!RecDate = Format$(dtRecDate & " " & tRecTime, "dd/mmm/yyyy hh:mm")
63340     Else
63350         tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy")
63360     End If
63370     tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
63380     If IsDate(tSampleTime) Then
63390         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
63400     Else
63410         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
63420     End If
          '+++ Junaid 20-05-2024
          '830   tb!SampleID = SampleIDWithOffset
63430     tb!SampleID = Trim(txtSampleID.Text)
          '--- Junaid
63440     tb!Chart = txtChart
63450     tb!PatName = Trim$(txtSurName & " " & txtForeName)
63460     tb!SurName = txtSurName & ""
63470     tb!ForeName = txtForeName & ""
63480     If IsDate(txtDoB) Then
63490         tb!DoB = Format$(txtDoB, "dd/mmm/yyyy")
63500     Else
63510         tb!DoB = Null
63520     End If
63530     tb!Age = txtAge.Text
63540     tb!Sex = Left$(txtSex, 1)
63550     tb!Addr0 = txtAddress(0)
63560     tb!Addr1 = txtAddress(1)
63570     tb!Ward = Left$(cmbWard, 50)
63580     tb!Clinician = Left$(cmbClinician, 50)
63590     tb!GP = Left$(cmbGP, 50)
63600     tb!ClDetails = Left$(txtClinDetails, 500)
63610     tb!Hospital = Hosp
63620     tb!Pregnant = chkPregnant
63630     tb!Operator = Left$(UserName, 20)
63640     tb!RecordDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
63650     tb!ExtSampleID = txtExtSampleID
63660     tb.Update
          '    '+++Junaid 21-12-2023
          '    txtIdentification(1).Text = txtSampleID.Text
          '    '---Junaid
63670     SaveComments

63680     LabNoUpdatePrvData txtChart, Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, Left$(txtSex, 1), txtLabNo
          '+++ Junaid 20-05-2024
          '1080  LogTimeOfPrinting SampleIDWithOffset, "D"
63690     LogTimeOfPrinting Trim(txtSampleID.Text), "D"
          '--- Junaid
63700     Screen.MousePointer = 0

63710     Exit Sub

SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

63720     intEL = Erl
63730     strES = Err.Description
63740     LogError "frmEditMicrobiology", "SaveDemographics", intEL, strES, sql

End Sub



Private Sub ValidateDemo()
63750     On Error GoTo ValidateDemo_Error
        
          Dim DVs As New DemogValidations
          Dim DV As New DemogValidation

63760     Set DV = New DemogValidation
63770     DV.SampleID = txtSampleID
63780     DV.EnteredBy = UserName
63790     DV.EnteredDateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
63800     DV.ValidatedBy = UserName
63810     DVs.Add DV
63820     DVs.Save DV

       
63830     Exit Sub

       
ValidateDemo_Error:

          Dim strES As String
          Dim intEL As Integer

63840     intEL = Erl
63850     strES = Err.Description
63860     LogError "frmEditMicrobiology", "ValidateDemo", intEL, strES
End Sub

Private Sub cmdSaveMicro_Click()
63870     bPrint.Enabled = False
63880     cmdIntrim.Enabled = False
63890     UpDown1.Enabled = False
63900     If UserHasAuthority(UserMemberOf, "MicroOtherSave") = False Then
63910         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
63920         bPrint.Enabled = True
63930         cmdIntrim.Enabled = True
63940         UpDown1.Enabled = True
63950         Exit Sub
63960     End If

63970     SaveMicro True, 0
63980     MsgBox "Record has been saved.", vbInformation
63990     LoadAllDetails
64000     bPrint.Enabled = True
64010     cmdIntrim.Enabled = True
64020     UpDown1.Enabled = True

End Sub

Private Sub SaveSensitivities(ByVal Validate As Integer)

          Dim intOrg As Integer
          Dim n As Integer
          Dim ReportCounter As Integer
          Dim sx As Sensitivity
          Dim Sxs As New Sensitivities

64030     On Error GoTo SaveSensitivities_Error
          '+++ Junaid 08-05-2024
64040     If txtSampleID.Text = "" Then
64050         Exit Sub
64060     End If
          '--- Junaid


64070     For intOrg = 1 To 4

64080         With grdAB(intOrg)

64090             ReportCounter = 0

64100             For n = 1 To .Rows - 1
64110                 If .TextMatrix(n, 0) <> "" Then
64120                     Set sx = New Sensitivity
64130                     sx.AntibioticCode = .TextMatrix(n, 6)
64140                     sx.AntibioticName = .TextMatrix(n, 0)
64150                     sx.CPOFlag = .TextMatrix(n, 2)
64160                     sx.DateTimeOfRecord = Format(Now, "dd/mmm/yyyy hh:nn:ss")
64170                     sx.UserCode = UserCode
                          '+++ Junaid 08-05-2024
                          '130                   sx.SampleID = SampleIDWithOffset
64180                     sx.SampleID = Val(txtSampleID.Text)
                          '--- Junaid
64190                     sx.IsolateNumber = intOrg
64200                     sx.RSI = .TextMatrix(n, 1)
64210                     sx.Result = Left$(.TextMatrix(n, 3), 10)
64220                     If IsDate(.TextMatrix(n, 4)) Then
64230                         sx.Rundate = Format$(.TextMatrix(n, 4), "dd/MMM/yyyy")
64240                         sx.RunDateTime = Format$(.TextMatrix(n, 4), "dd/MMM/yyyy HH:mm:ss")
64250                     Else
64260                         sx.Rundate = Format(Now, "dd/mmm/yyyy")
64270                         sx.RunDateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
64280                     End If
64290                     If .TextMatrix(n, 5) <> "" Then
64300                         sx.UserCode = .TextMatrix(n, 5)
64310                     Else
64320                         sx.UserCode = UserCode
64330                     End If

64340                     .row = n
64350                     .Col = 0
64360                     If .CellFontBold = True Then
64370                         sx.Secondary = 1
64380                     Else
64390                         sx.Secondary = 0
64400                     End If
64410                     If .CellBackColor = &HFFFFC0 Then
64420                         sx.Forced = 1
64430                     Else
64440                         sx.Forced = 0
64450                     End If

64460                     .Col = 2
64470                     If .CellPicture = 0 Then
64480                         If .TextMatrix(n, 1) = "R" Then
64490                             sx.Report = 1
64500                         ElseIf .TextMatrix(n, 1) = "S" Then
64510                             ReportCounter = ReportCounter + 1
64520                             If ReportCounter < 4 Then
64530                                 sx.Report = 1
64540                             Else
64550                                 sx.Report = 0
64560                             End If
64570                         End If
64580                     Else
64590                         If .CellPicture = imgSquareTick.Picture Then
64600                             sx.Report = 1
64610                         ElseIf .CellPicture = imgSquareCross.Picture Then
64620                             sx.Report = 0
64630                         Else
64640                             sx.Report = Null
64650                         End If
64660                     End If

64670                     If Validate = gYES Then
64680                         sx.Valid = 1
64690                         sx.AuthoriserCode = UserCode
64700                     ElseIf Validate = gNO Then
64710                         sx.Valid = 0
64720                         sx.AuthoriserCode = ""
64730                     End If
64740                     Sxs.Add sx
64750                 End If

64760             Next
64770         End With

64780     Next
64790     Sxs.Save

64800     Exit Sub

SaveSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

64810     intEL = Erl
64820     strES = Err.Description
64830     LogError "frmEditMicrobiology", "SaveSensitivities", intEL, strES

End Sub



Private Function LoadSensitivities() As Integer
          'Returns number of Isolates

          Dim intIsolate As Integer
          Dim n As Integer
          Dim t As Single
          Dim ReportCounter As Integer
          Dim Rows As Integer
          Dim sx As Sensitivity
          Dim S As String
          Dim Ax As ABDefinition
          Dim Axs(1 To 4) As New ABDefinitions
          Dim Loaded(1 To 4) As Boolean
          Dim abOK As Boolean
          Dim RetVal As Integer
          Dim ShowInGrid As Boolean

64840     On Error GoTo LoadSensitivities_Error

64850     RetVal = 0

64860     t = Timer

64870     For n = 1 To 4
64880         Loaded(n) = False
64890         With grdAB(n)
64900             .Visible = False
64910             .Rows = 2
64920             .AddItem ""
64930             .RemoveItem 1
64940         End With
64950     Next

64960     ReportCounter = 0

64970     fraMicroResult(Mic.CandS).Enabled = True

64980     If CheckIfValid() Then
64990         LockFraCS True
65000     Else
65010         LockFraCS False
65020     End If

65030     For Each sx In CurrentSensitivities

65040         intIsolate = sx.IsolateNumber
65050         If Not Loaded(intIsolate) Then
65060             Axs(intIsolate).Load cmbSite, cmbOrgGroup(intIsolate)
65070             Loaded(intIsolate) = True
65080         End If
65090         abOK = False
65100         Set Ax = Axs(intIsolate).Item(cmbSite, cmbOrgGroup(intIsolate), sx.AntibioticName)
65110         If Ax Is Nothing Then
65120             Set Ax = Axs(intIsolate).Item("Generic", cmbOrgGroup(intIsolate), sx.AntibioticName)
65130             If Not Ax Is Nothing Then
65140                 abOK = True
65150             End If
65160         Else
65170             abOK = True
65180         End If

65190         If Not sx.Forced = 1 Then
65200             If Not Ax Is Nothing Then
65210                 If abOK And sx.RSI <> "" And (sx.UserCode = "Ph" And Ax.PriSec = "P") Then    'And Not Sx.Secondary Then
65220                     sx.Report = True
65230                 End If
65240             End If
65250         End If
              '  If (abOK Or Sx.Forced) Then


65260         ShowInGrid = False
65270         If Not sx.Report And sx.UserCode <> "Ph" Then
65280             ShowInGrid = True
65290         End If
65300         If sx.Report Then
65310             ShowInGrid = True
65320         End If
65330         If abOK Then
65340             If sx.Forced Then
65350                 ShowInGrid = True
65360             End If
65370             If sx.UserCode <> "Ph" And sx.Secondary Then
65380                 ShowInGrid = True
65390             End If
65400         End If

65410         If ShowInGrid Then


65420             With grdAB(intIsolate)
65430                 S = sx.AntibioticName & vbTab & _
                          sx.RSI & vbTab & _
                          "" & vbTab & _
                          sx.Result & vbTab & _
                          sx.RunDateTime & vbTab & _
                          sx.UserCode & vbTab & _
                          sx.AntibioticCode
65440                 .AddItem S    'cpoflag goes between rsi and result
65450                 .row = .Rows - 1
65460                 .Col = 0
65470                 If sx.Forced Then
65480                     .CellBackColor = &HFFFFC0
65490                 End If
65500                 If sx.Secondary Then
65510                     .CellFontBold = True
65520                 End If
65530                 .Col = 2
10                    Set .CellPicture = IIf(sx.Report, imgSquareTick.Picture, imgSquareCross.Picture)
20                End With
30            End If
40        Next

50        For n = 1 To 4
60            If grdAB(n).Rows > 2 Then
70                grdAB(n).RemoveItem 1
80            End If
90            If cmbOrgGroup(n) <> "" Then
100               FillAbGrid (n)
110           End If
120           FillABSelect n
130           grdAB(n).Visible = True
140       Next

150       If LoadLockStatus(Mic.CandS) Then
160           fraMicroResult(Mic.CandS).Enabled = False
170       End If


180       For n = 1 To 4
190           If grdAB(n).Rows > 2 Then
200               RetVal = 1
210               Exit For
220           End If
230       Next
240       LoadSensitivities = RetVal

250       Exit Function

LoadSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditMicrobiology", "LoadSensitivities", intEL, strES

End Function

 Sub cmdSaveDemographics_Click()
290       bPrint.Enabled = False
300       cmdIntrim.Enabled = False
310       UpDown1.Enabled = False
320       pBar = 0

330       If UserHasAuthority(UserMemberOf, "MicroDemSave") = False Then
340           iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
350           bPrint.Enabled = True
360           cmdIntrim.Enabled = True
370           UpDown1.Enabled = True
380           Exit Sub
390       End If


400       If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
410           bPrint.Enabled = True
420           cmdIntrim.Enabled = True
430           UpDown1.Enabled = True
440           Exit Sub
450       End If

460       If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then
470           bPrint.Enabled = True
480           cmdIntrim.Enabled = True
490           UpDown1.Enabled = True
500           Exit Sub
510       End If

520       cmdSaveDemographics.Caption = "Saving"

          '110   If Not GetSampleIDWithOffset Then Exit Sub
          'DoEvents
          'DoEvents
530       SaveDemographics
          'DoEvents
          'DoEvents

540       UpdateMRU Me
550       Call SaveAntiBiotic(txtSampleID.Text, Trim(txtAntibiotics.Text), "Current Antibiotics")
560       Call SaveAntiBiotic(txtSampleID.Text, Trim(txtIntAntibiotics.Text), "Intended Antibiotic")

570       cmdSaveDemographics.Caption = "Save && &Hold"
580       cmdSaveDemographics.Enabled = False
590       cmdSaveInc.Enabled = False
          '      dtRunDate = Format$(Now, "dd/mm/yyyy")
          '      dtSampleDate = Format$(Now, "dd/mm/yyyy")
          'DoEvents
          'DoEvents
600       bPrint.Enabled = True
610       cmdIntrim.Enabled = True
620       UpDown1.Enabled = True

End Sub

Private Sub SaveDemographicInc()

630       On Error GoTo SaveDemographicInc_Error

640       pBar = 0

650       If UserHasAuthority(UserMemberOf, "MicroDemSave") = False Then
660           iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
670           Exit Sub
680       End If

690       If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
700           Exit Sub
710       End If

720       If lblChartNumber.BackColor = vbRed Then
730           If iMsg("Confirm this Patient has" & vbCrLf & _
                  lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
740               Exit Sub
750           End If
760       End If

770       If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

          '160   If Not GetSampleIDWithOffset Then Exit Sub

780       cmdSaveDemographics.Caption = "Saving"
          'DoEvents
          'DoEvents
790       SaveDemographics
          'DoEvents
          'DoEvents
800       UpdateMRU Me
          'DoEvents
          'DoEvents
810       Call SaveAntiBiotic(txtSampleID.Text, Trim(txtAntibiotics.Text), "Current Antibiotics")
820       Call SaveAntiBiotic(txtSampleID.Text, Trim(txtIntAntibiotics.Text), "Intended Antibiotic")
          'DoEvents
          'DoEvents
830       cmdSaveDemographics.Caption = "Save && &Hold"
840       cmdSaveDemographics.Enabled = False
850       cmdSaveInc.Enabled = False

860       txtSampleID = Format$(Val(txtSampleID) + 1)

          '240   If Not GetSampleIDWithOffset Then Exit Sub
870       LoadAllDetails
          'DoEvents
          'DoEvents
880       cmdSaveMicro.Enabled = False
890       cmdSaveHold.Enabled = False

900       Exit Sub

SaveDemographicInc_Error:

          Dim strES As String
          Dim intEL As Integer

910       intEL = Erl
920       strES = Err.Description
930       LogError "frmEditMicrobiology", "SaveDemographicInc", intEL, strES
          
          
End Sub


Private Sub cmdSaveInc_Click()

940       On Error GoTo cmdSaveInc_Click_Error
950       bPrint.Enabled = False
960       cmdIntrim.Enabled = False
970       UpDown1.Enabled = False
980       SaveDemographicInc
          '      dtRunDate = Format$(Now, "dd/mm/yyyy")
          '      dtSampleDate = Format$(Now, "dd/mm/yyyy")
          'DoEvents
          'DoEvents
990       bPrint.Enabled = True
1000      cmdIntrim.Enabled = True
1010      UpDown1.Enabled = True
1020      Exit Sub

cmdSaveInc_Click_Error:

          Dim strES As String
          Dim intEL As Integer
1030      bPrint.Enabled = True
1040      cmdIntrim.Enabled = True
1050      UpDown1.Enabled = True
1060      intEL = Erl
1070      strES = Err.Description
1080      LogError "frmEditMicrobiology", "cmdSaveInc_Click", intEL, strES
          

End Sub

Private Sub bsearch_Click()

1090      LogEvent "Name Search Click", "frmEditMicrobiology", "bsearch_Click"

1100      pBar = 0

1110      With frmPatHistory
1120          If SSTab1.Tab = 0 Then
1130              .chkShort.Value = 1
1140          Else
1150              .chkShort.Value = 0
1160          End If
1170          If HospName(0) = "Monaghan" And SSTab1.Tab = 0 Then
1180              .optBoth = True
1190          Else
1200              .optHistoric = True
1210          End If
1220          .lblDept = "M"
1230          .oFor(0) = True
1240          .txtName = Trim$(txtSurName & " " & txtForeName)
1250          .FromEdit = True
1260          .EditScreen = Me
1270          .bsearch = True
          
1280          If Not .NoPreviousDetails Then
1290              .Show 1
1300          Else
1310              FlashNoPrevious Me
1320          End If
1330      End With
1340      CheckCC
1350      LabNoUpdatePrvColor

End Sub

Private Sub cmdValidateMicro_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Validate As Integer
          Dim x As Integer
          Dim y As Integer
          Dim SID As String

1360      On Error GoTo cmdValidateMicro_Click_Error

          '20    If Not GetSampleIDWithOffset Then Exit Sub
          '+++ Junaid 20-05-2024
          '30    SID = Format$(Val(txtSampleID) + sysOptMicroOffset(0))
1370      SID = Format$(Val(txtSampleID))
          '--- Junaid
1380      If cmdValidateMicro.Caption = "&Validate" Then
1390          SaveMicro False, True
1400          If QueryGent() = 1 Then
                  'returns 0 if not paediatrics and not scbu
                  '        1 if dont report
                  '        2 if force report
1410              For x = 1 To 4
1420                  For y = 1 To grdAB(x).Rows - 1
1430                      If grdAB(x).TextMatrix(y, 0) = "Gentamicin" Then
1440                          grdAB(x).Col = 2
1450                          grdAB(x).row = y
1460                          Set grdAB(x).CellPicture = imgSquareCross.Picture
1470                      End If
1480                  Next
1490              Next
1500          End If
1510          With cmdHealthLink
1520              Set .Picture = imgHGreen.Picture
1530              ReleaseMicro SID, True

1540          End With


1550          VisibilityofCmdBtn
1560          If QueryCEF() Then Exit Sub
1570          cmdValidateMicro.Caption = "Un&Validate"
1580          cmdValidateMicro.BackColor = vbGreen
1590          LockFraCS 1
1600          Validate = 1
1610      Else
1620          sql = "Select Password from Users where " & _
                  "Name = '" & AddTicks(UserName) & "'"
1630          Set tb = New Recordset
1640          RecOpenServer 0, tb, sql
1650          If tb.EOF Then
1660              Exit Sub
1670          Else
1680              If UCase$(iBOX("Password Required", , , True)) = UCase$(tb!Password & "") Then
1690                  cmdValidateMicro.Caption = "&Validate"
1700                  cmdValidateMicro.BackColor = vbRed
1710                  With cmdHealthLink
1720                      Set .Picture = imgHRed.Picture
1730                      ReleaseMicro SID, False
1740                  End With
                      '+++ Junaid 20-05-2024
                      '410               UpdatePrintValidLog SampleIDWithOffset, "MICRO", 0, 0
1750                  UpdatePrintValidLog Trim(txtSampleID.Text), "MICRO", 0, 0
                      '--- Junaid
1760                  SaveSensitivities False
1770                  LoadSensitivities
1780                  LockFraCS 0
1790                  VisibilityofCmdBtn
1800                  Exit Sub
1810              Else
1820                  Exit Sub
1830              End If
1840          End If



1850          Validate = 0
1860      End If

1870      SaveMicro False, True
1880      SaveDemographics

          'Select Case SSTab1.Tab
          '    Case Mic.Urine: SaveUrine Validate
          '    Case Mic.Identification:
          '    Case Mic.Faeces: SaveFaeces: FillHistoricalFaeces
          '    Case Mic.CandS: SaveIsolates: SaveSensitivities Validate
          '    Case Mic.FOB: SaveFOB Validate
          '    Case Mic.RotaAdeno: SaveRotaAdeno Validate
          '    Case Mic.CDiff: SaveCdiff Validate
          '    Case Mic.OP: SaveOP Validate
          '    Case Mic.IDENTCAVAN: SaveIdentification
          '    Case Mic.RSV: SaveRSV Validate
          '    Case Mic.CSF: SaveCSF
          '    Case Mic.RedSub: SaveRedSub Validate
          'End Select

1890      PrintThis ("Save")
          '+++ Junaid 20-05-2024
          '560   UpdatePrintValidLog SampleIDWithOffset, "MICRO", Validate, 0
1900      UpdatePrintValidLog Trim(txtSampleID.Text), "MICRO", Validate, 0
          '--- Junaid

          'SaveComments
1910      UpdateMRU Me



1920      If Validate Then
1930          txtSampleID = Format$(Val(txtSampleID) + 1)

1940      End If
          '610   If Not GetSampleIDWithOffset Then Exit Sub
1950      LoadAllDetails



1960      cmdSaveMicro.Enabled = False
1970      cmdSaveHold.Enabled = False

1980      Exit Sub

cmdValidateMicro_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1990      intEL = Erl
2000      strES = Err.Description
2010      LogError "frmEditMicrobiology", "cmdValidateMicro_Click", intEL, strES, sql

End Sub

Private Sub bViewBB_Click()

2020      pBar = 0

2030      If Trim$(txtChart) <> "" Then
2040          frmViewBB.lchart = txtChart
2050          frmViewBB.Show 1
2060      End If

End Sub


Public Sub LoadAllDetails()

          Dim WasTab As Integer
          Dim SID As Long
          Dim GenResults As New GenericResults
          Dim Fxs As New FaecesResults

2070      On Error GoTo LoadAllDetails_Error

          '+++Junaid 28-08-2023
2080      dtRunDate.Value = Format$(Now, "dd/mm/yyyy")
2090      dtSampleDate.Value = Format$(Now, "dd/mm/yyyy")
2100      dtRecDate.Value = Format$(Now, "dd/mm/yyyy")
2110      tSampleTime.Mask = ""
2120      tSampleTime.Text = ""
2130      tSampleTime.Mask = "##:##"
2140      tRecTime.Mask = ""
2150      tRecTime.Text = ""
2160      tRecTime.Mask = "##:##"
          '---Junaid
2170      DoEvents
          '+++ Junaid 15-10-2023
2180      If m_Flag Then
2190          Call CheckUrineResults
2200          m_Flag = False
2210      End If

          '--- Junaid
2220      LoadingAllDetails = True
2230      ClearLabNoSelection

2240      txtMSC = ""
2250      txtClinDetails = ""
2260      txtConC = ""
2270      txtDemographicComment = ""
2280      txtUrineComment = ""
2290      txtExtSampleID = ""

          '30    ForceSaveability = False
          'Put Clear demographic function here
          'If txtSampleID < sysOptMicroOffsetOLD(0) Then Exit Sub

2300      LoadDemographics
2310      Call GetAntiBiotic(txtSampleID.Text)

2320      LockDemographics Me, False
2330      If Trim$(txtSurName) <> "" Or Trim$(txtForeName) <> "" Or Trim$(txtChart) <> "" Then
2340          LockDemographics Me, True
2350      End If
2360      LoadComments
          '160   EnableCopyFrom
2370      CheckIfPhoned
2380      CheckCC
          '190   SaveCurrentAntibiotics 'JUnaid 19-12-23
2390      lblRequestID.Caption = GetRequestID(txtSampleID.Text)


2400      If UserHasAuthority(UserMemberOf, "MicroOtherTabs") = True Then

2410          ClearCSF
2420          ClearUrine
2430          ClearFaeces
2440          ClearIndividualFaeces
2450          CheckPrintValidLog "MICRO"
2460          WasTab = SSTab1.Tab

2470          SSTab1.TabCaption(Mic.Urine) = "Urine"
2480          SSTab1.TabCaption(Mic.Identification) = "Identification"
2490          SSTab1.TabCaption(Mic.Faeces) = "Faeces"
2500          SSTab1.TabCaption(Mic.CandS) = "C && S"
2510          SSTab1.TabCaption(Mic.FOB) = "FOB"
2520          SSTab1.TabCaption(Mic.RotaAdeno) = "Rota/Adeno"
2530          SSTab1.TabCaption(Mic.CSF) = "CSF"
2540          SSTab1.TabCaption(Mic.CDiff) = "C.diff"
2550          SSTab1.TabCaption(Mic.OP) = "OP"
2560          SSTab1.TabCaption(Mic.IDENTCAVAN) = "Identification"
2570          SSTab1.TabCaption(Mic.RedSub) = "Red/Sub"
2580          SSTab1.TabCaption(Mic.RSV) = "RSV"

2590          SID = Val(txtSampleID) ' + sysOptMicroOffset(0)

2600          GenResults.Load Format$(SID)
2610          Fxs.Load Format$(SID)
2620          Set CurrentSensitivities = New Sensitivities
2630          CurrentSensitivities.Load SID
2640          txtCommentMicro.Text = CheckAutoCommentsMicro(SID)    ' Masood 01-Oct-2015
2650          SSTab1.TabVisible(Mic.Urine) = False
2660          SSTab1.TabVisible(Mic.IDENTCAVAN) = True
2670          SSTab1.TabVisible(Mic.Identification) = False
2680          SSTab1.TabVisible(Mic.Faeces) = False
2690          SSTab1.TabVisible(Mic.CandS) = True
2700          SSTab1.TabVisible(Mic.Faeces) = False
2710          SSTab1.TabVisible(Mic.FOB) = False
2720          SSTab1.TabVisible(Mic.CSF) = False
2730          SSTab1.TabVisible(Mic.RotaAdeno) = False
2740          SSTab1.TabVisible(Mic.CDiff) = False
2750          SSTab1.TabVisible(Mic.OP) = False
2760          SSTab1.TabVisible(Mic.RSV) = False
2770          SSTab1.TabVisible(Mic.RedSub) = False

2780          FaecesLoaded = False
2790          IdentLoaded = False
2800          CSLoaded = False
2810          FOBLoaded = False
2820          RotaAdenoLoaded = False
2830          CdiffLoaded = False
2840          OPLoaded = False
2850          IdentificationLoaded = False
2860          CSFLoaded = False


2870          If LoadUrine() > 0 Then

2880              If chkPregnant.Value = 0 Then CheckUrineAutoVal

2890              SSTab1.TabVisible(Mic.Urine) = True
2900              cmbSite = "Urine"
2910              If UrineResultsPresent() Then
2920                  SSTab1.TabCaption(Mic.Urine) = "<<Urine>>"
2930              End If
2940              LoadPrintStatus Mic.Urine
2950          End If

2960          If cmbSite = "Faeces" Then
2970              If UCase(HospName(0)) = "CAVAN" Then
2980                  SSTab1.TabVisible(Mic.FOB) = True
2990                  If LoadFOB(Fxs) Then
3000                      SSTab1.TabCaption(Mic.FOB) = "<<FOB>>"
3010                      FOBLoaded = True
3020                  End If
3030                  SSTab1.TabVisible(Mic.RotaAdeno) = True
3040                  If LoadRotaAdeno(Fxs) Then
3050                      SSTab1.TabCaption(Mic.RotaAdeno) = "<<Rota/Adeno>>"
3060                      RotaAdenoLoaded = True
3070                  End If
3080                  SSTab1.TabVisible(Mic.CDiff) = True
3090                  If LoadCDiff(GenResults, Fxs) Then
3100                      SSTab1.TabCaption(Mic.CDiff) = "<<C.diff>>"
3110                      CdiffLoaded = True
3120                  End If
3130                  SSTab1.TabVisible(Mic.OP) = True
3140                  If LoadOP(Fxs) Then
3150                      SSTab1.TabCaption(Mic.OP) = "<<OP>>"
3160                      OPLoaded = True
3170                  End If
3180                  SSTab1.TabVisible(Mic.Faeces) = True
3190                  If LoadFaeces() Then
3200                      SSTab1.TabCaption(Mic.Faeces) = "<<Faeces>>"
3210                      FaecesLoaded = True
3220                  End If
3230              End If
3240          ElseIf cmbSite = "CSF" Or UCase$(cmbSite) = "CEREBROSPINAL FLUID" Then
3250              SSTab1.TabVisible(Mic.CSF) = True
3260              If LoadCSF() Then
3270                  SSTab1.TabCaption(Mic.CSF) = "<<CSF>>"
3280                  CSFLoaded = True
3290              End If
3300          End If

3310          Select Case LoadIdentification()
                  Case 0:
3320                  IdentificationLoaded = False
3330              Case 1, 2, 3, 4:
3340                  SSTab1.TabCaption(Mic.IDENTCAVAN) = "<<Identification>>"
3350                  IdentificationLoaded = True
3360          End Select

3370          LoadIsolates
              'CSLoaded = False

3380          If LoadSensitivities() > 0 Then
3390              SSTab1.TabCaption(Mic.CandS) = "<<C && S>>"
3400              CSLoaded = True
3410          Else
3420              CSLoaded = False
3430          End If

3440          SSTab1.TabVisible(Mic.Identification) = False
3450          SSTab1.TabVisible(4) = True

              '1070  LoadComments

3460          If LoadRSV(GenResults) Then
3470              SSTab1.TabCaption(Mic.RSV) = "<<RSV>>"
3480          End If

3490          If LoadRedSub(GenResults) Then
3500              SSTab1.TabCaption(Mic.RedSub) = "<<R/S>>"
3510          End If

3520          FillHistoricalFaeces
3530          FillForConsultantValidation
3540          cmdCopySensitivities.Visible = False
3550          If IsAnyRecordPresent("Isolates", SID - 1) Then
3560              If IsAnyRecordPresent("Sensitivities", SID - 1) Then
3570                  If Not IsAnyRecordPresent("Isolates", SID) Then
3580                      If Not IsAnyRecordPresent("Sensitivities", SID) Then
3590                          cmdCopySensitivities.Caption = "Copy from " & Val(txtSampleID) - 1
3600                          cmdCopySensitivities.Visible = False
                              '1400                      cmdCopySensitivities.Visible = True
3610                      End If
3620                  End If
3630              End If
3640          End If

3650          LoadPrintStatus Mic.FOB
3660          LoadPrintStatus Mic.RotaAdeno
3670          LoadPrintStatus Mic.CDiff
3680          LoadPrintStatus Mic.OP
3690          LoadPrintStatus Mic.CSF
3700          LoadPrintStatus Mic.RSV
3710          LoadPrintStatus Mic.RedSub
3720          LoadPrintStatus Mic.CandS

3730          If SSTab1.TabVisible(WasTab) Then
3740              SSTab1.Tab = WasTab
3750          End If

3760          cmdIsoRepeat.Visible = IsAnyRecordPresent("IsolatesRepeats", SID)
3770          cmdIsoArchive.Visible = IsAnyRecordPresent("IsolatesArc", SID)
3780          cmdSensRepeat.Visible = IsAnyRecordPresent("SensitivitiesRepeats", SID)
3790          cmdSensArchive.Visible = IsAnyRecordPresent("SensitivitiesArc", SID)

              '1440  If ForceSaveability Then
              '1450    cmdSaveMicro.Enabled = True
              '1460    cmdSaveHold.Enabled = True
              '1470  End If

3800          SetViewReports txtSampleID
              '+++ Junaid
              '        cmdViewReports.Visible = True
              '--- Junaid

3810          If IsMicroReleased(SID) Then
3820              cmdHealthLink.BackColor = vbRed
                  'Set cmdHealthLink.Picture = imgHGreen.Picture
3830          Else
3840              cmdHealthLink.BackColor = vbWhite
                  'Set cmdHealthLink.Picture = imgHRed.Picture
3850          End If
3860          CheckIfExternalReleased txtSampleID
3870      End If

3880      CheckMicroUrineComment txtSampleID

3890      LoadingAllDetails = False
3900      MatchingDemoLoaded = False


3910      If UserHasAuthority(UserMemberOf, "EnableMicroNotepad") Then
3920          If IsNotepadExists(Trim(SID & ""), "") = True Then
3930              cmdPatientNotePad.BackColor = vbYellow
3940          Else
3950              cmdPatientNotePad.BackColor = &H8000000F
3960          End If
3970      Else
3980          cmdPatientNotePad.Visible = False
3990      End If


4000      ReportButtonColor
4010      VisibilityofCmdBtn
4020      Exit Sub

LoadAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

4030      intEL = Erl
4040      strES = Err.Description
4050      LogError "frmEditMicrobiology", "LoadAllDetails", intEL, strES
4060      LoadingAllDetails = False

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportButtonColor
' Author    : Masood
' Date      : 15/Jan/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ReportButtonColor()
          Dim ReportType As String
4070      On Error GoTo ReportButtonColor_Error


4080      ReportType = GetMicroReportType(txtSampleID)

4090      cmdIntrim.BackColor = &H8000000F
4100      bPrint.BackColor = &H8000000F

4110      If UCase(ReportType) = UCase("Final Report") Then
4120          bPrint.BackColor = vbGreen
4130      ElseIf UCase(ReportType) = UCase("Interim Report") Then
4140          cmdIntrim.BackColor = vbGreen
4150      End If


4160      Exit Sub


ReportButtonColor_Error:

          Dim strES As String
          Dim intEL As Integer

4170      intEL = Erl
4180      strES = Err.Description
4190      LogError "frmEditMicrobiology", "ReportButtonColor", intEL, strES

End Sub

Private Sub CheckIfExternalReleased(ByVal SID As String)

          Dim sql As String
          Dim tb As Recordset

4200      On Error GoTo CheckIfExternalReleased_Error
          '+++ Junaid 12-01-2024
          '20    Sql = "SELECT * FROM BiomnisRequests WHERE SampleID = " & CLng(SID) + sysOptMicroOffset(0)
4210      sql = "SELECT * FROM BiomnisRequests WHERE SampleID = '" & SID & "'"
          '--- Junaid
4220      Set tb = New Recordset
4230      RecOpenServer 0, tb, sql
4240      cmdOrderBiomnis(0).BackColor = vbWhite
4250      cmdOrderBiomnis(1).BackColor = vbWhite
4260      While Not tb.EOF
4270          If tb!SendTo & "" = "Biomnis" Then
4280              cmdOrderBiomnis(0).BackColor = vbRed
4290          ElseIf tb!SendTo & "" = "MAT: Mater Hospital" Then
4300              cmdOrderBiomnis(1).BackColor = vbRed
4310          End If
4320          tb.MoveNext
4330      Wend


4340      Exit Sub

CheckIfExternalReleased_Error:

          Dim strES As String
          Dim intEL As Integer

4350      intEL = Erl
4360      strES = Err.Description
4370      LogError "frmEditMicrobiology", "CheckIfExternalReleased", intEL, strES, sql
          
End Sub


Private Sub bcancel_Click()

4380      pBar = 0

4390      Me.Hide

End Sub

Private Sub cmbClinDetails_Click()

4400      txtClinDetails = txtClinDetails & cmbClinDetails & " "
4410      cmbClinDetails.ListIndex = -1

4420      cmdSaveDemographics.Enabled = True
4430      cmdSaveInc.Enabled = True

End Sub


Private Sub cmbClinDetails_LostFocus()

          Dim tb As Recordset
          Dim sql As String

4440      On Error GoTo cmbClinDetails_LostFocus_Error

4450      pBar = 0

4460      If Trim$(cmbClinDetails) = "" Then Exit Sub

4470      sql = "Select * from Lists where " & _
              "ListType = 'CD' " & _
              "and Code = '" & cmbClinDetails & "' and InUse = 1"
4480      Set tb = New Recordset
4490      RecOpenServer 0, tb, sql
4500      If Not tb.EOF Then
4510          cmbClinDetails = tb!Text & ""
4520      End If

4530      Exit Sub

cmbClinDetails_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

4540      intEL = Erl
4550      strES = Err.Description
4560      LogError "frmEditMicrobiology", "cmbClinDetails_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbClinician_Click()

4570      cmdSaveDemographics.Enabled = True
4580      cmdSaveInc.Enabled = True

End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

4590      cmdSaveDemographics.Enabled = True
4600      cmdSaveInc.Enabled = True

End Sub


Private Sub cmbClinician_LostFocus()

4610      pBar = 0
4620      cmbClinician = QueryKnown(cmbClinician, cmbHospital)

End Sub

Private Sub cmbGP_Change()

4630      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub cmbGP_Click()

4640      pBar = 0

4650      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

4660      cmdSaveDemographics.Enabled = True
4670      cmdSaveInc.Enabled = True

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

4680      cmdSaveDemographics.Enabled = True
4690      cmdSaveInc.Enabled = True

End Sub


Private Sub cmbGP_LostFocus()

          Dim strOrig As String
          Dim Gx As New GP

          Dim S As String
          Dim GXs As New GPs

4700      pBar = 0

4710      strOrig = cmbGP

4720      cmbGP = ""

4730      Gx.LoadCodeOrText strOrig
4740      cmbGP = Gx.Text
4750      If sysOptAllowGPFreeText(0) And cmbGP = "" Then
4760          cmbGP = strOrig
4770      End If

4780      If cmdCopyTo.Caption = "cc" Then
4790          If GXs.GpCCed(ListCodeFor("HO", cmbHospital), cmbGP) Then
4800              S = cmbWard & " " & cmbClinician
4810              S = Trim$(S) & " " & cmbGP
4820              S = Trim$(S)

4830              frmCopyTo.lblOriginal = S
4840              frmCopyTo.lblSampleID = txtSampleID
4850              frmCopyTo.Show 1

4860              CheckCC
4870          End If
4880      End If

End Sub


Private Sub cmdSetPrinter_Click()

4890      frmForcePrinter.From = Me
4900      frmForcePrinter.Show 1

4910      If pPrintToPrinter = "Automatic Selection" Then
4920          pPrintToPrinter = ""
4930      End If

4940      If pPrintToPrinter <> "" Then
4950          cmdSetPrinter.BackColor = vbRed
4960          cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
4970      Else
4980          cmdSetPrinter.BackColor = vbButtonFace
4990          pPrintToPrinter = ""
5000          cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
5010      End If

End Sub

Private Sub cMRU_Click()

5020      txtSampleID = cMRU

          '20    If Not GetSampleIDWithOffset Then Exit Sub

5030      LoadAllDetails

5040      cmdSaveDemographics.Enabled = False
5050      cmdSaveInc.Enabled = False
5060      cmdSaveMicro.Enabled = False
5070      cmdSaveHold.Enabled = False

End Sub


Private Sub cMRU_KeyPress(KeyAscii As Integer)

5080      KeyAscii = 0

End Sub


Private Sub cRooH_Click(Index As Integer)

5090      cmdSaveDemographics.Enabled = True
5100      cmdSaveInc.Enabled = True

End Sub

Private Sub cmbWard_Change()

5110      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub cmbWard_Click()

5120      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

5130      cmdSaveDemographics.Enabled = True
5140      cmdSaveInc.Enabled = True

End Sub


Private Sub cmbWard_KeyPress(KeyAscii As Integer)

5150      cmdSaveDemographics.Enabled = True
5160      cmdSaveInc.Enabled = True

End Sub


Private Sub cmbWard_LostFocus()

          Dim Hospital As String

5170      Hospital = ListCodeFor("HO", HospName(0))

5180      cmbWard = GetWard(cmbWard, Hospital)

5190      If Trim$(cmbWard) = "" Then
5200          cmbWard = "GP"
5210          Exit Sub
5220      End If

End Sub



Private Sub dtRunDate_CloseUp()

5230      pBar = 0

5240      cmdSaveDemographics.Enabled = True
5250      cmdSaveInc.Enabled = True

End Sub


Private Sub dtSampleDate_CloseUp()

5260      pBar = 0

5270      cmdSaveDemographics.Enabled = True
5280      cmdSaveInc.Enabled = True

End Sub


Private Sub Form_Activate()


5290      pBar = 0
5300      pBar.max = LogOffDelaySecs
5310      TimerBar.Enabled = True


5320      If UserMemberOf = "Managers" Then
5330          cmdSetValid.Visible = True
5340      Else
5350          cmdSetValid.Visible = False
5360      End If

5370      If GetOptionSetting("HealthLinkDeptMicro", "0") <> "1" Then
5380          cmdHealthLink.Visible = False
5390      End If
5400      ShowMenuLists

          '+++ Junaid
5410      cmdViewReports.Visible = True
          '--- Junaid

          'If InStr(UCase$(App.Path), "TEST") Then
          '    cmdPatientNotePad.Visible = True
          'Else
          '    cmdPatientNotePad.Visible = False
          'End If

End Sub

Private Sub FillOrganisms()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim Temp As String

5420      On Error GoTo FillOrganisms_Error

5430      sql = "Select Text from Lists where " & _
              "ListType = 'OR' and InUse = 1 " & _
              "order by ListOrder"
5440      Set tb = New Recordset
5450      RecOpenServer 0, tb, sql

5460      For n = 1 To 4
5470          cmbOrgGroup(n).Clear
5480          cmbOrgName(n).Clear
5490      Next

5500      Do While Not tb.EOF
5510          Temp = tb!Text & ""
5520          For n = 1 To 4
5530              cmbOrgGroup(n).AddItem Temp
5540          Next
5550          tb.MoveNext
5560      Loop
5570      For n = 1 To 4
5580          FixComboWidth cmbOrgGroup(n)
5590      Next n

5600      Exit Sub

FillOrganisms_Error:

          Dim strES As String
          Dim intEL As Integer

5610      intEL = Erl
5620      strES = Err.Description
5630      LogError "frmEditMicrobiology", "FillOrganisms", intEL, strES, sql

End Sub

Private Sub ClearFaeces()

          Dim x As Integer
          Dim y As Integer

5640      For x = 1 To 3
5650          For y = 1 To 3
5660              cmbDay1(y * 10 + x) = ""
5670              cmbDay2(y * 10 + x) = ""
5680          Next
5690          cmbDay3(x) = ""
5700      Next

End Sub
Private Sub FillCastsCrystalsMiscSite()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String

5710      On Error GoTo FillCastsCrystalsMiscSite_Error

5720      cmbCasts.Clear
5730      cmbCrystals.Clear
5740      cmbMisc(0).Clear
5750      cmbMisc(1).Clear
5760      cmbMisc(2).Clear
5770      cmbSite.Clear
5780      cmbClinDetails.Clear

5790      For n = 1 To 4
5800          cmbQualifier(n).Clear
5810          cmbIdentification(n).Clear
5820      Next

5830      sql = "SELECT Text, UPPER(ListType) ListType FROM Lists " & _
              "WHERE " & _
              "ListType IN ('IN', 'CA', 'CR', 'MI', 'SI', 'OV', 'CD', 'HO', 'MQ') " & _
              "AND InUse = 1 " & _
              "ORDER BY ListOrder"
5840      Set tb = New Recordset
5850      RecOpenClient 0, tb, sql
5860      Do While Not tb.EOF
5870          Select Case tb!ListType
                  Case "IN"
5880                  For n = 1 To 4
5890                      cmbIdentification(n).AddItem tb!Text & ""
5900                  Next
5910              Case "CA"
5920                  cmbCasts.AddItem tb!Text & ""
5930              Case "CR"
5940                  cmbCrystals.AddItem tb!Text & ""
5950              Case "MI"
5960                  cmbMisc(0).AddItem tb!Text & ""
5970                  cmbMisc(1).AddItem tb!Text & ""
5980                  cmbMisc(2).AddItem tb!Text & ""
5990              Case "SI"
6000                  cmbSite.AddItem tb!Text & ""
6010              Case "OV"
6020                  cmbOva(0).AddItem tb!Text & ""
6030                  cmbOva(1).AddItem tb!Text & ""
6040                  cmbOva(2).AddItem tb!Text & ""
6050              Case "CD"
6060                  cmbClinDetails.AddItem tb!Text & ""
6070              Case "HO"
6080                  cmbHospital.AddItem tb!Text & ""
6090              Case "MQ"
6100                  For n = 1 To 4
6110                      cmbQualifier(n).AddItem tb!Text & ""
6120                  Next
6130          End Select
6140          tb.MoveNext
6150      Loop

6160      For n = 1 To 4
6170          FixComboWidth cmbIdentification(n)
6180          FixComboWidth cmbQualifier(n)
6190      Next n
6200      FixComboWidth cmbCasts
6210      FixComboWidth cmbCrystals
6220      FixComboWidth cmbSite
6230      For n = 0 To 2
6240          FixComboWidth cmbMisc(n)
6250          FixComboWidth cmbOva(n)
6260      Next n
6270      FixComboWidth cmbClinDetails
6280      FixComboWidth cmbHospital


6290      Exit Sub

FillCastsCrystalsMiscSite_Error:

          Dim strES As String
          Dim intEL As Integer

6300      intEL = Erl
6310      strES = Err.Description
6320      LogError "frmEditMicrobiology", "FillCastsCrystalsMiscSite", intEL, strES, sql

End Sub
Private Sub FillFaecesLists()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer
          Dim y As Integer

6330      On Error GoTo FillFaecesLists_Error

6340      For x = 1 To 3
6350          For y = 1 To 3
6360              cmbDay1(y * 10 + x).Clear
6370              cmbDay2(y * 10 + x).Clear
6380          Next
6390          cmbDay3(x).Clear
6400      Next

6410      sql = "Select Text from Lists where " & _
              "ListType = 'FX' and InUse = 1 order by ListOrder"
6420      Set tb = New Recordset
6430      RecOpenServer 0, tb, sql
6440      Do While Not tb.EOF
6450          For x = 1 To 3
6460              For y = 1 To 2
6470                  cmbDay1(y * 10 + x).AddItem tb!Text & ""
6480                  cmbDay2(y * 10 + x).AddItem tb!Text & ""
6490              Next
6500          Next
6510          tb.MoveNext
6520      Loop

6530      sql = "Select Text from Lists where " & _
              "ListType = 'FS' and InUse = 1 order by ListOrder"
6540      Set tb = New Recordset
6550      RecOpenServer 0, tb, sql
6560      Do While Not tb.EOF
6570          For x = 1 To 3
6580              cmbDay1(30 + x).AddItem tb!Text & ""
6590          Next
6600          tb.MoveNext
6610      Loop

6620      sql = "Select Text from Lists where " & _
              "ListType = 'FP' and InUse = 1 order by ListOrder"
6630      Set tb = New Recordset
6640      RecOpenServer 0, tb, sql
6650      Do While Not tb.EOF
6660          For x = 1 To 3
6670              cmbDay2(30 + x).AddItem tb!Text & ""
6680              cmbDay3(x).AddItem tb!Text & ""
6690          Next
6700          tb.MoveNext
6710      Loop

6720      For x = 1 To 3
6730          FixComboWidth cmbDay3(x)
6740          For y = 1 To 3
6750              FixComboWidth cmbDay1(y * 10 + x)
6760              FixComboWidth cmbDay2(y * 10 + x)
6770          Next
6780      Next


6790      Exit Sub

FillFaecesLists_Error:

          Dim strES As String
          Dim intEL As Integer

6800      intEL = Erl
6810      strES = Err.Description
6820      LogError "frmEditMicrobiology", "FillFaecesLists", intEL, strES, sql

End Sub

Private Sub FillAbGrid(ByVal Index As Integer)

          Dim n As Integer
          Dim y As Integer
          Dim Found As Boolean
          Dim Ax As ABDefinition
          Dim Axs As New ABDefinitions

6830      On Error GoTo FillAbGrid_Error

6840      Axs.Load cmbSite, cmbOrgGroup(Index)

6850      For Each Ax In Axs
6860          Found = False
6870          For n = 1 To grdAB(Index).Rows - 1
6880              If Trim$(grdAB(Index).TextMatrix(n, 0)) = Ax.AntibioticName Then
6890                  If IsChild() And Not Ax.AllowIfChild Then
6900                      grdAB(Index).TextMatrix(n, 2) = "C"
6910                  ElseIf IsPregnant() And Not Ax.AllowIfPregnant Then
6920                      grdAB(Index).TextMatrix(n, 2) = "P"
6930                  ElseIf IsOutPatient() And Not Ax.AllowIfOutPatient Then
6940                      grdAB(Index).TextMatrix(n, 2) = "O"
6950                  End If
6960                  Found = True
6970                  Exit For
6980              End If
6990          Next
7000          If Not Found And Ax.PriSec = "P" Then
7010              grdAB(Index).AddItem Trim$(Ax.AntibioticName) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & Ax.Code & ""
7020              grdAB(Index).row = grdAB(Index).Rows - 1
7030              grdAB(Index).Col = 2
7040              If IsChild() And Not Ax.AllowIfChild Then
7050                  Set grdAB(Index).CellPicture = imgSquareCross.Picture
                      '170         grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "C"
7060              ElseIf IsPregnant() And Not Ax.AllowIfPregnant Then
7070                  Set grdAB(Index).CellPicture = imgSquareCross.Picture
                      '200         grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "P"
7080              ElseIf IsOutPatient() And Not Ax.AllowIfOutPatient Then
7090                  Set grdAB(Index).CellPicture = imgSquareCross.Picture
                      '230         grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "O"
7100              Else
7110                  Set grdAB(Index).CellPicture = Me.Picture
7120              End If
7130          End If
7140      Next

          'Set sx = CurrentSensitivities.Item(Index, Ax.Code)
          'If Not sx Is Nothing Then
          '  grdAB(Index).TextMatrix(grdAB(Index).Rows - 1, 1) = sx.RSI
          'End If

7150      For n = 0 To lstABsInUse.ListCount - 1
7160          If cmbOrgGroup(Index) <> "Negative Results" And lstABsInUse.List(n) <> "Antibiotic Not Stated" And lstABsInUse.List(n) <> "None" Then
7170              Found = False
7180              For y = 1 To grdAB(Index).Rows - 1
7190                  If grdAB(Index).TextMatrix(y, 0) = lstABsInUse.List(n) Then
7200                      Found = True
7210                      Exit For
7220                  End If
7230              Next
7240              If Not Found Then
                      '390         grdAB(Index).AddItem lstABsInUse.List(n)
                      '400         grdAB(Index).Row = grdAB(Index).Rows - 1
7250              Else
7260                  grdAB(Index).row = y
7270              End If
7280              grdAB(Index).Col = 2
                  '470       Set grdAB(Index).CellPicture = imgSquareTick.Picture
7290          End If
7300      Next
7310      Exit Sub

FillAbGrid_Error:

          Dim strES As String
          Dim intEL As Integer

7320      intEL = Erl
7330      strES = Err.Description
7340      LogError "frmEditMicrobiology", "FillAbGrid", intEL, strES

End Sub

Private Sub Form_Deactivate()

7350      pBar = 0
7360      TimerBar.Enabled = False

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

7370      pBar = 0

End Sub

Private Sub Form_Load()

          Dim i As Integer
          '10        CheckIdentificationInDb
          '20        CheckIdentificationArcInDb
7380      m_Flag = False
7390      Me.Caption = "NetAcquire - Microbiology (" & HospName(0) & ")"
7400      FillLists
7410      FillOrganisms
7420      FillCurrentABs
7430      FillMSandConsultantComment
7440      FillMRU Me
          '          DoEvents
          '          DoEvents
          '          Call FormatGrid
7450      fmeQuestions.Visible = False
7460      With lblChartNumber
7470          .BackColor = &H8000000F
7480          .ForeColor = vbBlack
7490          Select Case UCase$(HospName(0))
                  Case "CAVAN"
7500                  .Caption = "Cavan Chart #"
7510              Case "MONAGHAN"
7520                  .Caption = "Monaghan Chart #"
7530          End Select
7540      End With

          '180       dtRunDate = Format$(Now, "dd/mm/yyyy")
          '190       dtSampleDate = Format$(Now, "dd/mm/yyyy")

7550      UpDown1.max = 2147483647

7560      If pFromViewReportSID <> "" Then
7570          txtSampleID = pFromViewReportSID
7580      Else
7590          txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")
7600      End If
          '260       If Not GetSampleIDWithOffset Then Exit Sub

7610      SSTab1.TabVisible(0) = UserHasAuthority(UserMemberOf, "MicroDemTab")
7620      For i = 1 To 12
7630          SSTab1.TabVisible(i) = UserHasAuthority(UserMemberOf, "MicroOtherTabs")
7640      Next i


7650      LoadAllDetails
          '          DoEvents
          '          DoEvents
7660      cmdSaveDemographics.Enabled = False
7670      cmdSaveInc.Enabled = False
7680      cmdSaveMicro.Enabled = False
7690      cmdSaveHold.Enabled = False

7700      fraDipStick.Visible = sysOptDipStick(0)
7710      fraUrineSpecific.Visible = frmOptUrineSpecific

7720      cmdValidateMicro.Enabled = True

7730      LoadListBacteria
7740      LoadListRCC
7750      LoadListWCC
          '          DoEvents
          '          DoEvents
7760      Activated = False
7770      MatchingDemoLoaded = False
7780      FormLoaded = True



End Sub
Private Sub LoadDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleDate As String
          Dim RooH As Boolean

7790      On Error GoTo LoadDemographics_Error

7800      RooH = IsRoutine()
7810      cRooH(0) = RooH
7820      cRooH(1) = Not RooH
7830      bViewBB.Enabled = False
7840      lstABsInUse.Clear
7850      cmbSite = ""
7860      txtSiteDetails = ""
7870      lstABsInUse.Clear
7880      lblABsInUse = ""
7890      cmbClinDetails = ""

7900      If Trim$(txtSampleID) = "" Then Exit Sub

7910      If FormLoaded Then txtLabNo = ""          'Val(FndMaxID("demographics", "LabNo", ""))

          Dim SDS As New SiteDetails
          '+++ Junaid 20-05-2024
          '140   SDS.Load SampleIDWithOffset
7920      SDS.Load Trim(txtSampleID.Text)
          '--- Junaid
7930      If SDS.Count > 0 Then
7940          cmbSite = SDS(1).Site
7950          txtSiteDetails.Text = Trim(SDS(1).SiteDetails)
7960          DoEvents
7970      End If

          Dim CURS As New CurrentAntibiotics
          Dim Cur As CurrentAntibiotic
          '+++ Junaid 20-05-2024
          '190   CURS.Load SampleIDWithOffset
7980      CURS.Load Trim(txtSampleID.Text)
          '--- Junaid
7990      For Each Cur In CURS
8000          lstABsInUse.AddItem Cur.Antibiotic
8010          lblABsInUse = lblABsInUse & Cur.Antibiotic & " "
8020      Next
          '+++ Junaid 20-05-2024
          '240   sql = "Select * from Demographics where " & _
          '            "SampleID = '" & SampleIDWithOffset & "'"
8030      sql = "Select * from Demographics where " & _
              "SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
8040      Set tb = New Recordset
8050      RecOpenClient 0, tb, sql
8060      If tb.EOF Or txtSampleID < sysOptMicroOffsetOLD(0) Then
8070          mNewRecord = True
              '290       dtRunDate = Format$(Now, "dd/mm/yyyy")
              '300       dtSampleDate = Format$(Now, "dd/mm/yyyy")
8080          dtRecDate = dtSampleDate
8090          tRecTime.Mask = ""
8100          tRecTime.Text = ""
8110          tRecTime.Mask = "##:##"
8120          txtChart = ""
8130          txtSurName = ""
8140          txtForeName = ""
8150          txtAddress(0) = ""
8160          txtAddress(1) = ""
8170          txtSex = ""
8180          txtDoB = ""
8190          txtAge = ""
8200          cmbWard = "GP"
8210          cmbClinician = ""
8220          cmbGP = ""
8230          cmbHospital = HospName(0)
8240          txtClinDetails = ""
8250          txtDemographicComment = ""
8260          tSampleTime.Mask = ""
8270          tSampleTime.Text = ""
8280          tSampleTime.Mask = "##:##"
8290          lblChartNumber.Caption = HospName(0) & " Chart #"
8300          lblChartNumber.BackColor = &H8000000F
8310          lblChartNumber.ForeColor = vbBlack
8320          chkPregnant = 0
8330      Else
8340          If Trim$(ConvertNull(tb!Hospital, "") & "") <> "" Then
8350              cmbHospital = Trim$(tb!Hospital)
8360              lblChartNumber = Trim$(tb!Hospital) & " Chart #"
8370              If UCase$(tb!Hospital) = UCase$(HospName(0)) Then
8380                  lblChartNumber.BackColor = &H8000000F
8390                  lblChartNumber.ForeColor = vbBlack
8400              Else
8410                  lblChartNumber.BackColor = vbRed
8420                  lblChartNumber.ForeColor = vbYellow
8430              End If
8440          Else
8450              cmbHospital = HospName(0)
8460              lblChartNumber.Caption = HospName(0) & " Chart #"
8470              lblChartNumber.BackColor = &H8000000F
8480              lblChartNumber.ForeColor = vbBlack
8490          End If
8500          If IsDate(tb!SampleDate) Then
8510              dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
8520          Else
8530              dtSampleDate = Format$(Now, "dd/mm/yyyy")
8540          End If
8550          If IsDate(tb!Rundate) Then
8560              dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
8570          Else
8580              dtRunDate = Format$(Now, "dd/mm/yyyy")
8590          End If
8600          If IsDate(tb!RecDate) Then
8610              dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
8620          Else
8630              dtRecDate = dtRunDate
8640          End If
8650          mNewRecord = False
8660          cRooH(0) = ConvertNull(tb!RooH, "0")
8670          cRooH(1) = Not ConvertNull(tb!RooH, "0")
8680          txtChart = ConvertNull(tb!Chart, "") & ""
8690          txtSurName = SurName(tb!PatName & "")
8700          txtForeName = ForeName(tb!PatName & "")
8710          txtAddress(0) = ConvertNull(tb!Addr0, "") & ""
8720          txtAddress(1) = ConvertNull(tb!Addr1, "") & ""
8730          If ConvertNull(tb!LabNo, "") & "" <> "" Then
8740              txtLabNo = ConvertNull(tb!LabNo, "") & ""
8750          End If
8760          Select Case Left$(Trim$(UCase$(tb!Sex & "")), 1)
                  Case "M": txtSex = "Male"
8770              Case "F": txtSex = "Female"
8780              Case Else: txtSex = ""
8790          End Select
8800          txtDoB = Format$(tb!DoB, "dd/mm/yyyy")
8810          txtAge = ConvertNull(tb!Age, "") & ""
8820          cmbWard = ConvertNull(tb!Ward, "") & ""
8830          cmbClinician = ConvertNull(tb!Clinician, "") & ""
8840          cmbGP = ConvertNull(tb!GP, "") & ""
8850          txtClinDetails = ConvertNull(tb!ClDetails, "") & ""
8860          txtExtSampleID = ConvertNull(tb!ExtSampleID, "") & ""
8870          If IsDate(ConvertNull(tb!SampleDate, 0)) Then
8880              dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
8890              If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
8900                  tSampleTime = Format$(tb!SampleDate, "hh:mm")
8910              Else
8920                  tSampleTime.Mask = ""
8930                  tSampleTime.Text = ""
8940                  tSampleTime.Mask = "##:##"
8950              End If
8960          Else
8970              dtSampleDate = Format$(Now, "dd/mm/yyyy")
8980              tSampleTime.Mask = ""
8990              tSampleTime.Text = ""
9000              tSampleTime.Mask = "##:##"
9010          End If
9020          If IsDate(tb!RecDate & "") Then
9030              dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
9040              If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
9050                  tRecTime = Format$(tb!RecDate, "hh:mm")
9060              Else
9070                  tRecTime.Mask = ""
9080                  tRecTime.Text = ""
9090                  tRecTime.Mask = "##:##"
9100              End If
9110          Else
9120              dtRecDate = dtSampleDate
9130              tRecTime.Mask = ""
9140              tRecTime.Text = ""
9150              tRecTime.Mask = "##:##"
9160          End If
9170          If IsNull(tb!Pregnant) Then
9180              chkPregnant = 0
9190          Else
9200              chkPregnant = IIf(tb!Pregnant, 1, 0)
9210          End If
9220      End If
9230      cmdSaveDemographics.Enabled = False
9240      cmdSaveInc.Enabled = False

9250      If sysOptBloodBank(0) Then
9260          If Trim$(txtChart) <> "" Then
9270              sql = "Select  * from PatientDetails where " & _
                      "PatNum = '" & txtChart & "'"
9280              Set tb = New Recordset
9290              RecOpenClientBB 0, tb, sql
9300              bViewBB.Enabled = Not tb.EOF
9310          End If
9320      End If
9330      SetViewScans Val(txtSampleID), cmdViewScan
9340      CheckIfExternalReleased txtSampleID
9350      Screen.MousePointer = 0

9360      Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

9370      intEL = Erl
9380      strES = Err.Description
9390      LogError "frmEditMicrobiology", "LoadDemographics", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

9400      pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

9410      If Val(txtSampleID) > Val(GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")) Then
9420          SaveSetting "NetAcquire", "StartUp", "LastUsedMicro", txtSampleID
9430      End If

9440      pPrintToPrinter = ""

9450      Activated = False

End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

9460      pBar = 0

End Sub


Private Sub Frame2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

9470      pBar = 0

End Sub

Private Sub fraMicroscopy_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

9480      pBar = 0

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

9490      pBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

9500      pBar = 0

End Sub

Private Sub fraSampleID_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

9510      pBar = 0

9520      If cmdSaveMicro.Enabled Then
9530          MoveCursorToSaveButton
9540      End If

End Sub

Private Sub Frame7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

9550      pBar = 0

End Sub

Private Sub fraUrineSpecific_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

9560      pBar = 0

End Sub

Private Sub irelevant_Click(Index As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim strDept As String
          Dim strDirection As String

9570      On Error GoTo irelevant_Click_Error

9580      strDept = "Urine"

          '30    If Not GetSampleIDWithOffset Then Exit Sub
9590      If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
9600          If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
9610              SaveDemographics
9620              cmdSaveDemographics.Enabled = False
9630              cmdSaveInc.Enabled = False
9640          End If
9650      End If

9660      strDirection = IIf(Index = 0, "<", ">")

9670      Select Case SSTab1.Tab
              Case Mic.Demographics:
                  '+++ Junaid 20-05-2024
                  '130       sql = "SELECT TOP 1 SampleID FROM Demographics WHERE " & _
                  '                "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
                  '                "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
9680              sql = "SELECT TOP 1 SampleID FROM Demographics WHERE " & _
                      "SampleID " & strDirection & " " & Val(txtSampleID.Text) & " " & _
                      "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
                  '--- Junaid
9690          Case Mic.Urine: strDept = "Urine"
                  '+++ Junaid 20-05-2024
                  '150       sql = "SELECT TOP 1 SampleID FROM UrineResults50 WHERE " & _
                  '                "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
                  '                "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
9700              sql = "SELECT TOP 1 SampleID FROM UrineResults50 WHERE " & _
                      "SampleID " & strDirection & " " & Val(txtSampleID.Text) & " " & _
                      "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
                  '--- Junaid
9710          Case Mic.IDENTCAVAN: strDept = "UrineIdent"

                  '170       sql = "SELECT TOP 1 SampleID FROM UrineIdent50 WHERE " & _
                  '                "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
                  '                "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
9720              sql = "SELECT TOP 1 SampleID FROM UrineIdent50 WHERE " & _
                      "SampleID " & strDirection & " " & Val(txtSampleID.Text) & " " & _
                      "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")

9730          Case Mic.Faeces: strDept = "Faeces"
                  '+++ Junaid 20-05-2024
                  '190       sql = "SELECT TOP 1 SampleID FROM FaecesResults50 WHERE " & _
                  '                "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
                  '                "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
9740              sql = "SELECT TOP 1 SampleID FROM FaecesResults50 WHERE " & _
                      "SampleID " & strDirection & " " & Val(txtSampleID.Text) & " " & _
                      "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
                  '--- Junaid
9750          Case Mic.CandS:
                  '+++ Junaid 20-05-2024
                  '210       sql = "SELECT TOP 1 SampleID FROM Isolates WHERE " & _
                  '                "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
                  '                "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
9760              sql = "SELECT TOP 1 SampleID FROM Isolates WHERE " & _
                      "SampleID " & strDirection & " " & Val(txtSampleID.Text) & " " & _
                      "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")
                  '--- Junaid
9770          Case Else
9780              If Index = 0 Then
9790                  If txtSampleID > 1 Then
9800                      txtSampleID = Val(txtSampleID - 1)
9810                  End If
9820              Else
9830                  txtSampleID = Val(txtSampleID) + 1
9840              End If
                  '300       If Not GetSampleIDWithOffset Then Exit Sub
9850              LoadAllDetails
9860              cmdSaveDemographics.Enabled = False
9870              cmdSaveInc.Enabled = False
9880              cmdSaveMicro.Enabled = False
9890              cmdSaveHold.Enabled = False
9900              Exit Sub

9910      End Select

9920      Set tb = New Recordset
9930      RecOpenClient 0, tb, sql
9940      If Not tb.EOF Then
              '410       txtSampleID = Val(tb!SampleID & "") - sysOptMicroOffset(0)
9950          txtSampleID = Val(tb!SampleID & "")
9960      End If

          '430   If Not GetSampleIDWithOffset Then Exit Sub
9970      LoadAllDetails

9980      cmdSaveDemographics.Enabled = False
9990      cmdSaveInc.Enabled = False
10000     cmdSaveMicro.Enabled = False
10010     cmdSaveHold.Enabled = False

10020     Exit Sub

irelevant_Click_Error:

          Dim strES As String
          Dim intEL As Integer

10030     intEL = Erl
10040     strES = Err.Description
10050     LogError "frmEditMicrobiology", "irelevant_Click", intEL, strES, sql

End Sub

Private Sub iRunDate_Click(Index As Integer)

10060     If Index = 0 Then
10070         dtRunDate = DateAdd("d", -1, dtRunDate)
10080     Else
10090         If DateDiff("d", dtRunDate, Now) > 0 Then
10100             dtRunDate = DateAdd("d", 1, dtRunDate)
10110         End If
10120     End If

10130     cmdSaveInc.Enabled = True
10140     cmdSaveDemographics.Enabled = True

End Sub

Private Sub iSampleDate_Click(Index As Integer)

10150     If Index = 0 Then
10160         dtSampleDate = DateAdd("d", -1, dtSampleDate)
10170     Else
10180         If DateDiff("d", dtSampleDate, Now) > 0 Then
10190             dtSampleDate = DateAdd("d", 1, dtSampleDate)
10200         End If
10210     End If

10220     cmdSaveInc.Enabled = True
10230     cmdSaveDemographics.Enabled = True

End Sub


Private Sub iToday_Click(Index As Integer)

10240     If Index = 0 Then
10250         dtRunDate = Format$(Now, "dd/mm/yyyy")
10260     ElseIf Index = 1 Then
10270         If DateDiff("d", dtRunDate, Now) > 0 Then
10280             dtSampleDate = dtRunDate
10290         Else
10300             dtSampleDate = Format$(Now, "dd/mm/yyyy")
10310         End If
10320     Else
10330         dtRecDate = Format$(Now, "dd/mm/yyyy")
10340     End If

10350     cmdSaveInc.Enabled = True
10360     cmdSaveDemographics.Enabled = True

End Sub


Private Sub lblToxinB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10370     With lblToxinB
10380         Select Case .Caption
                  Case ""
10390                 .Caption = "Not Detected"
10400                 .BackColor = vbGreen
10410             Case "Not Detected"
10420                 .Caption = "Positive"
10430                 .BackColor = vbRed
10440             Case "Positive"
10450                 .Caption = "Inconclusive"
10460                 .BackColor = vbYellow
10470             Case "Inconclusive"
10480                 .Caption = "Rejected"
10490                 .BackColor = &H8000000F
10500             Case "Rejected"
10510                 .Caption = ""
10520                 .BackColor = &H8000000F
10530         End Select
10540         lblToxinA.Caption = .Caption
10550         lblToxinA.BackColor = .BackColor
10560     End With

10570     ShowUnlock Mic.CDiff

End Sub


Private Sub lstABsInUse_Click()

          Dim n As Integer
          Dim CURS As New CurrentAntibiotics
          '+++ Junaid 20-05-2024
          '10    CURS.Delete SampleIDWithOffset, lstABsInUse.List(lstABsInUse.ListIndex)
10580     CURS.Delete Trim(txtSampleID.Text), lstABsInUse.List(lstABsInUse.ListIndex)
          '--- Junaid

10590     lstABsInUse.RemoveItem lstABsInUse.ListIndex

10600     lblABsInUse = ""
10610     For n = 0 To lstABsInUse.ListCount - 1
10620         lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
10630     Next

10640     cmdSaveDemographics.Enabled = True
10650     cmdSaveInc.Enabled = True

End Sub

Private Sub mnuBacteria_Click()

10660     With frmListsGeneric
10670         .ListType = "BB"
10680         .ListTypeName = "Bacteria Entry"
10690         .ListTypeNames = "Bacteria Entries"
10700         .Show 1
10710     End With

10720     LoadListBacteria

End Sub

Private Sub mnuCSFAppearance_Click()

10730     With frmListsGeneric
10740         .ListType = "FA"
10750         .ListTypeName = "CSF Appearance"
10760         .ListTypeNames = "CSF Appearances"
10770         .Show 1
10780     End With

10790     FillListCSFAppearance

End Sub

Private Sub mnuCSFGram_Click()

10800     With frmListsGeneric
10810         .ListType = "FG"
10820         .ListTypeName = "Gram Stain"
10830         .ListTypeNames = "Gram Stains"
10840         .Show 1
10850     End With

10860     FillListCSFGram

End Sub

Private Sub mnuDefaultsMicro_Click()

10870     frmAntibioticLists.Show 1

End Sub

Private Sub mnuExit_Click()

10880     Me.Hide

End Sub

Private Sub mnuListPrestonCCDA_Click()
10890     With frmListsFaeces
10900         .o(2) = True
10910         .Show 1
10920     End With
End Sub

Private Sub mnuListSMAC_Click()
10930     With frmListsFaeces
10940         .o(1) = True
10950         .Show 1
10960     End With
End Sub

Private Sub mnuListXLDDCA_Click()
10970     With frmListsFaeces
10980         .o(0) = True
10990         .Show 1
11000     End With
End Sub

Private Sub mnuMicroGWQuantity_Click()
11010     With frmMicroLists
11020         .o(10).Value = True
11030         .Show 1
11040     End With
End Sub

Private Sub mnuMicroIDGram_Click()
11050     With frmMicroLists
11060         .o(3).Value = True
11070         .Show 1
11080     End With
End Sub

Private Sub mnuMicroWetPrep_Click()
11090     With frmMicroLists
11100         .o(4).Value = True
11110         .Show 1
11120     End With
End Sub

Private Sub mnuOrganisms_Click()

11130     With frmListsGeneric
11140         .ListType = "IN"
11150         .ListTypeName = "Organism"
11160         .ListTypeNames = "Organisms"
11170         .Show 1
11180     End With

11190     LoadListOrganism

End Sub

Private Sub mnuRCC_Click()

11200     With frmListsGeneric
11210         .ListType = "RR"
11220         .ListTypeName = "RCC Entry"
11230         .ListTypeNames = "RCC Entries"
11240         .Show 1
11250     End With

11260     LoadListRCC

End Sub

Private Sub mnuUrineCasts_Click()
11270     With frmMicroLists
11280         .o(5).Value = True
11290         .Show 1
11300     End With
End Sub

Private Sub mnuUrineCrystals_Click()
11310     With frmMicroLists
11320         .o(6).Value = True
11330         .Show 1
11340     End With
End Sub

Private Sub mnuUrineMisc_Click()
11350     With frmMicroLists
11360         .o(7).Value = True
11370         .Show 1
11380     End With
End Sub

Private Sub mnuWCC_Click()

11390     With frmListsGeneric
11400         .ListType = "WW"
11410         .ListTypeName = "WCC Entry"
11420         .ListTypeNames = "WCC Entries"
11430         .Show 1
11440     End With

11450     LoadListWCC

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

          Dim GenResults As New GenericResults
          Dim Fxs As New FaecesResults

11460     If LoadingAllDetails Then Exit Sub

11470     ShowMenuLists

11480     GenResults.Load Format$(Val(txtSampleID))
11490     Fxs.Load Format$(Val(txtSampleID))

11500     Select Case PreviousTab

              Case Mic.Demographics:
11510             If cmdSaveDemographics.Enabled Then
11520                 cmdSaveDemographics_Click
11530             End If
11540             cmdSaveHold.Enabled = False

11550         Case Mic.Urine: If cmdSaveMicro.Enabled Then SaveUrine False

11560         Case Mic.Identification:

11570         Case Mic.Faeces: If cmdSaveMicro.Enabled Then SaveFaeces

11580         Case Mic.CandS: If cmdSaveMicro.Enabled Then SaveIsolates: SaveSensitivities gNO

11590         Case Mic.FOB: If cmdSaveMicro.Enabled Then SaveFOB gDONTCARE

11600         Case Mic.RotaAdeno: If cmdSaveMicro.Enabled Then SaveRotaAdeno gDONTCARE

11610         Case Mic.CDiff: If cmdSaveMicro.Enabled Then SaveCdiff gDONTCARE

11620         Case Mic.CSF: If cmdSaveMicro.Enabled Then SaveCSF

11630         Case Mic.OP: If cmdSaveMicro.Enabled Then SaveOP gDONTCARE

11640         Case Mic.IDENTCAVAN:
                  '200       If cmdSaveMicro.Enabled Then SaveIdentification
11650             If cmbSite = "Urine" Or cmbSite = "Faeces" Then
11660                 cmdGramPrep.Visible = False
11670             Else
11680                 cmdGramPrep.Visible = True
11690             End If

11700         Case Mic.RSV: If cmdSaveMicro.Enabled Then SaveRSV gDONTCARE

11710         Case Mic.RedSub: If cmdSaveMicro.Enabled Then SaveRedSub gDONTCARE

11720     End Select

          '290   If Not GetSampleIDWithOffset Then Exit Sub

11730     cmdValidateMicro.Visible = False

11740     Select Case SSTab1.Tab
              Case Mic.Demographics:    'Demographics

11750         Case Mic.Urine:    'Urine
11760             cmdValidateMicro.Visible = True
11770             If LoadUrine() > 0 Then
11780                 If UrineResultsPresent() Then
11790                     SSTab1.TabCaption(Mic.Urine) = "<<Urine>>"
11800                 End If
11810             End If

                  'NOT USED IN CAVAN
11820         Case Mic.Identification:    'Identification

11830         Case Mic.Faeces:    'Faeces
11840             If Not FaecesLoaded Then
11850                 LoadFaeces
11860                 FaecesLoaded = True
11870             End If

11880         Case Mic.CandS:    'Sensitivities
11890             cmdValidateMicro.Visible = True
11900             If Not CSLoaded Then
11910                 LoadSensitivities
11920                 CSLoaded = True
11930             End If

11940         Case Mic.FOB:
11950             cmdValidateMicro.Visible = True
11960             If Not FOBLoaded Then
11970                 LoadFOB Fxs
11980                 FOBLoaded = True
11990             End If

12000         Case Mic.RotaAdeno:
12010             cmdValidateMicro.Visible = True
12020             If Not RotaAdenoLoaded Then
12030                 LoadRotaAdeno Fxs
12040                 RotaAdenoLoaded = True
12050             End If

12060         Case Mic.CDiff:
12070             cmdValidateMicro.Visible = True
12080             If Not CdiffLoaded Then
12090                 LoadCDiff GenResults, Fxs
12100                 CdiffLoaded = True
12110             End If

12120         Case Mic.OP:
12130             cmdValidateMicro.Visible = True
12140             If Not OPLoaded Then
12150                 LoadOP Fxs
12160                 OPLoaded = True
12170             End If

12180         Case Mic.IDENTCAVAN:
12190             If Not IdentificationLoaded Then
12200                 LoadIdentification
12210                 IdentificationLoaded = True
12220             End If

12230         Case Mic.RSV:
12240             cmdValidateMicro.Visible = True
12250             LoadRSV GenResults

12260         Case Mic.RedSub:
12270             cmdValidateMicro.Visible = True
12280             LoadRedSub GenResults

12290         Case Mic.CSF:
12300             cmdValidateMicro.Visible = True
12310             LoadCSF

12320     End Select

12330     CheckPrintValidLog "MICRO"
12340     VisibilityofCmdBtn

          '890   cmdSaveMicro.Enabled = ForceSaveability
          '900   cmdSaveHold.Enabled = ForceSaveability
12350     cmdSaveMicro.Enabled = False
12360     cmdSaveHold.Enabled = False

End Sub

Private Sub CheckPrintValidLog(ByVal Dept As String)

          Dim tb As Recordset
          Dim sql As String
          Dim LogDept As String

          'A Rota/Adeno
          'B Biochemistry
          'C Coagulation
          'D Culture/Sensitivity
          'F FOB
          'G cDiff
          'H Haematology
          'I Immunology
          'O Ova/Parasites
          'R Reducing Substances
          'S ESR
          'U Urine
          'V RSV
          'X External

12370     On Error GoTo CheckPrintValidLog_Error

12380     Select Case UCase$(Dept)
              Case "MICRO": LogDept = "M"
                  '  Case "REDSUB":    LogDept = "R"
                  '  Case "RSV":       LogDept = "V"
                  '  Case "OP":        LogDept = "O"
                  '  Case "CDIFF":     LogDept = "G"
                  '  Case "FOB":       LogDept = "F"
                  '  Case "URINE":     LogDept = "U"
                  '  Case "CANDS":     LogDept = "D"
12390     End Select

12400     cmdValidateMicro.Caption = "&Validate"
12410     cmdValidateMicro.BackColor = vbRed

          '60        Sql = "SELECT Valid FROM PrintValidLog WHERE " & _
          '                "SampleID = '" & txtSampleID + sysOptMicroOffset(0) & "' " & _
          '                "AND Department = '" & LogDept & "' " & _
          '                "AND Valid = 1"
12420     If txtSampleID.Text = "" Then
12430         Exit Sub
12440     End If
12450     sql = "SELECT IsNULL(Valid,0) Valid FROM PrintValidLog WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "AND Department = '" & LogDept & "' " & _
              "AND IsNULL(Valid,0) = 1"
12460     Set tb = New Recordset
12470     RecOpenServer 0, tb, sql
12480     If Not tb Is Nothing Then
12490         If Not tb.EOF Then
12500             cmdValidateMicro.Caption = "Un&Validate"
12510             cmdValidateMicro.BackColor = vbGreen
          
12520         End If
12530     End If

12540     Exit Sub

CheckPrintValidLog_Error:

          Dim strES As String
          Dim intEL As Integer

12550     intEL = Erl
12560     strES = Err.Description
12570     LogError "frmEditMicrobiology", "CheckPrintValidLog", intEL, strES, sql

End Sub
Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

12580     pBar = 0

End Sub


Private Sub Timer1_Timer()
12590     m_Time = m_Time + 5
12600     If m_Time = 1800 Then
12610         Call CheckUrineResults
12620         m_Time = 0
12630     End If
End Sub

Private Sub tRecTime_GotFocus()

12640     tRecTime.SelStart = 0
12650     tRecTime.SelLength = 0

End Sub


Private Sub tRecTime_KeyPress(KeyAscii As Integer)

12660     pBar = 0

12670     cmdSaveDemographics.Enabled = True
12680     cmdSaveInc.Enabled = True

End Sub


Private Sub tRecTime_LostFocus()
12690     SetDatesColour
End Sub

Private Sub tSampleTime_LostFocus()
12700     SetDatesColour
End Sub

Private Sub txtaddress_Change(Index As Integer)

12710     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub txtaddress_KeyPress(Index As Integer, KeyAscii As Integer)

12720     cmdSaveDemographics.Enabled = True
12730     cmdSaveInc.Enabled = True

End Sub


Private Sub txtaddress_LostFocus(Index As Integer)

12740     txtAddress(Index) = Initial2Upper(txtAddress(Index))

End Sub


Private Sub txtAdeno_KeyPress(KeyAscii As Integer)

12750     KeyAscii = 0

End Sub

Private Sub txtAdeno_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

12760     With txtAdeno
12770         Select Case .Text
                  Case ""
12780                 .Text = "Negative"
12790                 .BackColor = vbGreen
12800             Case "Negative"
12810                 .Text = "Positive"
12820                 .BackColor = vbRed
12830             Case "Positive"
12840                 .Text = ""
12850                 .BackColor = &H8000000F
12860         End Select
12870     End With

12880     ShowUnlock Mic.RotaAdeno

End Sub


Private Sub txtage_Change()

12890     lblAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

12900     cmdSaveDemographics.Enabled = True
12910     cmdSaveInc.Enabled = True

End Sub


Private Sub txtBacteria_Change()

12920     If Not ClearingUrine Then
              'If chkPregnant.Value = 0 Then CheckUrineAutoVal
12930     End If

End Sub

Private Sub txtBacteria_Click()

          Dim n As Integer
          Dim x As Integer

12940     If UBound(ListBacteria) = 0 Then
12950         mnuBacteria_Click
12960     End If

12970     For n = 0 To UBound(ListBacteria)
12980         If txtBacteria = ListBacteria(n) Then
12990             If n = UBound(ListBacteria) Then
13000                 x = 0
13010             Else
13020                 x = n + 1
13030             End If
13040             txtBacteria = ListBacteria(x)
13050             Exit For
13060         End If
13070     Next

13080     ShowUnlock Mic.Urine

End Sub

Private Sub txtBacteria_KeyUp(KeyCode As Integer, Shift As Integer)

13090     ShowUnlock Mic.Urine

End Sub

Private Sub txtBacteria_LostFocus()

          Dim sql As String
          Dim tb As Recordset

13100     On Error GoTo txtBacteria_LostFocus_Error

13110     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'BB' " & _
              "AND Code = '" & AddTicks(txtBacteria) & "'"
13120     Set tb = New Recordset
13130     RecOpenServer 0, tb, sql
13140     If Not tb.EOF Then
13150         txtBacteria = tb!Text & ""
13160     End If

13170     Exit Sub

txtBacteria_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

13180     intEL = Erl
13190     strES = Err.Description
13200     LogError "frmEditMicrobiology", "txtBacteria_LostFocus", intEL, strES, sql

End Sub


Private Sub txtBenceJones_KeyUp(KeyCode As Integer, Shift As Integer)

13210     ShowUnlock Mic.Urine

End Sub


Private Sub txtchart_Change()

13220     lblChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

13230     cmdSaveDemographics.Enabled = True
13240     cmdSaveInc.Enabled = True

End Sub


Private Sub txtchart_LostFocus()

13250     If Trim$(txtChart) = "" Then Exit Sub
13260     If Trim$(txtSurName) <> "" Then Exit Sub

13270     LoadPatientFromChart Me, mNewRecord
13280     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
13290         If MatchingDemoLoaded = False Then LoadMatchingDemo
13300     End If

End Sub


Private Sub txtClinDetails_KeyPress(KeyAscii As Integer)

13310     cmdSaveDemographics.Enabled = True
13320     cmdSaveInc.Enabled = True

End Sub





Private Sub txtClinDetails_LostFocus()

13330     CheckLorem txtClinDetails

End Sub


Private Sub txtConC_GotFocus()

13340     If txtConC.Text = "Consultant Comments" Then
13350         txtConC.Text = ""
13360     End If

End Sub


Private Sub txtConC_KeyUp(KeyCode As Integer, Shift As Integer)

13370     cmdSaveMicro.Enabled = True
13380     cmdSaveHold.Enabled = True

End Sub


Private Sub txtConC_LostFocus()

13390     If Trim$(txtConC) = "" Then
13400         txtConC = "Consultant Comments"
13410     End If

13420     CheckLorem txtConC

End Sub


Private Sub txtCSFRCC_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

13430     cmdSaveMicro.Enabled = True
13440     cmdSaveHold.Enabled = True

13450     ShowUnlock Mic.CSF

End Sub

Private Sub txtCSFWCC_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

13460     cmdSaveMicro.Enabled = True
13470     cmdSaveHold.Enabled = True

13480     ShowUnlock Mic.CSF

End Sub

Private Sub txtCSFWCCDiff_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

13490     cmdSaveMicro.Enabled = True
13500     cmdSaveHold.Enabled = True

13510     ShowUnlock Mic.CSF

End Sub

Private Sub txtDemographicComment_LostFocus()

13520     CheckLorem txtDemographicComment

End Sub

Private Sub txtDoB_Change()

13530     lblDoB = txtDoB
13540     LabNoUpdatePrviousData = ""
13550     LabNoUpdatePrvColor

End Sub




Private Sub txtDoB_KeyPress(KeyAscii As Integer)

13560     cmdSaveDemographics.Enabled = True
13570     cmdSaveInc.Enabled = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtDoB_LostFocus
' Author    : Masood
' Date      : 02/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtDoB_LostFocus()

13580     On Error GoTo txtDoB_LostFocus_Error


13590     txtDoB = Convert62Date(txtDoB, BACKWARD)
13600     txtAge = CalcAge(txtDoB, dtSampleDate)
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
13610     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
13620         If MatchingDemoLoaded = False Then LoadMatchingDemo
13630     End If

       
13640     Exit Sub

       
txtDoB_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

13650     intEL = Erl
13660     strES = Err.Description
13670     LogError "frmEditMicrobiology", "txtDoB_LostFocus", intEL, strES

End Sub


Private Sub TimerBar_Timer()

13680     pBar = pBar + 1

13690     If pBar = pBar.max Then
             'Zyam 15-06-24
             'ChkIfAnyChildFormShown
             'Zyam 15-06-24
13700         Me.Hide
13710         Exit Sub
13720     End If

End Sub
'Zyam 15-06-24
'Private Sub ChkIfAnyChildFormShown()
'Dim frm As Form
'
'For Each frm In Forms
'    If frm.Name <> Me.Name And frm.Name <> "frmMain" Then
'        frm.Hide
'    End If
'Next frm
'End Sub

'Zyam 15-06-24


'---------------------------------------------------------------------------------------
' Procedure : txtExtSampleID_LostFocus
' Author    : Masood
' Date      : 21/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtExtSampleID_LostFocus()

          '10
13730     On Error GoTo txtExtSampleID_LostFocus_Error


13740     If txtExtSampleID <> "" Then
13750         If cmbSite = "" Then
13760             iMsg "Please select site first", vbExclamation
13770             Exit Sub
13780         End If
13790         With frmGporders
13800             .SampleID = txtSampleID
13810             .SampleIDExt = txtExtSampleID
13820             .MicroSite = cmbSite
13830             .ClinicalDetails = ""
13840             .DisiplinesQuery = " AND P.Department IN ('Microbiology')"
13850             Set .EditScreen = Me
13860             .Show 1
13870         End With
13880     End If
13890     Exit Sub
13900     Call LoadPatientFromOrderCom(Me, False, txtExtSampleID)


13910     Exit Sub


txtExtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

13920     intEL = Erl
13930     strES = Err.Description
13940     LogError "frmEditMicrobiology", "txtExtSampleID_LostFocus", intEL, strES
End Sub

Private Sub txtFatGlobules_KeyPress(KeyAscii As Integer)

13950     ShowUnlock Mic.Urine

End Sub


Private Sub txtForeName_KeyPress(KeyAscii As Integer)

13960     If SSTab1.Tab <> 0 Then
13970         SSTab1.Tab = 0
13980     End If

13990     cmdSaveDemographics.Enabled = True
14000     cmdSaveInc.Enabled = True

End Sub


Private Sub txtForeName_LostFocus()

14010     On Error GoTo txtForeName_LostFocus_Error

          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
14020     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
14030         If MatchingDemoLoaded = False Then LoadMatchingDemo
14040     End If

14050     Exit Sub
txtForeName_LostFocus_Error:
         
14060     LogError "frmEditMicrobiology", "txtForeName_LostFocus", Erl, Err.Description



End Sub

Private Sub txtGlucose_KeyPress(KeyAscii As Integer)

14070     ShowUnlock Mic.Urine

End Sub

Private Sub txtHCGLevel_KeyPress(KeyAscii As Integer)

14080     ShowUnlock Mic.Urine

End Sub


Private Sub txtIdentification_KeyPress(Index As Integer, KeyAscii As Integer)

14090     cmdSaveMicro.Enabled = True
14100     cmdSaveHold.Enabled = True

End Sub


Private Sub txtMSC_GotFocus()

14110     If txtMSC.Text = "Medical Scientist Comments" Then
14120         txtMSC.Text = ""
14130     End If

End Sub


Private Sub txtMSC_KeyUp(KeyCode As Integer, Shift As Integer)

14140     cmdSaveMicro.Enabled = True
14150     cmdSaveHold.Enabled = True

End Sub

Private Sub txtMSC_LostFocus()

14160     If Trim$(txtMSC) = "" Then
14170         txtMSC = "Medical Scientist Comments"
14180     End If

14190     CheckLorem txtMSC

End Sub


Private Sub txtSampleID_Change()
    '    MsgBox txtSampleID.Text
End Sub

Private Sub txtsampleid_KeyPress(KeyAscii As Integer)
14200     If KeyAscii = 13 Then
14210         Call txtsampleid_LostFocus
14220         txtDemographicComment.SetFocus
14230     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSex_LostFocus
' Author    : Masood
' Date      : 02/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSex_LostFocus()
14240     On Error GoTo txtSex_LostFocus_Error


          '20        Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)

       
14250     Exit Sub

       
txtSex_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

14260     intEL = Erl
14270     strES = Err.Description
14280     LogError "frmEditMicrobiology", "txtSex_LostFocus", intEL, strES

End Sub

Private Sub txtSurName_Change()

14290     lblName = Trim$(txtSurName & " " & txtForeName)
14300     LabNoUpdatePrviousData = ""
14310     LabNoUpdatePrvColor

End Sub

Private Sub txtForeName_Change()

14320     lblName = Trim$(txtSurName & " " & txtForeName)
14330     LabNoUpdatePrviousData = ""
14340     LabNoUpdatePrvColor

End Sub


Private Sub txtSurName_KeyPress(KeyAscii As Integer)

14350     If SSTab1.Tab <> 0 Then
14360         SSTab1.Tab = 0
14370     End If

14380     cmdSaveDemographics.Enabled = True
14390     cmdSaveInc.Enabled = True

End Sub

Private Sub txtSurname_LostFocus()

          Dim strSurName As String
          Dim strForeName As String
          Dim strSex     As String


14400     On Error GoTo txtSurname_LostFocus_Error

14410     strSurName = txtSurName
14420     strForeName = txtForeName
14430     strSex = txtSex

14440     NameLostFocus strSurName, strForeName, strSex

14450     txtSurName = strSurName
14460     txtForeName = strForeName
14470     txtSex = strSex
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
14480     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
14490         If MatchingDemoLoaded = False Then LoadMatchingDemo
14500     End If

14510     Exit Sub
txtSurname_LostFocus_Error:
         
14520     LogError "frmEditMicrobiology", "txtSurname_LostFocus", Erl, Err.Description

End Sub


Private Sub txtPregnancy_Click()
          Dim TempComment As String
14530     TempComment = Trim(txtUrineComment)
14540     TempComment = Replace(TempComment, "Please repeat specimen in 24-48 hours.", "")
14550     TempComment = Replace(TempComment, "Specimen Unsuitable - Please repeat.", "")
14560     Select Case Left$(txtPregnancy & " ", 1)
              Case " ":
14570             txtPregnancy = "Negative"
14580             txtHCGLevel = "<20"
14590             txtUrineComment = TempComment
14600         Case "N":
14610             txtPregnancy = "Positive"
14620             txtHCGLevel = ">=20"
14630             txtUrineComment = TempComment
14640         Case "P":
14650             txtPregnancy = "Equivocal"
14660             txtHCGLevel = ""
14670             txtUrineComment = TempComment & IIf(Len(TempComment) > 0, " ", "") & "Please repeat specimen in 24-48 hours."
14680         Case "E":
14690             txtPregnancy = "Specimen Unsuitable"
14700             txtHCGLevel = ""
14710             txtUrineComment = TempComment & IIf(Len(TempComment) > 0, " ", "") & "Specimen Unsuitable - Please repeat."
14720         Case "S":
14730             txtPregnancy = ""
14740             txtHCGLevel = ""
14750             txtUrineComment = TempComment
14760     End Select

14770     ShowUnlock Mic.Urine

End Sub


Private Sub txtBilirubin_Click()

14780     Select Case txtBilirubin
              Case "": txtBilirubin = "+"
14790         Case "+": txtBilirubin = "++"
14800         Case "++": txtBilirubin = "+++"
14810         Case "+++": txtBilirubin = "++++"
14820         Case "++++": txtBilirubin = "Nil"
14830         Case "Nil": txtBilirubin = ""
14840     End Select

14850     ShowUnlock Mic.Urine

End Sub


Private Sub txtBilirubin_KeyPress(KeyAscii As Integer)

14860     Select Case KeyAscii
              Case vbKey0, vbKeyNumpad0, vbKeyN: txtBilirubin = "Nil"
14870         Case vbKey1, vbKeyNumpad1: txtBilirubin = "+"
14880         Case vbKey2, vbKeyNumpad2: txtBilirubin = "++"
14890         Case vbKey3, vbKeyNumpad3: txtBilirubin = "+++"
14900         Case vbKey4, vbKeyNumpad4: txtBilirubin = "++++"
14910         Case Else: txtBilirubin = ""
14920     End Select
14930     KeyAscii = 0

14940     ShowUnlock Mic.Urine

End Sub


Private Sub txtBloodHb_Click()

14950     Select Case txtBloodHb
              Case "": txtBloodHb = "+"
14960         Case "+": txtBloodHb = "++"
14970         Case "++": txtBloodHb = "+++"
14980         Case "+++": txtBloodHb = "++++"
14990         Case "++++": txtBloodHb = "Nil"
15000         Case "Nil": txtBloodHb = ""
15010     End Select

15020     ShowUnlock Mic.Urine

End Sub


Private Sub txtBloodHb_KeyPress(KeyAscii As Integer)

15030     Select Case KeyAscii
              Case vbKey0, vbKeyNumpad0, vbKeyN: txtBloodHb = "Nil"
15040         Case vbKey1, vbKeyNumpad1: txtBloodHb = "+"
15050         Case vbKey2, vbKeyNumpad2: txtBloodHb = "++"
15060         Case vbKey3, vbKeyNumpad3: txtBloodHb = "+++"
15070         Case vbKey4, vbKeyNumpad4: txtBloodHb = "++++"
15080         Case Else: txtBloodHb = ""
15090     End Select
15100     KeyAscii = 0

15110     ShowUnlock Mic.Urine

End Sub


Private Sub txtGlucose_Click()

15120     If txtGlucose = "" Then
15130         txtGlucose = "Pos"
15140     Else
15150         txtGlucose.SelStart = 0
15160         txtGlucose.SelLength = Len(txtGlucose)
15170     End If
15180     ShowUnlock Mic.Urine

End Sub


Private Sub txtKetones_Click()

15190     Select Case txtKetones
              Case "": txtKetones = "+"
15200         Case "+": txtKetones = "++"
15210         Case "++": txtKetones = "+++"
15220         Case "+++": txtKetones = "++++"
15230         Case "++++": txtKetones = "Nil"
15240         Case "Nil": txtKetones = ""
15250     End Select

15260     ShowUnlock Mic.Urine

End Sub


Private Sub txtKetones_KeyPress(KeyAscii As Integer)

15270     Select Case KeyAscii
              Case vbKey0, vbKeyNumpad0, vbKeyN: txtKetones = "Nil"
15280         Case vbKey1, vbKeyNumpad1: txtKetones = "+"
15290         Case vbKey2, vbKeyNumpad2: txtKetones = "++"
15300         Case vbKey3, vbKeyNumpad3: txtKetones = "+++"
15310         Case vbKey4, vbKeyNumpad4: txtKetones = "++++"
15320         Case Else: txtKetones = ""
15330     End Select
15340     KeyAscii = 0

15350     ShowUnlock Mic.Urine

End Sub


Private Sub txtpH_Click()

15360     Select Case txtpH
              Case "": txtpH = "Acid"
15370         Case "Acid": txtpH = "Alkaline"
15380         Case "Alkaline": txtpH = "Neutral"

15390         Case "Neutral":
15400             If iMsg("Is Sample Unsuitable?", vbQuestion + vbYesNo) = vbYes Then
15410                 txtpH = "Unsuitable"
15420                 txtProtein = ""
15430                 txtGlucose = ""
15440                 txtKetones = ""
15450                 txtUrobilinogen = ""
15460                 txtBilirubin = ""
15470             Else
15480                 txtpH = ""
15490             End If

15500         Case Else: txtpH = ""
15510     End Select

15520     ShowUnlock Mic.Urine

End Sub

Private Sub txtpH_KeyPress(KeyAscii As Integer)

15530     KeyAscii = 0

15540     Select Case txtpH
              Case "": txtpH = "Acid"
15550         Case "Acid": txtpH = "Alkaline"
15560         Case "Alkaline": txtpH = "Neutral"
15570         Case "Neutral": txtpH = "Unsuitable"
15580         Case Else: txtpH = ""
15590     End Select

End Sub


Private Sub txtProtein_Click()

15600     If txtProtein = "" Then
15610         txtProtein = "Pos"
15620     Else
15630         txtProtein.SelStart = 0
15640         txtProtein.SelLength = Len(txtProtein)
15650     End If

15660     ShowUnlock Mic.Urine

End Sub


Private Sub txtRCC_Click()

          Dim n As Integer
          Dim x As Integer

15670     If UBound(ListRCC) = 0 Then
15680         mnuRCC_Click
15690     End If

15700     For n = 0 To UBound(ListRCC)
15710         If txtRCC = ListRCC(n) Then
15720             If n = UBound(ListRCC) Then
15730                 x = 0
15740             Else
15750                 x = n + 1
15760             End If
15770             txtRCC = ListRCC(x)
15780             Exit For
15790         End If
15800     Next

15810     ShowUnlock Mic.Urine

End Sub

Private Sub txtRCC_KeyUp(KeyCode As Integer, Shift As Integer)

15820     ShowUnlock Mic.Urine

End Sub


Private Sub txtRCC_LostFocus()

          Dim sql As String
          Dim tb As Recordset

15830     On Error GoTo txtRCC_LostFocus_Error

15840     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'RR' " & _
              "AND Code = '" & AddTicks(txtRCC) & "'"
15850     Set tb = New Recordset
15860     RecOpenServer 0, tb, sql
15870     If Not tb.EOF Then
15880         txtRCC = tb!Text & ""
15890     End If

15900     Exit Sub

txtRCC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

15910     intEL = Erl
15920     strES = Err.Description
15930     LogError "frmEditMicrobiology", "txtRCC_LostFocus", intEL, strES, sql

End Sub


Private Sub txtRota_KeyPress(KeyAscii As Integer)

15940     KeyAscii = 0

End Sub

Private Sub txtRota_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

15950     With txtRota
15960         Select Case .Text
                  Case ""
15970                 .Text = "Negative"
15980                 .BackColor = vbGreen
15990             Case "Negative"
16000                 .Text = "Positive"
16010                 .BackColor = vbRed
16020             Case "Positive"
16030                 .Text = ""
16040                 .BackColor = &H8000000F
16050         End Select
16060     End With

16070     ShowUnlock Mic.RotaAdeno

End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtsampleid_GotFocus
' Author    : XPMUser
' Date      : 04/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtsampleid_GotFocus()

16080     On Error GoTo txtsampleid_GotFocus_Error


16090     If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
16100         If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
                  '40            If Not GetSampleIDWithOffset Then Exit Sub
16110             SaveDemographics
16120             cmdSaveDemographics.Enabled = False
16130             cmdSaveInc.Enabled = False
16140         End If
16150     End If
16160     ClearLabNoSelection

       
16170     Exit Sub

       
txtsampleid_GotFocus_Error:

          Dim strES As String
          Dim intEL As Integer

16180     intEL = Erl
16190     strES = Err.Description
16200     LogError "frmEditMicrobiology", "txtsampleid_GotFocus", intEL, strES
End Sub

Private Sub txtsampleid_LostFocus()

16210     If Not ComingFromUnlock Then

16220         txtSampleID = Format$(Val(txtSampleID))
16230         If txtSampleID = 0 Then Exit Sub
          
          
              '40        If Not GetSampleIDWithOffset Then Exit Sub
              '    txtLabNo = Val(FndMaxID("demographics", "LabNo", "")) + 1
16240         LoadAllDetails
            
16250         LabNoUpdatePrviousData = ""
16260         cmdSaveDemographics.Enabled = False
16270         cmdSaveInc.Enabled = False
              '70    cmdSaveMicro.Enabled = ForceSaveability
              '80    cmdSaveHold.Enabled = ForceSaveability
16280         cmdSaveMicro.Enabled = False
16290         cmdSaveHold.Enabled = False


16300     End If

16310     ComingFromUnlock = False

End Sub

'Abubaker +++ 09/10/2023

'Private Sub ReloadAndLogData()
'On Error GoTo ErrorHandler
'
'    Dim Sql As String
'    Dim tb As Recordset
'    Dim currentTime As String
'    Dim sampleIDValue As String
'    Dim lblNameValue As String
'
'    Sql = "Select SampleID from Demographics Where SampleID = '" & txtSampleID.Text & "'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, Sql
'
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            Call txtsampleid_LostFocus
'            currentTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
'            sampleIDValue = txtSampleID.Text
'            lblNameValue = lblName.Caption
'
'            Sql = "INSERT INTO BugLog (DateTime, SampleID, LblName) VALUES ('" & currentTime & "', '" & sampleIDValue & "', '" & lblNameValue & "');"
'            Cnxn(0).Execute Sql
'        End If
'    End If
'
'    Exit Sub
'ErrorHandler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "frmEditMicrobiology", "ReloadAndLogData", intEL, strES, Sql
'    '    MsgBox Err.Description
'End Sub

'Abubaker --- 09/10/2023

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

16320     cmdSaveDemographics.Enabled = True
16330     cmdSaveInc.Enabled = True

End Sub


Private Sub txtSampleID_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

16340     If cmdSaveMicro.Enabled Then
16350         MoveCursorToSaveButton
16360     End If

End Sub

Private Sub txtSex_Change()

16370     lblSex = txtSex

End Sub

Private Sub txtsex_Click()

16380     Select Case Trim$(txtSex)
              Case "": txtSex = "Male"
16390         Case "Male": txtSex = "Female"
16400         Case "Female": txtSex = ""
16410         Case Else: txtSex = ""
16420     End Select

16430     cmdSaveDemographics.Enabled = True
16440     cmdSaveInc.Enabled = True

End Sub


Private Sub txtsex_KeyPress(KeyAscii As Integer)

16450     KeyAscii = 0
16460     txtsex_Click

End Sub


Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

16470     cmdSaveDemographics.Enabled = True
16480     cmdSaveInc.Enabled = True

End Sub

Private Sub txtUrineComment_KeyPress(KeyAscii As Integer)

16490     ShowUnlock Mic.Urine

End Sub

Private Sub txtUrineComment_KeyUp(KeyCode As Integer, Shift As Integer)
16500     If Trim(txtBacteria.Text) = "" Or Trim(txtWCC.Text) = "" Or Trim(txtRCC.Text) = "" Then
16510         txtUrineComment.Text = ""
16520     End If
End Sub

Private Sub txtUrineComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String
          Dim S As Variant
          Dim n As Integer

16530     On Error GoTo txtUrineComment_LostFocus_Error

16540     If Trim$(txtUrineComment) = "" Then Exit Sub

16550     S = Split(txtUrineComment, " ")

16560     For n = 0 To UBound(S)
16570         sql = "Select * from Lists where " & _
                  "ListType = 'HA' " & _
                  "and Code = '" & S(n) & "' and InUse = 1"
16580         Set tb = New Recordset
16590         RecOpenServer 0, tb, sql
16600         If Not tb.EOF Then
16610             S(n) = tb!Text & ""
16620         End If
16630     Next

16640     txtUrineComment = Join(S, " ")

16650     CheckLorem txtUrineComment

16660     Exit Sub

txtUrineComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

16670     intEL = Erl
16680     strES = Err.Description
16690     LogError "frmEditMicrobiology", "txtUrineComment_LostFocus", intEL, strES, sql

End Sub


Private Sub txtSG_KeyPress(KeyAscii As Integer)

16700     ShowUnlock Mic.Urine

End Sub


Private Sub txtSiteDetails_Change()

16710     lblSiteDetails = cmbSite & " " & Trim(txtSiteDetails.Text)

End Sub

Private Sub txtSiteDetails_KeyPress(KeyAscii As Integer)

16720     cmdSaveDemographics.Enabled = True
16730     cmdSaveInc.Enabled = True

End Sub


Private Sub txtUrobilinogen_Click()

16740     Select Case txtUrobilinogen
              Case "": txtUrobilinogen = "+"
16750         Case "+": txtUrobilinogen = "++"
16760         Case "++": txtUrobilinogen = "+++"
16770         Case "+++": txtUrobilinogen = "++++"
16780         Case "++++": txtUrobilinogen = "Nil"
16790         Case "Nil": txtUrobilinogen = ""
16800     End Select

16810     ShowUnlock Mic.Urine

End Sub


Private Sub txtUrobilinogen_KeyPress(KeyAscii As Integer)

16820     Select Case KeyAscii
              Case vbKey0, vbKeyNumpad0, vbKeyN: txtUrobilinogen = "Nil"
16830         Case vbKey1, vbKeyNumpad1: txtUrobilinogen = "+"
16840         Case vbKey2, vbKeyNumpad2: txtUrobilinogen = "++"
16850         Case vbKey3, vbKeyNumpad3: txtUrobilinogen = "+++"
16860         Case vbKey4, vbKeyNumpad4: txtUrobilinogen = "++++"
16870         Case Else: txtUrobilinogen = ""
16880     End Select
16890     KeyAscii = 0

16900     ShowUnlock Mic.Urine

End Sub


Private Sub txtWCC_Change()

16910     If Not ClearingUrine Then
              'If chkPregnant.Value = 0 Then CheckUrineAutoVal
16920     End If

End Sub

Private Sub txtWCC_Click()

          Dim n As Integer
          Dim x As Integer

16930     If UBound(ListWCC) = 0 Then
16940         mnuWCC_Click
16950     End If

16960     For n = 0 To UBound(ListWCC)
16970         If txtWCC = ListWCC(n) Then
16980             If n = UBound(ListWCC) Then
16990                 x = 0
17000             Else
17010                 x = n + 1
17020             End If
17030             txtWCC = ListWCC(x)
17040             Exit For
17050         End If
17060     Next

17070     ShowUnlock Mic.Urine

End Sub

Private Sub txtWCC_KeyPress(KeyAscii As Integer)

17080     ShowUnlock Mic.Urine

End Sub


Private Sub txtWCC_LostFocus()

          Dim sql As String
          Dim tb As Recordset

17090     On Error GoTo txtWCC_LostFocus_Error

17100     sql = "SELECT Text FROM Lists WHERE " & _
              "ListType = 'WW' " & _
              "AND Code = '" & AddTicks(txtWCC) & "'"
17110     Set tb = New Recordset
17120     RecOpenServer 0, tb, sql
17130     If Not tb.EOF Then
17140         txtWCC = tb!Text & ""
17150     End If

17160     Exit Sub

txtWCC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

17170     intEL = Erl
17180     strES = Err.Description
17190     LogError "frmEditMicrobiology", "txtWCC_LostFocus", intEL, strES, sql

End Sub

Private Sub UpDown1_GotFocus()
    '
    'If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
    '  If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
    '    GetSampleIDWithOffset
    '    SaveDemographics
    '    cmdSaveDemographics.Enabled = False
    '    cmdSaveInc.Enabled = False
    '  End If
    'End If
    '
End Sub

Private Sub UpDown1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    'If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
    '  If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
    '    GetSampleIDWithOffset
    '    SaveDemographics
    '    cmdSaveDemographics.Enabled = False
    '    cmdSaveInc.Enabled = False
    '  End If
    'End If

End Sub


Private Sub UpDown1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

17200     If cmdSaveMicro.Enabled Then
17210         MoveCursorToSaveButton
17220     End If

End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

17230     pBar = 0

          '20    If Not GetSampleIDWithOffset Then Exit Sub
17240     m_Flag = False
17250     LoadAllDetails

17260     cmdSaveDemographics.Enabled = False
17270     cmdSaveInc.Enabled = False
17280     cmdSaveMicro.Enabled = False
17290     cmdSaveHold.Enabled = False
17300     ClearLabNoSelection
End Sub



Public Property Let FromViewReportSID(ByVal strNewValue As String)

17310     pFromViewReportSID = strNewValue

End Property
Public Property Let PrintToPrinter(ByVal strNewValue As String)

17320     pPrintToPrinter = strNewValue

End Property

Public Property Get PrintToPrinter() As String

17330     PrintToPrinter = pPrintToPrinter

End Property

Private Sub udHistoricalFaecesView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

17340     FillHistoricalFaeces

End Sub

Private Sub SetDatesColour()

17350     On Error GoTo SetDatesColour_Error

17360     If CheckDateSequence(dtSampleDate, dtRecDate, dtRunDate, tSampleTime, tRecTime) Then
17370         Frame7.ForeColor = vbButtonText
17380         Frame7.Font.Bold = False
17390         label1(15).ForeColor = vbButtonText
17400         label1(15).Font.Bold = False
17410         label1(16).ForeColor = vbButtonText
17420         label1(16).Font.Bold = False
17430         lblDateError.Visible = False
17440     Else
17450         Frame7.ForeColor = vbRed
17460         Frame7.Font.Bold = True
17470         label1(15).ForeColor = vbRed
17480         label1(15).Font.Bold = True
17490         label1(16).ForeColor = vbRed
17500         label1(16).Font.Bold = True
17510         lblDateError.Visible = True
17520     End If

17530     Exit Sub

SetDatesColour_Error:

          Dim strES As String
          Dim intEL As Integer

17540     intEL = Erl
17550     strES = Err.Description
17560     LogError "basShared", "SetDatesColour", intEL, strES

End Sub
Private Sub ShowMenuLists()

17570     mnuLists.Visible = False



17580     Select Case SSTab1.Tab
              Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12:
17590             mnuLists.Visible = UserHasAuthority(UserMemberOf, "MicroLists")

17600     End Select


End Sub




Private Sub LabNoUpdatePrvColor()
17610     On Error GoTo LabNoUpdatePrv_Error


          'If UCase(LabNoUpdatePrviousData) = UCase(txtSurName & txtForeName & txtDoB) Then
17620     If LabNoUpdatePrviousData = "1" Then
17630         txtLabNo.BackColor = vbGreen
              '40        lAddWardGP = FindLatestAddress(txtChart, Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, Left$(txtSex, 1), txtLabNo)

17640     Else
17650         txtLabNo.BackColor = vbRed
17660     End If


17670     Exit Sub


LabNoUpdatePrv_Error:

          Dim strES As String
          Dim intEL As Integer

17680     intEL = Erl
17690     strES = Err.Description
17700     LogError "frmEditAll", "LabNoUpdatePrv", intEL, strES
End Sub



Private Sub LabNoUpdatePrvData(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String)
17710     On Error GoTo LabNoUpdatePrvData_Error
          Dim sql As String

17720     If txtMultiSeltdDemoForLabNoUpd <> "" Then
17730         sql = txtMultiSeltdDemoForLabNoUpd
17740         Cnxn(0).Execute (sql)
17750     End If

          'If txtLabNo.BackColor = vbGreen Then
          '    sql = "UPDATE demographics "
          '    sql = sql & " SET PatName ='" & Name & "'"
          '    sql = sql & ",  DoB ='" & DoB & "'"
          '    sql = sql & ",  Sex ='" & Sex & "'"
          '    sql = sql & ",  LabNo ='" & LabNo & "'"
          '    sql = sql & ",  Chart ='" & ChartNo & "'"
          '    sql = sql & " WHERE "
          '    sql = sql & " UPPER(PatName) ='" & UCase(Name) & "'"
          '    sql = sql & " AND DoB ='" & DoB & "'"
          '    sql = sql & " AND UPPER(Sex) ='" & UCase(Sex) & "'"
          '    sql = sql & " AND UPPER(Chart) ='" & UCase(ChartNo) & "'"
          '
          '    Cnxn(0).Execute Sql
          'End If
17760     ClearLabNoSelection
17770     Exit Sub


LabNoUpdatePrvData_Error:

          Dim strES As String
          Dim intEL As Integer

17780     intEL = Erl
17790     strES = Err.Description
17800     LogError "frmEditAll", "LabNoUpdatePrvData", intEL, strES, sql

End Sub

Private Function FindLatestAddress(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String) As String
          Dim sql As String

17810     On Error GoTo FindLatestAddress_Error
        
          Dim tb As New ADODB.Recordset
17820     sql = "Select Addr0 from demographics  "
17830     sql = sql & " WHERE "
17840     sql = sql & " UPPER(PatName) ='" & UCase(Name) & "'"
17850     sql = sql & " AND DoB ='" & DoB & "'"
17860     sql = sql & " AND UPPER(Sex) ='" & UCase(Sex) & "'"
17870     sql = sql & " AND UPPER(Chart) ='" & UCase(ChartNo) & "'"
17880     sql = sql & " ORDER BY DateTimeDemographics DESC  "
17890     Set tb = New Recordset
17900     RecOpenServer 0, tb, sql

17910     If Not tb.EOF Then
17920         FindLatestAddress = tb!Addr0
17930     End If


17940     Exit Function


FindLatestAddress_Error:

          Dim strES As String
          Dim intEL As Integer

17950     intEL = Erl
17960     strES = Err.Description
17970     LogError "frmEditAll", "FindLatestAddress", intEL, strES, sql
End Function

'---------------------------------------------------------------------------------------
' Procedure : ABExistsInCurrent
' Author    : XPMUser
' Date      : 27/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ABExistsInCurrent(GrdIndex As Integer)

17980     On Error GoTo ABExistsInCurrent_Error
          Dim i As Integer
          Dim Li As Integer
17990     For Li = 0 To lstABsInUse.ListCount - 1
18000         With grdAB(GrdIndex)
18010             For i = 0 To .Rows - 1
18020                 If UCase(.TextMatrix(i, 0)) = UCase(lstABsInUse.List(Li)) Then
18030                     .Col = 2
18040                     .row = i
18050                     Set .CellPicture = imgSquareTick.Picture
18060                 End If
18070             Next i
18080         End With
18090     Next Li

18100     Exit Sub


ABExistsInCurrent_Error:

          Dim strES As String
          Dim intEL As Integer

18110     intEL = Erl
18120     strES = Err.Description
18130     LogError "frmEditMicrobiology", "ABExistsInCurrent", intEL, strES
End Sub


Private Sub LoadMatchingDemo()
18140     On Error GoTo LoadMatchingDemo_Error
          Dim SearchConditon As String


18150     SearchConditon = " AND D.Dob = '" & Format(txtDoB, "dd/MMM/yyyy") & "'"
18160     If Val(txtLabNo & "") <> 0 Then
18170         If FndMatchingRecords(SearchConditon) > 0 Then
18180             With frmPatHistoryChart
18190                 .LabNoUpd = txtLabNo
18200                 Set .EditScreen = Me
18210                 .PatientHistory = SearchConditon
18220                 If .Visible = False Then
18230                     .Show 1
18240                 End If
18250                 MatchingDemoLoaded = True
18260             End With
18270             Exit Sub
18280         End If
18290     End If

18300     ClearLabNoSelection

18310     Exit Sub


LoadMatchingDemo_Error:

          Dim strES As String
          Dim intEL As Integer

18320     intEL = Erl
18330     strES = Err.Description
18340     LogError "frmEditAll", "LoadMatchingDemo", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : FndMatchingRecords
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FndMatchingRecords(Condition As String)

18350     On Error GoTo FndMatchingRecords_Error
          Dim sql As String
          Dim tb As ADODB.Recordset

18360     sql = "select Count(PatName) as Cnt from Demographics D " & _
              " WHERE D.SAMPLEID <> '-9' " & Condition & " "
          '"  GROUP BY D.PatName,D.Chart,D.Addr0,D.DoB,D.Sex "
18370     Set tb = New Recordset
18380     RecOpenClient 0, tb, sql

18390     If tb.EOF = False Then
18400         FndMatchingRecords = tb!Cnt
18410     End If


18420     Exit Function


FndMatchingRecords_Error:

          Dim strES As String
          Dim intEL As Integer

18430     intEL = Erl
18440     strES = Err.Description
18450     LogError "frmEditAll", "FndMatchingRecords", intEL, strES, sql
End Function

Private Sub ClearLabNoSelection()
18460     txtMultiSeltdDemoForLabNoUpd = ""
18470     gMDemoLabNoUpd.Clear
18480     gMDemoLabNoUpd.Rows = 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DemographicsUniLabNoSelect
' Author    : Masood
' Date      : 02/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function DemographicsUniLabNoSelect(PatName As String, DoB As String, Sex As String, Chart As String, LabNo As String) As Double

18490     On Error GoTo DemographicsUniLabNoSelect_Error
          Dim sql As String
          Dim tb As New ADODB.Recordset
18500     If PatName = "" Or DoB = "" Or Sex = "" Then
18510         Exit Function
18520     End If


18530     sql = "select Top 1 ISNULL(LabNo,0) as LabNo  from DemographicsUniLabNo As D  " & _
              " WHERE ISNULL(LabNo,0)  <> 0 AND  D.PatName='" & AddTicks(PatName) & "' AND DoB = '" & Format(DoB, "dd/MMM/yyyy") & "'" & _
              " ORDER BY DateTimeOfRecord DESC "
              
18540     Set tb = New Recordset
18550     RecOpenClient 0, tb, sql

18560     If tb.EOF = False Then
18570         DemographicsUniLabNoSelect = tb!LabNo
18580     Else
18590         LabNo = Val(FndMaxID("demographics", "LabNo", ""))
18600         Call DemographicsUniLabNoInsertValues("", UserName, PatName, DoB, Sex, Chart, LabNo)
18610         DemographicsUniLabNoSelect = LabNo
18620     End If

18630     txtLabNo = DemographicsUniLabNoSelect

18640     Exit Function


DemographicsUniLabNoSelect_Error:

          Dim strES As String
          Dim intEL As Integer

18650     intEL = Erl
18660     strES = Err.Description
18670     LogError "frmEditMicrobiology", "DemographicsUniLabNoSelect", intEL, strES
End Function


'---------------------------------------------------------------------------------------
' Procedure : VisibilityofCmdBtn
' Author    : Masood
' Date      : 03/Mar/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VisibilityofCmdBtn()
18680     On Error GoTo VisibilityofCmdBtn_Error


18690     If UCase(cmdValidateMicro.Caption) = UCase("Validate") Or UCase(cmdValidateMicro.Caption) = UCase("&Validate") Then
18700         cmdValidateMicro.Visible = False
18710         cmdIntrim.Visible = True
18720         bPrint.Visible = True
18730     Else
18740         cmdValidateMicro.Visible = True
18750         cmdIntrim.Visible = False
18760         bPrint.Visible = True

18770     End If

       
18780     Exit Sub

       
VisibilityofCmdBtn_Error:

          Dim strES As String
          Dim intEL As Integer

18790     intEL = Erl
18800     strES = Err.Description
18810     LogError "frmEditMicrobiology", "VisibilityofCmdBtn", intEL, strES

End Sub

Private Sub GetAntiBiotic(p_SampleID As String)
18820     On Error GoTo GetAntiBiotic_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          Dim tbWS As ADODB.Recordset
          
18830     fmeAntibiotics.Visible = False
18840     sql = "Select IsNULL(R.Question,'') Question, IsNULL(R.Answer,'') Answer from ocmQuestions R "
18850     sql = sql & " Inner Join OCMRequestDetails D ON D.RequestID = R.RID "
18860     sql = sql & " Where D.SampleID = '" & p_SampleID & "'"
18870     Set tb = New Recordset
18880     RecOpenServer 0, tb, sql
          
18890     If Not tb Is Nothing Then
18900         If Not tb.EOF Then
18910             fmeAntibiotics.Visible = True
18920             While Not tb.EOF
18930                 If Trim(ConvertNull(tb!Question, "")) = "Current Antibiotics" Then
18940                     txtAntibiotics.Text = tb!Answer
18950                 ElseIf Trim(ConvertNull(tb!Question, "")) = "Intended Antibiotic" Then
18960                     txtIntAntibiotics.Text = tb!Answer
18970                 End If
18980                 tb.MoveNext
18990             Wend
19000         End If
19010     End If
          
          '    Sql = "Select IsNULL(ProfileID,'') ProfileID, IsNULL(TestDescription,'') TestDescription from ocmRequestDetails Where DepartmentID = 'Micro' And SampleID = '" & p_SampleID & "'"
          '    Set tbWS = New Recordset
          '    RecOpenServer 0, tbWS, Sql
          '    If Not tbWS Is Nothing Then
          '        If Not tbWS.EOF Then
          '            cmbSite.Text = tbWS!ProfileID
          '            txtSiteDetails.Text = tbWS!TestDescription
          '        End If
          '    End If
          
19020     Exit Sub

       
GetAntiBiotic_Error:

          Dim strES As String
          Dim intEL As Integer

19030     intEL = Erl
19040     strES = Err.Description
19050     LogError "frmEditMicrobiology", "GetAntiBiotic", intEL, strES
End Sub

Private Sub SaveAntiBiotic(p_SampleID As String, p_Answer As String, p_Question As String)
19060     On Error GoTo GetAntiBiotic_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          
19070     sql = "Select RID from ocmQuestions R "
19080     sql = sql & " Inner Join OCMRequestDetails D ON D.RequestID = R.RID "
19090     sql = sql & " Where D.SampleID = '" & p_SampleID & "'"
19100     Set tb = New Recordset
19110     RecOpenServer 0, tb, sql
19120     If Not tb Is Nothing Then
19130         If Not tb.EOF Then
19140             sql = "Update ocmQuestions Set Answer = '" & p_Answer & "' Where TRIM(Question) = '" & p_Question & "' And RID = '" & ConvertNull(tb!RID, "") & "'"
19150             Cnxn(0).Execute sql
19160         End If
19170     End If
          
          
          
19180     Exit Sub

       
GetAntiBiotic_Error:

          Dim strES As String
          Dim intEL As Integer

19190     intEL = Erl
19200     strES = Err.Description
19210     LogError "frmEditMicrobiology", "SaveAntiBiotic", intEL, strES
End Sub

Private Sub ShowAntiBiotic()
19220     On Error GoTo GetAntiBiotic_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          
19230     sql = "Select Distinct AntibioticName from Antibiotics "
19240     Set tb = New Recordset
19250     RecOpenServer 0, tb, sql
19260     If Not tb Is Nothing Then
19270         If Not tb.EOF Then
19280             lstAntibio.AddItem ""
19290             While Not tb.EOF
19300                 lstAntibio.AddItem ConvertNull(tb!AntibioticName, "")
19310                 tb.MoveNext
19320             Wend
19330         End If
19340     End If
          
19350     Exit Sub

       
GetAntiBiotic_Error:

          Dim strES As String
          Dim intEL As Integer

19360     intEL = Erl
19370     strES = Err.Description
19380     LogError "frmEditMicrobiology", "ShowAntiBiotic", intEL, strES
End Sub

Private Sub FormatGrid()
19390     On Error GoTo ERROR_FormatGrid
          
19400     flxQuestions.Rows = 1
19410     flxQuestions.row = 0
          
19420     flxQuestions.ColWidth(fcsLine_NO) = 50
          
19430     flxQuestions.TextMatrix(0, fcsSr) = "Sr.#"
19440     flxQuestions.ColWidth(fcsSr) = 500
19450     flxQuestions.ColAlignment(fcsSr) = flexAlignLeftCenter
          
19460     flxQuestions.TextMatrix(0, fcsQes) = "Qustions"
19470     flxQuestions.ColWidth(fcsQes) = 2000
19480     flxQuestions.ColAlignment(fcsQes) = flexAlignLeftCenter
          
19490     flxQuestions.TextMatrix(0, fcsAns) = "Answers"
19500     flxQuestions.ColWidth(fcsAns) = 3500
19510     flxQuestions.ColAlignment(fcsAns) = flexAlignLeftCenter
              
19520     Exit Sub
ERROR_FormatGrid:
          Dim strES As String
          Dim intEL As Integer
          
19530     intEL = Erl
19540     strES = Err.Description
19550     LogError "frmEditMicrobiology", "FormatGrid", intEL, strES
End Sub

Private Sub GetOtherQuestion(p_SampleID As String)
19560     On Error GoTo GetAntiBiotic_Error

          Dim sql As String
          Dim l_Count As Integer
          Dim tb As ADODB.Recordset
          
19570     flxQuestions.Rows = 1
19580     flxQuestions.row = 0
19590     l_Count = 0
          
19600     sql = "Select IsNULL(R.Question,'') Question, IsNULL(R.Answer,'') Answer from ocmQuestions R "
19610     sql = sql & " Inner Join OCMRequestDetails D ON D.RequestID = R.RID "
19620     sql = sql & " Where D.SampleID = '" & p_SampleID & "'"
19630     Set tb = New Recordset
19640     RecOpenServer 0, tb, sql
          
19650     If Not tb Is Nothing Then
19660         If Not tb.EOF Then
19670             While Not tb.EOF
19680                 l_Count = l_Count + 1
19690                 If Trim(ConvertNull(tb!Question, "")) <> "Current Antibiotics" And Trim(ConvertNull(tb!Question, "")) <> "Intended Antibiotic" Then
19700                     flxQuestions.AddItem ("" & vbTab & l_Count & vbTab & Trim(ConvertNull(tb!Question, "")) & vbTab & Trim(ConvertNull(tb!Answer, "")))
19710                 End If
19720                 tb.MoveNext
19730             Wend
19740         End If
19750     End If
          
19760     Exit Sub

       
GetAntiBiotic_Error:

          Dim strES As String
          Dim intEL As Integer

19770     intEL = Erl
19780     strES = Err.Description
19790     LogError "frmEditMicrobiology", "GetOtherQuestion", intEL, strES
End Sub

Private Function GetRequestID(p_SampleID As String) As String
19800     On Error GoTo GetRequestID_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          
19810     GetRequestID = ""
19820     sql = "Select Distinct IsNULL(RequestID,'') RequestID from ocmRequestDetails Where SampleID = '" & p_SampleID & "'"
19830     Set tb = New Recordset
19840     RecOpenServer 0, tb, sql
19850     If Not tb Is Nothing Then
19860         If Not tb.EOF Then
19870             GetRequestID = "Request ID: " & tb!RequestID
19880         End If
19890     End If
          
19900     Exit Function

GetRequestID_Error:
          
          Dim strES As String
          Dim intEL As Integer
          
19910     GetRequestID = ""
19920     intEL = Erl
19930     strES = Err.Description
19940     LogError "frmEditMicrobiology", "GetRequestID", intEL, strES, sql
End Function

'+++Junaid 15-10-2023
Private Sub CheckUrineResults()
19950     On Error GoTo ErrorHandler:

          Dim sql As String
          Dim tb As Recordset
          Dim tbrq As Recordset
          Dim tbob As Recordset
          Dim tbre As Recordset
          Dim l_Count As Integer
          
19960     sql = "Select SampleID From Demographics Where CONVERT(VARCHAR(10),RecordDateTime,111) = CONVERT(VARCHAR(10),GetDate(),111) And ForMicro = '1'"
19970     Set tb = New Recordset
19980     RecOpenServer 0, tb, sql
19990     If Not tb Is Nothing Then
20000         If Not tb.EOF Then
20010             While Not tb.EOF
20020                 sql = "Select * from UrineRequests50 Where SampleID = '" & ConvertNull(tb!SampleID, "") & "'"
20030                 Set tbrq = New Recordset
20040                 RecOpenServer 0, tbrq, sql
20050                 If Not tbrq Is Nothing Then
20060                     If Not tbrq.EOF Then
20070                         sql = "Select * from Observations Where SampleID = '" & ConvertNull(tb!SampleID, "") & "'"
20080                         Set tbob = New Recordset
20090                         RecOpenServer 0, tbob, sql
20100                         If Not tbob Is Nothing Then
20110                             If Not tbob.EOF Then
20120                                 sql = "Select G.* from UrineResults50 G Left Join PrintValidLOG P ON P.SampleID = G.SampleID Where G.SampleID = '" & ConvertNull(tb!SampleID, "") & "'"
20130                                 Set tbre = New Recordset
20140                                 RecOpenServer 0, tbre, sql
20150                                 If tbre Is Nothing Then
20160                                     sql = "Delete from Observations Where SampleID = '" & ConvertNull(tb!SampleID, "") & "'"
20170                                     Cnxn(0).Execute sql
20180                                     MsgBox "Sample ID " & ConvertNull(tb!SampleID, "") & " had ambiguous comments", vbInformation
20190                                 Else
20200                                     If tbre.EOF Then
20210                                         sql = "Delete from Observations Where SampleID = '" & ConvertNull(tb!SampleID, "") & "'"
20220                                         Cnxn(0).Execute sql
20230                                         MsgBox "Sample ID " & ConvertNull(tb!SampleID, "") & " had ambiguous comments", vbInformation
20240                                     Else
20250                                         While Not tbre.EOF
20260                                             If (ConvertNull(tbre!TestName, "") = "WCC" And Val(Trim(ConvertNull(tbre!Result, ""))) > 39) Or (ConvertNull(tbre!TestName, "") = "Bacteria" And Val(Trim(ConvertNull(tbre!Result, ""))) > 150) Then
20270                                                 If Left(ConvertNull(tbob!Comment, ""), 50) = "This urine specimen has not met the automated CGH " Then
20280                                                     sql = "Delete from Observations Where SampleID = '" & ConvertNull(tb!SampleID, "") & "'"
20290                                                     Cnxn(0).Execute sql
20300                                                     MsgBox "Sample ID " & ConvertNull(tb!SampleID, "") & " had ambiguous comments", vbInformation
20310                                                 End If
20320                                             End If
20330                                             tbre.MoveNext
20340                                         Wend
20350                                     End If
20360                                 End If
20370                             End If
20380                         End If
20390                     End If
20400                 End If
20410                 tb.MoveNext
20420             Wend
20430         End If
20440     End If
          '    DoEvents
          '    DoEvents
          '    Unload frmWait

20450     Exit Sub
ErrorHandler:

          Dim strES As String
          Dim intEL As Integer
          '    Unload frmWait
20460     intEL = Erl
20470     strES = Err.Description
20480     LogError "frmEditMicrobiology", "CheckUrineResults", intEL, strES, sql
20490     MsgBox strES
End Sub
'--- Junaid

