VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewResults 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NetAcquire"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox lblBioHaemCoagComment 
      Height          =   1515
      Left            =   6900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   108
      Top             =   3300
      Width           =   4980
   End
   Begin VB.CommandButton cmdSetPrinter 
      BackColor       =   &H0000C000&
      Height          =   885
      Left            =   8850
      Picture         =   "frmViewResults.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   107
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   6900
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid gDem 
      Height          =   3675
      Left            =   15810
      TabIndex        =   105
      Top             =   1230
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   6482
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<Active |<RunDate             |<SampleID    |<Address     |<GP       |<Ward     |<Cnxn     |<SampleDate   |<SampleTime"
   End
   Begin VB.CommandButton cmdFAX 
      Caption         =   "FAX"
      Height          =   885
      Index           =   2
      Left            =   10440
      Picture         =   "frmViewResults.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "FAX This Report"
      Top             =   6900
      Width           =   675
   End
   Begin VB.CommandButton cmdFAX 
      Caption         =   "FAX"
      Height          =   885
      Index           =   1
      Left            =   5310
      Picture         =   "frmViewResults.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "FAX This Report"
      Top             =   6900
      Width           =   675
   End
   Begin VB.CommandButton cmdFAX 
      Caption         =   "FAX"
      Height          =   885
      Index           =   0
      Left            =   1650
      Picture         =   "frmViewResults.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   102
      ToolTipText     =   "FAX This Report"
      Top             =   6900
      Width           =   675
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   345
      Left            =   8670
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "Log as Phoned"
      Top             =   1440
      Width           =   1305
   End
   Begin VB.CommandButton cmdMedibridge 
      BackColor       =   &H000080FF&
      Caption         =   "Med. Rept"
      Height          =   945
      Left            =   10590
      Picture         =   "frmViewResults.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdExternal 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ext. Rept"
      Height          =   945
      Left            =   11250
      Picture         =   "frmViewResults.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   3
      Left            =   12690
      Picture         =   "frmViewResults.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   95
      ToolTipText     =   "Right Click to toggle Print/FAX"
      Top             =   6870
      Width           =   1245
   End
   Begin VB.Frame Frame13 
      Caption         =   "Blood Gas"
      Height          =   5535
      Left            =   12240
      TabIndex        =   80
      Top             =   1230
      Width           =   2055
      Begin VB.TextBox txtPh 
         Height          =   330
         Left            =   795
         TabIndex        =   87
         Top             =   285
         Width           =   1050
      End
      Begin VB.TextBox txtPco2 
         Height          =   330
         Left            =   795
         TabIndex        =   86
         Top             =   690
         Width           =   1050
      End
      Begin VB.TextBox txtPo2 
         Height          =   330
         Left            =   795
         TabIndex        =   85
         Top             =   1095
         Width           =   1050
      End
      Begin VB.TextBox txtHco3 
         Height          =   330
         Left            =   795
         TabIndex        =   84
         Top             =   1500
         Width           =   1050
      End
      Begin VB.TextBox txtO2Sat 
         Height          =   330
         Left            =   795
         TabIndex        =   83
         Top             =   2310
         Width           =   1050
      End
      Begin VB.TextBox txtBE 
         Height          =   330
         Left            =   795
         TabIndex        =   82
         Top             =   1905
         Width           =   1050
      End
      Begin VB.TextBox txtTotCo2 
         Height          =   330
         Left            =   795
         TabIndex        =   81
         Top             =   2715
         Width           =   1050
      End
      Begin VB.Label lblBGAComment 
         BorderStyle     =   1  'Fixed Single
         Height          =   2355
         Left            =   60
         TabIndex        =   96
         Top             =   3090
         Width           =   1935
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Ph"
         Height          =   195
         Left            =   540
         TabIndex        =   94
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PCO2"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   93
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "PO2"
         Height          =   195
         Left            =   420
         TabIndex        =   92
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "HCO3"
         Height          =   195
         Left            =   300
         TabIndex        =   91
         Top             =   1575
         Width           =   435
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "BE"
         Height          =   195
         Left            =   525
         TabIndex        =   90
         Top             =   1980
         Width           =   210
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "O2SAT"
         Height          =   195
         Left            =   210
         TabIndex        =   89
         Top             =   2430
         Width           =   525
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Tot CO2"
         Height          =   195
         Left            =   135
         TabIndex        =   88
         Top             =   2790
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   10590
      Picture         =   "frmViewResults.frx":1EBE
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   2460
      Width           =   1275
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Biochemistry Cumulative"
      Height          =   885
      Index           =   0
      Left            =   120
      Picture         =   "frmViewResults.frx":2528
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   6900
      Width           =   1245
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Haematology Cumulative"
      Height          =   885
      Index           =   1
      Left            =   3270
      Picture         =   "frmViewResults.frx":296A
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   6900
      Width           =   1245
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Coagulation Cumulative"
      Height          =   885
      Index           =   2
      Left            =   6900
      Picture         =   "frmViewResults.frx":2DAC
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6900
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   0
      Left            =   2370
      Picture         =   "frmViewResults.frx":31EE
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Print This Report"
      Top             =   6900
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   1
      Left            =   6030
      Picture         =   "frmViewResults.frx":3858
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Print This Report"
      Top             =   6900
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   2
      Left            =   11160
      Picture         =   "frmViewResults.frx":3EC2
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Print This Report"
      Top             =   6900
      Width           =   705
   End
   Begin VB.Frame Frame2 
      Height          =   5475
      Left            =   3270
      TabIndex        =   18
      Top             =   1350
      Width           =   3435
      Begin VB.TextBox tPlt 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   41
         Top             =   3870
         Width           =   825
      End
      Begin VB.TextBox tPdw 
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
         Left            =   2340
         TabIndex        =   40
         Top             =   4170
         Width           =   825
      End
      Begin VB.TextBox tMPV 
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
         Left            =   660
         TabIndex        =   39
         Top             =   4170
         Width           =   825
      End
      Begin VB.TextBox tPLCR 
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
         Left            =   2340
         TabIndex        =   38
         Top             =   3870
         Width           =   825
      End
      Begin VB.TextBox tMCV 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   37
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox tRDWSD 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   36
         Top             =   570
         Width           =   825
      End
      Begin VB.TextBox tRDWCV 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   35
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox tMCH 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   34
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox tHct 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   33
         Top             =   1170
         Width           =   825
      End
      Begin VB.TextBox tHgb 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   32
         Top             =   570
         Width           =   825
      End
      Begin VB.TextBox tRBC 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   660
         MaxLength       =   5
         TabIndex        =   31
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox tMCHC 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   30
         Top             =   1170
         Width           =   825
      End
      Begin VB.TextBox tLymP 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   29
         Top             =   2160
         Width           =   825
      End
      Begin VB.TextBox tMonoA 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   28
         Top             =   2430
         Width           =   825
      End
      Begin VB.TextBox tLymA 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   27
         Top             =   2160
         Width           =   825
      End
      Begin VB.TextBox tBasP 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   26
         Top             =   3300
         Width           =   825
      End
      Begin VB.TextBox tMonoP 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   25
         Top             =   2445
         Width           =   825
      End
      Begin VB.TextBox tNeutP 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   24
         Top             =   2730
         Width           =   825
      End
      Begin VB.TextBox tEosA 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   23
         Top             =   3000
         Width           =   825
      End
      Begin VB.TextBox tWBC 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1830
         Width           =   825
      End
      Begin VB.TextBox tNeutA 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   21
         Top             =   2730
         Width           =   825
      End
      Begin VB.TextBox tEosP 
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   20
         Top             =   3015
         Width           =   825
      End
      Begin VB.TextBox tBasA 
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
         Left            =   660
         TabIndex        =   19
         Top             =   3300
         Width           =   825
      End
      Begin VB.Label lblNotValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Not Validated"
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
         Left            =   660
         TabIndex        =   79
         Top             =   1530
         Width           =   1395
      End
      Begin VB.Image imgHaemGraphs 
         Height          =   480
         Left            =   2520
         Picture         =   "frmViewResults.frx":452C
         ToolTipText     =   "Graphs for this Sample"
         Top             =   1590
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "ESR"
         Height          =   195
         Left            =   225
         TabIndex        =   65
         Top             =   4710
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Retics"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   5085
         Width           =   450
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Monospot"
         Height          =   195
         Left            =   1560
         TabIndex        =   63
         Top             =   4680
         Width           =   705
      End
      Begin VB.Label lesr 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   660
         TabIndex        =   62
         Top             =   4650
         Width           =   825
      End
      Begin VB.Label lmonospot 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2355
         TabIndex        =   61
         Top             =   4650
         Width           =   810
      End
      Begin VB.Label lretics 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   660
         TabIndex        =   60
         Top             =   5010
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Pdw"
         Height          =   195
         Left            =   1920
         TabIndex        =   59
         Top             =   4230
         Width           =   315
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MPV"
         Height          =   195
         Left            =   270
         TabIndex        =   58
         Top             =   4230
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Plt"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   57
         Top             =   3930
         Width           =   180
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "PLCR"
         Height          =   195
         Left            =   1830
         TabIndex        =   56
         Top             =   3930
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "RDW CV"
         Height          =   195
         Left            =   1620
         TabIndex        =   55
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MCHC"
         Height          =   195
         Index           =   0
         Left            =   1815
         TabIndex        =   54
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MCH"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   53
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hct"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   52
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MCV"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "RBC"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   50
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "RDW SD"
         Height          =   195
         Index           =   0
         Left            =   1605
         TabIndex        =   49
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Hgb"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Neut"
         Height          =   195
         Left            =   1740
         TabIndex        =   47
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Mono"
         Height          =   195
         Left            =   1710
         TabIndex        =   46
         Top             =   2490
         Width           =   420
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Lymph"
         Height          =   195
         Left            =   1650
         TabIndex        =   45
         Top             =   2220
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "WBC"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   44
         Top             =   1890
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Eos"
         Height          =   195
         Left            =   1800
         TabIndex        =   43
         Top             =   3030
         Width           =   270
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Bas"
         Height          =   195
         Left            =   1770
         TabIndex        =   42
         Top             =   3300
         Width           =   270
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2550
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   11745
      Begin VB.Label lblDemogComment 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   660
         TabIndex        =   66
         Top             =   840
         Width           =   10155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   2955
         TabIndex        =   14
         Top             =   540
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   9780
         TabIndex        =   13
         Top             =   210
         Width           =   285
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3285
         TabIndex        =   12
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10095
         TabIndex        =   11
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   4425
         TabIndex        =   10
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   7290
         TabIndex        =   9
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   1170
         TabIndex        =   8
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   2835
         TabIndex        =   7
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5025
         TabIndex        =   5
         Top             =   510
         Width           =   5790
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8205
         TabIndex        =   4
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1575
         TabIndex        =   3
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3285
         TabIndex        =   2
         Top             =   210
         Width           =   3540
      End
      Begin VB.Label lblSampleID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1575
         TabIndex        =   1
         Top             =   210
         Width           =   1200
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   5355
      Left            =   120
      TabIndex        =   15
      Top             =   1470
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   9446
      _Version        =   393216
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
      FormatString    =   "<Parameter          |<Result       "
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
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1290
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gCoag 
      Height          =   1905
      Left            =   6900
      TabIndex        =   17
      Top             =   4920
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   3360
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Parameter  |<Result  |<Units      |<Comment     |^   |^      "
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
   Begin VB.Label lblProcessed 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processed in Hogwarts"
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   7590
      TabIndex        =   106
      Top             =   1440
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblWard 
      Caption         =   "lblWard"
      Height          =   225
      Left            =   15960
      TabIndex        =   101
      Top             =   480
      Width           =   2040
   End
   Begin VB.Label lblGP 
      Caption         =   "lblGP"
      Height          =   225
      Left            =   15960
      TabIndex        =   100
      Top             =   210
      Width           =   2025
   End
   Begin VB.Image imgEarliest 
      Height          =   285
      Left            =   7740
      Picture         =   "frmViewResults.frx":496E
      Stretch         =   -1  'True
      ToolTipText     =   "View Earliest Record"
      Top             =   2880
      Width           =   435
   End
   Begin VB.Image imgPrevious 
      Height          =   285
      Left            =   8190
      Picture         =   "frmViewResults.frx":4C78
      Stretch         =   -1  'True
      ToolTipText     =   "View Previous Record"
      Top             =   2880
      Width           =   435
   End
   Begin VB.Image imgNext 
      Height          =   285
      Left            =   9060
      Picture         =   "frmViewResults.frx":4F82
      Stretch         =   -1  'True
      ToolTipText     =   "View Next Record"
      Top             =   2880
      Width           =   435
   End
   Begin VB.Image imgLatest 
      Height          =   285
      Left            =   9510
      Picture         =   "frmViewResults.frx":528C
      Stretch         =   -1  'True
      ToolTipText     =   "View Most Recent Record"
      Top             =   2880
      Width           =   435
   End
   Begin VB.Label lblTimeTaken 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not Specified"
      Height          =   255
      Left            =   8670
      TabIndex        =   71
      Top             =   1950
      Width           =   1275
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Time Taken"
      Height          =   195
      Left            =   7740
      TabIndex        =   70
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Sample Date"
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   69
      Top             =   2310
      Width           =   915
   End
   Begin VB.Label lblSampleDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8670
      TabIndex        =   68
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label lblRecordInfo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Record 8888 of 8888"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7740
      TabIndex        =   67
      Top             =   2580
      Width           =   2205
   End
End
Attribute VB_Name = "frmViewResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private CurrentRecordNumber As Long

Private BioValBy As String
Private CoagValBy As String
Private HaemValBy As String


Private pPrintToPrinter As String
Public Property Let PrintToPrinter(ByVal strNewValue As String)

36600 pPrintToPrinter = strNewValue

End Property
Public Property Get PrintToPrinter() As String

36610 PrintToPrinter = pPrintToPrinter

End Property

Private Sub cmdSetPrinter_Click()

36620 frmForcePrinter.From = Me
36630 frmForcePrinter.Show 1

36640 If pPrintToPrinter = "Automatic Selection" Then
36650   pPrintToPrinter = ""
36660 End If

36670 If pPrintToPrinter <> "" Then
36680   cmdSetPrinter.BackColor = vbRed
36690   cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
36700 Else
36710   cmdSetPrinter.BackColor = vbGreen
36720   pPrintToPrinter = ""
36730   cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
36740 End If
        
End Sub

Private Sub CheckCumulative()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

36750 On Error GoTo CheckCumulative_Error

36760 For n = 0 To 2
36770   cmdCum(n).Visible = False
36780 Next

      '50    For n = 0 To intOtherHospitalsInGroup
        
      '60      If cmdCum(0).Visible = False Then
36790     sql = "SELECT COUNT(DISTINCT D.SampleID) AS Tot " & _
                "FROM Demographics AS D, BioResults AS R WHERE " & _
                "D.SampleID = R.SampleID " & _
                "AND Chart = '" & lblChart & "'" & _
                "AND PatName = '" & AddTicks(lblName) & "' "
36800     If IsDate(lblDoB) Then
36810       sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
36820     Else
36830       sql = sql & "AND (DoB IS NULL OR DoB = '') "
36840     End If
36850     Set tb = New Recordset
36860     RecOpenClient 0, tb, sql
36870     cmdCum(0).Visible = tb!Tot > 1
      '160     End If
      '
      '170     If cmdCum(1).Visible = False Then
36880     sql = "SELECT COUNT(DISTINCT D.SampleID) AS Tot " & _
                "FROM Demographics AS D, HaemResults AS R WHERE " & _
                "D.SampleID = R.SampleID " & _
                "AND Chart = '" & lblChart & "'" & _
                "AND PatName = '" & AddTicks(lblName) & "' "
36890     If IsDate(lblDoB) Then
36900       sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
36910     Else
36920       sql = sql & "AND (DoB IS NULL OR DoB = '') "
36930     End If
36940     Set tb = New Recordset
36950     RecOpenClient 0, tb, sql
36960     cmdCum(1).Visible = tb!Tot > 1
      '270     End If
      '
      '280     If cmdCum(2).Visible = False Then
36970     sql = "SELECT COUNT(DISTINCT D.SampleID) AS Tot " & _
                "FROM Demographics AS D, CoagResults AS R WHERE " & _
                "D.SampleID = R.SampleID " & _
                "AND Chart = '" & lblChart & "'" & _
                "AND PatName = '" & AddTicks(lblName) & "' "
36980     If IsDate(lblDoB) Then
36990       sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
37000     Else
37010       sql = sql & "AND (DoB IS NULL OR DoB = '') "
37020     End If
37030     Set tb = New Recordset
37040     RecOpenClient 0, tb, sql
37050     cmdCum(2).Visible = tb!Tot > 1
      '380     End If
      '
      '390   Next

37060 Exit Sub

CheckCumulative_Error:

      Dim strES As String
      Dim intEL As Integer

37070 intEL = Erl
37080 strES = Err.Description
37090 LogError "frmViewResults", "CheckCumulative", intEL, strES, sql

End Sub

Private Sub CheckExternal()

      Dim sql As String
      Dim tb As Recordset

37100 On Error GoTo CheckExternal_Error

37110 sql = "Select * from ExtResults where " & _
            "SampleID = '" & lblSampleID & "'"
37120 Set tb = New Recordset
37130 RecOpenServer 0, tb, sql
37140 If Not tb.EOF Then
37150   cmdExternal.Visible = True
37160 End If

37170 Exit Sub

CheckExternal_Error:

      Dim strES As String
      Dim intEL As Integer

37180 intEL = Erl
37190 strES = Err.Description
37200 LogError "frmViewResults", "CheckExternal", intEL, strES, sql

End Sub

Private Sub CheckIfPhoned()

37210 If CheckPhoneLog(lblSampleID).SampleID <> 0 Then
37220   cmdPhone.BackColor = vbYellow
37230   cmdPhone.Caption = "Results Phoned"
37240   cmdPhone.ToolTipText = "Results Phoned"
37250 Else
37260   cmdPhone.BackColor = &H8000000F
37270   cmdPhone.Caption = "Phone Results"
37280   cmdPhone.ToolTipText = "Phone Results"
37290 End If

End Sub

Private Sub CheckIfFaxed()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

37300 On Error GoTo CheckIfFaxed_Error

37310 For n = 0 To 2
37320   cmdFAX(n).BackColor = &H8000000F
37330   cmdFAX(n).Caption = "&Fax"
37340 Next

37350 sql = "Select Discipline from FaxLog where " & _
            "SampleID = '" & Val(lblSampleID) & "'"
37360 Set tb = Cnxn(0).Execute(sql)
37370 Do While Not tb.EOF
37380   If InStr(tb!Discipline, "B") Then
37390     cmdFAX(0).BackColor = vbYellow
37400     cmdFAX(0).Caption = "Faxed"
37410   End If
37420   If InStr(tb!Discipline, "H") Then
37430     cmdFAX(1).BackColor = vbYellow
37440     cmdFAX(1).Caption = "Faxed"
37450   End If
37460   If InStr(tb!Discipline, "C") Then
37470     cmdFAX(2).BackColor = vbYellow
37480     cmdFAX(2).Caption = "Faxed"
37490   End If
37500   tb.MoveNext
37510 Loop

37520 Exit Sub

CheckIfFaxed_Error:

      Dim strES As String
      Dim intEL As Integer

37530 intEL = Erl
37540 strES = Err.Description
37550 LogError "frmViewResults", "CheckIfFaxed", intEL, strES

End Sub
Private Sub CheckMedibridge()

      Dim sql As String
      Dim tb As Recordset

37560 On Error GoTo CheckMedibridge_Error

37570 sql = "Select * from MedibridgeResults where " & _
            "SampleID = '" & lblSampleID & "'"
37580 Set tb = New Recordset
37590 RecOpenServer 0, tb, sql
37600 If Not tb.EOF Then
37610   cmdMedibridge.Visible = True
37620 End If

37630 Exit Sub

CheckMedibridge_Error:

      Dim strES As String
      Dim intEL As Integer

37640 intEL = Erl
37650 strES = Err.Description
37660 LogError "frmViewResults", "CheckMedibridge", intEL, strES, sql

End Sub

Private Sub FillDemographics()
        
37670 On Error GoTo FillDemographics_Error

37680 lblSampleID = gDem.TextMatrix(CurrentRecordNumber, 2)
37690 lblSampleDate = gDem.TextMatrix(CurrentRecordNumber, 7)
37700 If IsDate(gDem.TextMatrix(CurrentRecordNumber, 8)) Then
37710   lblTimeTaken = gDem.TextMatrix(CurrentRecordNumber, 8)
37720 Else
37730   lblTimeTaken = "Not Specified"
37740 End If
37750 lblAddress = gDem.TextMatrix(CurrentRecordNumber, 3)
37760 lblGP = gDem.TextMatrix(CurrentRecordNumber, 4)
37770 lblWard = gDem.TextMatrix(CurrentRecordNumber, 5)

37780 If CurrentRecordNumber = 1 Then
37790   lblRecordInfo = "Most Recent Record."
37800 ElseIf CurrentRecordNumber = gDem.Rows - 1 Then
37810   lblRecordInfo = "Earliest Record."
37820 Else
37830   lblRecordInfo = "Record " & CurrentRecordNumber & " of " & gDem.Rows - 1
37840 End If

37850 imgNext.Visible = False
37860 imgLatest.Visible = False
37870 If gDem.Rows = 2 Then
37880   imgEarliest.Visible = False
37890   imgPrevious.Visible = False
37900 Else
37910   imgEarliest.Visible = True
37920   imgPrevious.Visible = True
37930 End If

37940 Exit Sub

FillDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

37950 intEL = Erl
37960 strES = Err.Description
37970 LogError "frmViewResults", "FillDemographics", intEL, strES

End Sub

Private Sub LoadAllResults()

      Dim Cn As Integer

37980 On Error GoTo LoadAllResults_Error

37990 Cn = gDem.TextMatrix(CurrentRecordNumber, 6)

38000 lblSampleDate = gDem.TextMatrix(CurrentRecordNumber, 7)

38010 LoadBiochemistry Cn
38020 LoadCoag Cn
38030 LoadHaem Cn

38040 If sysOptDeptBga(0) Then
38050   LoadBloodGas
38060 End If

38070 cmdExternal.Visible = False
38080 If sysOptDeptExt(0) Then CheckExternal

38090 cmdMedibridge.Visible = False
38100 If GetOptionSetting("DeptMediBridge", "0") <> "0" Then
38110   CheckMedibridge
38120 End If

38130 LoadComments

38140 CheckIfPhoned
38150 CheckIfFaxed

38160 lblProcessed.Visible = False
38170 If Cn <> 0 Then
38180   lblProcessed.Caption = "Processed in " & HospName(Cn)
38190   lblProcessed.Visible = True
38200 End If

38210 Exit Sub

LoadAllResults_Error:

      Dim strES As String
      Dim intEL As Integer

38220 intEL = Erl
38230 strES = Err.Description
38240 LogError "frmViewResults", "LoadAllResults", intEL, strES

End Sub

Private Sub LoadBloodGas()

      Dim Bx As BGAResult
      Dim Bxs As New BGAResults

38250 On Error GoTo LoadBloodGas_Error

38260 ClearBloodGas
38270 cmdPrint(3).Enabled = False

38280 Set Bx = Bxs.LoadResults(lblSampleID)
38290 If Not Bx Is Nothing Then
38300   With Bx
38310     txtpH = .pH
38320     txtPo2 = .PO2
38330     txtPco2 = .PCO2
38340     txtHco3 = .HCO3
38350     txtBE = .BE
38360     txtO2Sat = .O2SAT
38370     txtTotCo2 = .TotCO2
38380     lblTimeTaken = Format(.RunDateTime, "hh:mm")
38390     lblSampleDate = Format$(.Rundate, "dd/mm/yyyy")
38400   End With
38410   cmdPrint(3).Enabled = True
38420 End If

38430 Exit Sub

LoadBloodGas_Error:

      Dim strES As String
      Dim intEL As Integer

38440 intEL = Erl
38450 strES = Err.Description
38460 LogError "frmViewResults", "LoadBloodGas", intEL, strES

End Sub

Private Sub ClearBloodGas()

38470 txtBE = ""
38480 txtHco3 = ""
38490 txtO2Sat = ""
38500 txtPco2 = ""
38510 txtpH = ""
38520 txtPo2 = ""
38530 txtTotCo2 = ""
38540 lblBGAComment = ""

End Sub

Private Sub LoadComments()

      Dim OBs As Observations
        Dim AutoComment As String
        
38550 On Error GoTo LoadComments_Error

38560 lblDemogComment = ""
38570 lblBioHaemCoagComment = ""
38580 lblBGAComment = ""

38590 Set OBs = New Observations
38600 Set OBs = OBs.Load(lblSampleID, "Demographic")
38610 If Not OBs Is Nothing Then
38620   lblDemogComment = OBs.Item(1).Comment
38630 End If
38640 Set OBs = New Observations
38650 Set OBs = OBs.Load(lblSampleID, "BloodGas")
38660 If Not OBs Is Nothing Then
38670   lblBGAComment = OBs.Item(1).Comment
38680 End If
38690 Set OBs = New Observations
38700 Set OBs = OBs.Load(lblSampleID, "Biochemistry")
38710 If Not OBs Is Nothing Then
38720   lblBioHaemCoagComment = OBs.Item(1).Comment & vbCrLf
38730 End If
38740 AutoComment = CheckAutoComments(lblSampleID, 2)
38750 If Trim$(AutoComment) <> "" Then
38760   lblBioHaemCoagComment = lblBioHaemCoagComment & AutoComment & vbCrLf
38770 End If

38780 Set OBs = New Observations
38790 Set OBs = OBs.Load(lblSampleID, "Haematology")
38800 If Not OBs Is Nothing Then
38810   lblBioHaemCoagComment = lblBioHaemCoagComment & OBs.Item(1).Comment & vbCrLf
38820 End If

38830 Set OBs = New Observations
38840 Set OBs = OBs.Load(lblSampleID, "Coagulation")
38850 If Not OBs Is Nothing Then
38860   lblBioHaemCoagComment = lblBioHaemCoagComment & OBs.Item(1).Comment & vbCrLf
38870 End If
38880 AutoComment = CheckAutoComments(lblSampleID, 3)
38890 If Trim$(AutoComment) <> "" Then
38900   lblBioHaemCoagComment = lblBioHaemCoagComment & AutoComment & vbCrLf
38910 End If

38920 Set OBs = New Observations
38930 Set OBs = OBs.Load(lblSampleID, "Film")
38940 If Not OBs Is Nothing Then
38950   lblBioHaemCoagComment = lblBioHaemCoagComment & OBs.Item(1).Comment
38960 End If

38970 Exit Sub

LoadComments_Error:

      Dim strES As String
      Dim intEL As Integer

38980 intEL = Erl
38990 strES = Err.Description
39000 LogError "frmViewResults", "LoadComments", intEL, strES

End Sub

Private Sub FillgDem()

      Dim n As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim t As Single

39010 On Error GoTo FillgDem_Error

39020 With gDem
39030   .Visible = False
39040   .Rows = 2
39050   .AddItem ""
39060   .RemoveItem 1
39070 End With

39080 sql = "Select * from Demographics where " & _
            "Chart = '" & lblChart & "' " & _
            "and PatName = '" & AddTicks(lblName) & "' "
39090 If IsDate(lblDoB) Then
39100   sql = sql & "and DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
39110 Else
39120   sql = sql & "and (DoB is null or DoB = '') "
39130 End If

39140 t = Timer

39150 For n = 0 To intOtherHospitalsInGroup
39160   Set tb = New Recordset
39170   RecOpenClient n, tb, sql
39180   Do While Not tb.EOF
39190     If tb!SampleID <> lblSampleID Then
39200       s = ""
39210     Else
39220       s = "A"
39230     End If
39240     s = s & vbTab & Format$(tb!Rundate, "dd/mm/yy") & vbTab & _
              tb!SampleID & vbTab & _
              tb!Addr0 & " " & tb!Addr1 & vbTab & _
              tb!GP & vbTab & _
              tb!Ward & vbTab & _
              Format$(n) & vbTab
39250     If IsDate(tb!SampleDate) Then
39260       s = s & Format(tb!SampleDate, "dd/MM/yy") & vbTab
39270       If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
39280         s = s & Format$(tb!SampleDate, "dd/MM/yy hh:mm")
39290       Else
39300         s = s & "Not Specified"
39310       End If
39320     Else
39330       s = s & "Not Specified"
39340     End If
39350     gDem.AddItem s
          
39360     tb.MoveNext
39370   Loop
39380 Next

39390 With gDem
39400   If .Rows > 2 Then
39410     .RemoveItem 1
39420     .Col = 7
39430     .Sort = 9
39440     .Visible = True
39450   End If
39460 End With

39470 Debug.Print Timer - t

39480 For n = 1 To gDem.Rows - 1
39490   If gDem.TextMatrix(n, 0) = "A" Then
39500     CurrentRecordNumber = n
39510     Exit For
39520   End If
39530 Next

39540 If CurrentRecordNumber = 1 Then
39550     lblRecordInfo = "Most Recent Record."
39560     imgNext.Visible = False
39570     imgLatest.Visible = False
39580 ElseIf CurrentRecordNumber = gDem.Rows - 1 Then
39590   lblRecordInfo = "Earliest Record."
39600   imgEarliest.Visible = False
39610   imgPrevious.Visible = False
39620   imgNext.Visible = True
39630   imgLatest.Visible = True
39640 Else
39650   lblRecordInfo = "Record " & CurrentRecordNumber & " of " & gDem.Rows - 1
39660 End If

39670 If gDem.Rows = 2 Then
39680   imgEarliest.Visible = False
39690   imgPrevious.Visible = False
39700 End If

39710 Exit Sub

FillgDem_Error:

      Dim strES As String
      Dim intEL As Integer

39720 intEL = Erl
39730 strES = Err.Description
39740 LogError "frmViewResults", "FillgDem", intEL, strES, sql

End Sub
Private Sub cmdCancel_Click()

39750 Unload Me

End Sub

Private Sub cmdCum_Click(Index As Integer)

39760 Select Case Index
        Case 0:
39770       With frmFullHistory
39780         .Dept = "Bio"
39790         .lblSex = lblSex
39800         .lblChart = lblChart
39810         .lblDoB = lblDoB
39820         .lblName = lblName
39830         .Show 1
39840       End With
39850   Case 1:
39860     With frmFullHaem
39870       .lblChart = lblChart
39880       .lblDoB = lblDoB
39890       .lblName = lblName
39900       .Show 1
39910     End With
39920   Case 2:
39930       With frmFullHistory
39940         .Dept = "Coag"
39950         .lblSex = lblSex
39960         .lblChart = lblChart
39970         .lblDoB = lblDoB
39980         .lblName = lblName
39990         .Show 1
40000       End With
40010 End Select

End Sub

Private Sub cmdFAX_Click(Index As Integer)

40020 FAX Index

End Sub

Private Sub cmdMedibridge_Click()

40030 With frmViewMedibridge
40040   .SampleID = Val(lblSampleID) ' + sysOptMicroOffset(0)
40050   .Show 1
40060 End With

End Sub

Private Sub cmdPhone_Click()

40070 With frmPhoneLog
40080   .SampleID = lblSampleID
40090   If lblGP <> "" Then
40100     .GP = lblGP
40110     .WardOrGP = "GP"
40120   Else
40130     .GP = lblWard
40140     .WardOrGP = "Ward"
40150   End If
40160   .Show 1
40170 End With

40180 CheckIfPhoned

End Sub

Private Sub cmdPrint_Click(Index As Integer)
        
      Dim sql As String
      Dim tb As Recordset
      Dim Ward As String
      Dim Clin As String
      Dim GP As String

40190 On Error GoTo cmdPrint_Click_Error

40200 If DateDiff("d", GetOptionSetting("WardEnqV7Date", "01/May/2011"), lblSampleDate) > 0 Then

        Dim f As Form

40210   Set f = New frmReportViewer

40220   f.Dept = "Biochemistry"
40230   f.SampleID = lblSampleID
40240   f.InhibitChoosePrinter = False
40250   f.PrintToPrinter = pPrintToPrinter
40260   f.Show 1

40270   Set f = Nothing

40280   Exit Sub
40290 End If

40300 GetWardClinGP lblSampleID, Ward, Clin, GP

40310 sql = "Select * from PrintPending where " & _
            "Department = '" & Choose(Index + 1, "B", "H", "C", "E") & "' " & _
            "and SampleID = '" & lblSampleID & "'"
40320 Set tb = New Recordset
40330 RecOpenClient 0, tb, sql
40340 If tb.EOF Then
40350   tb.AddNew
40360 End If
40370 tb!PrintOnCondition = gONLYVALID
40380 tb!SampleID = lblSampleID
40390 tb!Ward = Ward
40400 tb!Clinician = Clin
40410 tb!GP = GP
40420 tb!Department = Choose(Index + 1, "B", "H", "C", "E")
40430 tb!Initiator = UserName
40440 tb!UsePrinter = pPrintToPrinter
40450 tb!ThisIsCopy = 1
40460 tb.Update

40470 cmdPrint(Index).Enabled = False

40480 Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40490 intEL = Erl
40500 strES = Err.Description
40510 LogError "frmViewResults", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdExternal_Click()

40520 With frmExternalReport
40530   .lblChart = lblChart
40540   .lblName = lblName
40550   .lblDoB = lblDoB
40560   .Show 1
40570 End With

End Sub

Private Sub cmdPrint_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
                                                                                          '
                                                                                          '10    If Button = vbRightButton Then
                                                                                          '20      If cmdPrint(index).Caption = "&Print" Then
                                                                                          '30        cmdPrint(index).Caption = "&FAX"
                                                                                          '40        cmdPrint(index).Picture = frmMain.ImageList1.ListImages("Fax").Picture
                                                                                          '50      Else
                                                                                          '60        cmdPrint(index).Caption = "&Print"
                                                                                          '70        cmdPrint(index).Picture = frmMain.ImageList1.ListImages("Printer").Picture
                                                                                          '80      End If
                                                                                          '90    End If
                                                                                          '

End Sub

Private Sub Form_Activate()

40580 pBar.max = LogOffDelaySecs
40590 pBar = 0

40600 Timer1.Enabled = True

40610 If Activated Then Exit Sub
40620 Activated = True

      'If IsIDE Then
40630   FillgDem
      'Else
      '  LoadInitialDemographics
      'End If
40640 CheckCumulative
40650 LoadAllResults

End Sub
Private Sub FAX(ByVal Index As Integer)

      Dim tb As Recordset
      Dim tbO As Recordset
      Dim sql As String
      Dim FaxNumber As String
      Dim Ward As String
      Dim Clin As String
      Dim GenP As String
      Dim Initiator As String
      Dim Department As String
      Dim f As Form
40660 On Error GoTo FAX_Error

40670 ReDim strFaxNumbers(0 To 0) As String
      Dim n As Integer

40680 GetWardClinGP lblSampleID, Ward, Clin, GenP
        
40690 Department = Choose(Index + 1, "B", "H", "C")

      Dim Gx As New GP
40700 Gx.LoadName GenP

40710 FaxNumber = Gx.FAX
40720 If FaxNumber = "" Then
40730   FaxNumber = IsFaxable("Wards", GenP)
40740 End If

40750  sql = "IF NOT EXISTS (SELECT Name FROM sysobjects WHERE xtype = 'U' AND Name = 'TempFax') " & _
            "CREATE TABLE TempFax (Fax nvarchar(50)) " & _
            "TRUNCATE TABLE TempFax " & _
            "INSERT INTO TempFax " & _
            "SELECT DISTINCT(Fax) FROM GPs WHERE Fax IS NOT NULL AND Fax <> '' " & _
            "INSERT INTO TempFax " & _
            "SELECT DISTINCT(Fax) FROM Wards WHERE Fax IS NOT NULL AND Fax <> ''"
40760 Cnxn(0).Execute sql
40770 sql = "SELECT DISTINCT(Fax) from TempFax " & _
            "ORDER BY Fax"
40780 Set tb = New Recordset
40790 RecOpenServer 0, tb, sql
40800 n = -1
40810 Do While Not tb.EOF
40820   n = n + 1
40830   ReDim Preserve strFaxNumbers(0 To n)
40840   strFaxNumbers(n) = tb!FAX
40850   tb.MoveNext
40860 Loop

40870 Set f = New fcdrDBox
40880 With f
40890   .Default = FaxNumber
40900   .ListOrCombo = "List"
40910   .Options = strFaxNumbers
40920   .Prompt = "Confirm FAX Number" & vbCrLf & "(Leave blank to Cancel FAX)"
40930   .Show 1
40940   FaxNumber = .ReturnValue
40950 End With
40960 Unload f
40970 Set f = Nothing

40980 If FaxNumber = "" Then
40990   iMsg "FAX Cancelled!", vbInformation
41000   Exit Sub
41010 End If

41020 If Index < 3 Then
41030   sql = "Select * from PrintPending where " & _
        "Department = '" & Department & "' " & _
        "and SampleID = '" & lblSampleID & "' " & _
        "and UsePrinter = 'FAX'"
41040   Set tb = New Recordset
41050   RecOpenClient 0, tb, sql
41060   If tb.EOF Then
41070     tb.AddNew
41080   End If
41090   tb!SampleID = lblSampleID
41100   tb!Ward = Ward
41110   tb!Clinician = Clin
41120   tb!GP = GenP
41130   tb!UsePrinter = "FAX"
41140   tb!FaxNumber = FaxNumber

41150   Select Case Index
          Case 0: sql = "Select Operator, Faxed from BioResults " & _
            "where SampleID = '" & lblSampleID & "'"
41160 Set tbO = New Recordset
41170 RecOpenServer 0, tbO, sql
41180 If Not tbO.EOF Then
41190   Initiator = tbO!Operator & ""
41200   tbO!FAXed = 1
41210   tbO.Update
41220 End If
41230 tb!Department = "B"
41240 tb!Initiator = Initiator
          
41250     Case 1: sql = "Select Operator, Faxed from HaemResults " & _
            "where SampleID = '" & lblSampleID & "'"
41260 Set tbO = New Recordset
41270 RecOpenServer 0, tbO, sql
41280 If Not tbO.EOF Then
41290   Initiator = tbO!Operator & ""
41300   tbO!FAXed = 1
41310   tbO.Update
41320 End If
41330 tb!Department = "H"
41340 tb!Initiator = Initiator
          
41350     Case 2: sql = "Update CoagResults " & _
            "Set Faxed = 1 where " & _
            "SampleID = '" & lblSampleID & "'"
41360 Cnxn(0).Execute sql
41370 tb!Department = "C"
41380 tb!Initiator = ""
        
41390   End Select
41400   tb.Update
        
41410   UpdateFaxLog lblSampleID, Department, FaxNumber
        
41420   cmdFAX(Index).BackColor = vbYellow
41430   cmdFAX(Index).Caption = "Faxed"

41440 End If

41450 Exit Sub

FAX_Error:

      Dim strES As String
      Dim intEL As Integer

41460 intEL = Erl
41470 strES = Err.Description
41480 LogError "frmViewResults", "FAX", intEL, strES, sql

End Sub
Private Sub LoadHaem(ByVal Connection As String)

      Dim tb As Recordset
      Dim sql As String

41490 On Error GoTo LoadHaem_Error

41500 If Trim(lblSampleID) = "" Then Exit Sub

41510 ClearHaem
41520 imgHaemGraphs.Visible = False
41530 lblNotValid.Visible = False

41540 cmdPrint(1).Enabled = False

41550 sql = "Select * from HaemResults where " & _
            "SampleID = '" & lblSampleID & "'"
41560 Set tb = New Recordset
41570 RecOpenClient Connection, tb, sql
41580 If Not tb.EOF Then
41590   If tb!Valid = 0 Or IsNull(tb!Valid) Then
41600     lblNotValid.Visible = True
41610     imgHaemGraphs.Visible = False
41620   Else
41630     cmdPrint(1).Enabled = True
41640     HaemValBy = tb!Operator & ""
41650     If Not IsNull(tb!gwb1) Or Not IsNull(tb!gwb2) Or Not IsNull(tb!gRBC) Or Not IsNull(tb!gplt) Or Not IsNull(tb!gplth) Then
41660       imgHaemGraphs.Visible = True
41670     End If
          
41680     If Trim$(tb!rbc & "") <> "" Then
41690       Colourise "RBC", tRBC, tb!rbc, lblSex, lblDoB
41700     End If
          
41710     If Trim$(tb!Hgb & "") <> "" Then
41720       Colourise "Hgb", tHgb, tb!Hgb, lblSex, lblDoB
41730     End If
          
41740     If Trim$(tb!MCV & "") <> "" Then
41750       Colourise "MCV", tMCV, tb!MCV, lblSex, lblDoB
41760     End If
          
41770     If Trim$(tb!hct & "") <> "" Then
41780       Colourise "Hct", tHct, tb!hct, lblSex, lblDoB
41790     End If
          
41800     If Trim$(tb!RDWCV & "") <> "" Then
41810       Colourise "RDWCV", tRDWCV, tb!RDWCV, lblSex, lblDoB
41820     End If
          
41830     If Trim$(tb!rdwsd & "") <> "" Then
41840       Colourise "RDWSD", tRDWSD, tb!rdwsd, lblSex, lblDoB
41850     End If
          
41860     If Trim$(tb!mch & "") <> "" Then
41870       Colourise "MCH", tMCH, tb!mch, lblSex, lblDoB
41880     End If
          
41890     If Trim$(tb!mchc & "") <> "" Then
41900       Colourise "MCHC", tMCHC, tb!mchc, lblSex, lblDoB
41910     End If
          
41920     If Trim$(tb!plt & "") <> "" Then
41930       Colourise "plt", tPlt, tb!plt, lblSex, lblDoB
41940     End If
          
41950     If Trim$(tb!mpv & "") <> "" Then
41960       Colourise "MPV", tMPV, tb!mpv, lblSex, lblDoB
41970     End If
          
41980     If Trim$(tb!plcr & "") <> "" Then
41990       Colourise "PLCR", tPLCR, tb!plcr, lblSex, lblDoB
42000     End If
          
42010     If Trim$(tb!pdw & "") <> "" Then
42020       Colourise "Pdw", tPdw, tb!pdw, lblSex, lblDoB
42030     End If
          
42040     If Trim$(tb!WBC & "") <> "" Then
42050       Colourise "WBC", tWBC, tb!WBC, lblSex, lblDoB
42060     End If
          
42070     If Trim$(tb!LymA & "") <> "" Then
42080       Colourise "LymA", tLymA, tb!LymA, lblSex, lblDoB
42090     End If
          
42100     If Trim$(tb!LymP & "") <> "" Then
42110       Colourise "LymP", tLymP, tb!LymP, lblSex, lblDoB
42120     End If
          
42130     If Trim$(tb!MonoA & "") <> "" Then
42140       Colourise "MonoA", tMonoA, tb!MonoA, lblSex, lblDoB
42150     End If
          
42160     If Trim$(tb!MonoP & "") <> "" Then
42170       Colourise "MonoP", tMonoP, tb!MonoP, lblSex, lblDoB
42180     End If
          
42190     If Trim$(tb!NeutA & "") <> "" Then
42200       Colourise "NeutA", tNeutA, tb!NeutA, lblSex, lblDoB
42210     End If
          
42220     If Trim$(tb!NeutP & "") <> "" Then
42230       Colourise "NeutP", tNeutP, tb!NeutP, lblSex, lblDoB
42240     End If
          
42250     If Trim$(tb!EosA & "") <> "" Then
42260       Colourise "EosA", tEosA, tb!EosA, lblSex, lblDoB
42270     End If
          
42280     If Trim$(tb!EosP & "") <> "" Then
42290       Colourise "EosP", tEosP, tb!EosP, lblSex, lblDoB
42300     End If
          
42310     If Trim$(tb!BasA & "") <> "" Then
42320       Colourise "BasA", tBasA, tb!BasA, lblSex, lblDoB
42330     End If
          
42340     If Trim$(tb!BasP & "") <> "" Then
42350       Colourise "BasP", tBasP, tb!BasP, lblSex, lblDoB
42360     End If
42370     lesr = tb!ESR & ""
          
42380     lretics = tb!RetA & ""
42390     If Trim$(tb!RetP & "") <> "" Then
42400       lretics = lretics & "(" & tb!RetP & "%)"
42410     End If
42420     lmonospot = tb!MonoSpot & ""
42430   End If
42440 End If

42450 Screen.MousePointer = 0

42460 Exit Sub

LoadHaem_Error:

      Dim strES As String
      Dim intEL As Integer

42470 intEL = Erl
42480 strES = Err.Description
42490 LogError "frmViewResults", "LoadHaem", intEL, strES, sql

End Sub
Private Sub Colourise(ByVal Analyte As String, _
                      ByVal Destination As TextBox, _
                      ByVal strValue As String, _
                      ByVal Sex As String, _
                      ByVal DoB As String)
          
      Dim Value As Single

42500 On Error GoTo Colourise_Error

42510 Value = Val(strValue)

42520 Destination.Text = strValue
42530 If Trim$(strValue) = "" Then
42540   Destination.BackColor = &HFFFFFF
42550   Destination.ForeColor = &H0&
42560   Exit Sub
42570 End If

42580 Select Case InterpH(Value, Analyte, Sex, DoB)
        Case "X":
42590     Destination.BackColor = vbBlack
42600     Destination.ForeColor = vbWhite
42610   Case "H":
42620     Destination.BackColor = &HFFFF&
42630     Destination.ForeColor = &HFF&
42640   Case "L"
42650     Destination.BackColor = &HFFFF00
42660     Destination.ForeColor = &HC00000
42670   Case Else
42680     Destination.BackColor = &HFFFFFF
42690     Destination.ForeColor = &H0&
42700 End Select

42710 Exit Sub

Colourise_Error:

      Dim strES As String
      Dim intEL As Integer

42720 intEL = Erl
42730 strES = Err.Description
42740 LogError "frmViewResults", "Colourise", intEL, strES
        
End Sub

Private Sub ClearHaem()

42750 tWBC = ""
42760 tWBC.BackColor = &HFFFFFF
42770 tWBC.ForeColor = &H0&

42780 tRBC = ""
42790 tRBC.BackColor = &HFFFFFF
42800 tRBC.ForeColor = &H0&

42810 tHgb = ""
42820 tHgb.BackColor = &HFFFFFF
42830 tHgb.ForeColor = &H0&

42840 tMCV = ""
42850 tMCV.BackColor = &HFFFFFF
42860 tMCV.ForeColor = &H0&

42870 tHct = ""
42880 tHct.BackColor = &HFFFFFF
42890 tHct.ForeColor = &H0&

42900 tRDWCV = ""
42910 tRDWCV.BackColor = &HFFFFFF
42920 tRDWCV.ForeColor = &H0&

42930 tRDWSD = ""
42940 tRDWSD.BackColor = &HFFFFFF
42950 tRDWSD.ForeColor = &H0&

42960 tMCH = ""
42970 tMCH.BackColor = &HFFFFFF
42980 tMCH.ForeColor = &H0&

42990 tMCHC = ""
43000 tMCHC.BackColor = &HFFFFFF
43010 tMCHC.ForeColor = &H0&

43020 tPlt = ""
43030 tPlt.BackColor = &HFFFFFF
43040 tPlt.ForeColor = &H0&

43050 tMPV = ""
43060 tMPV.BackColor = &HFFFFFF
43070 tMPV.ForeColor = &H0&

43080 tPLCR = ""
43090 tPLCR.BackColor = &HFFFFFF
43100 tPLCR.ForeColor = &H0&

43110 tPdw = ""
43120 tPdw.BackColor = &HFFFFFF
43130 tPdw.ForeColor = &H0&

43140 tLymA = ""
43150 tLymA.BackColor = &HFFFFFF
43160 tLymA.ForeColor = &H0&

43170 tLymP = ""
43180 tLymP.BackColor = &HFFFFFF
43190 tLymP.ForeColor = &H0&

43200 tMonoA = ""
43210 tMonoA.BackColor = &HFFFFFF
43220 tMonoA.ForeColor = &H0&

43230 tMonoP = ""
43240 tMonoP.BackColor = &HFFFFFF
43250 tMonoP.ForeColor = &H0&

43260 tNeutA = ""
43270 tNeutA.BackColor = &HFFFFFF
43280 tNeutA.ForeColor = &H0&

43290 tNeutP = ""
43300 tNeutP.BackColor = &HFFFFFF
43310 tNeutP.ForeColor = &H0&

43320 tEosA = ""
43330 tEosA.BackColor = &HFFFFFF
43340 tEosA.ForeColor = &H0&

43350 tEosP = ""
43360 tEosP.BackColor = &HFFFFFF
43370 tEosP.ForeColor = &H0&

43380 tBasA = ""
43390 tBasA.BackColor = &HFFFFFF
43400 tBasA.ForeColor = &H0&

43410 tBasP = ""
43420 tBasP.BackColor = &HFFFFFF
43430 tBasP.ForeColor = &H0&

43440 lesr = ""
43450 lretics = ""
43460 lmonospot = ""

End Sub

Private Sub LoadCoag(ByVal Connection As Integer)

      Dim Cxs As New CoagResults
      Dim Cx As CoagResult
      Dim s As String
      Dim FormatStr As String

43470 On Error GoTo LoadCoag_Error

43480 gCoag.Rows = 2
43490 gCoag.AddItem ""
43500 gCoag.RemoveItem 1

43510 cmdPrint(2).Enabled = False

43520 Set Cxs = Cxs.Load(lblSampleID, gDONTCARE, gDONTCARE, "Results", Connection)
43530 If Cxs.Count <> 0 Then
43540   cmdPrint(2).Enabled = True
43550   For Each Cx In Cxs

43560     s = Cx.TestName & vbTab
43570     Select Case Cx.DP
            Case 0: FormatStr = "###0"
43580       Case 1: FormatStr = "##0.0"
43590       Case 2: FormatStr = "#0.00"
43600       Case 3: FormatStr = "0.000"
43610     End Select
43620     If Cx.Valid Then
43630       s = s & Format(Cx.Result, FormatStr)
43640       CoagValBy = Cx.OperatorCode
43650     Else
43660       s = s & "NV"
43670     End If
43680     s = s & vbTab & Cx.Units & vbTab
43690     s = s & vbTab
43700     If Val(Cx.Result) <> 0 Then
43710       If Cx.Result < Cx.Low Then
43720         s = s & "L"
43730       ElseIf Cx.Result > Cx.High Then
43740         s = s & "H"
43750       End If
43760     End If
43770     s = s & vbTab & _
              IIf(Cx.Valid, "V", "") & vbTab & _
              IIf(Cx.Printed, "P", "")
43780     gCoag.AddItem s
43790   Next
43800 End If

43810 If gCoag.Rows > 2 Then
43820   gCoag.RemoveItem 1
43830 End If

43840 Exit Sub

LoadCoag_Error:

      Dim strES As String
      Dim intEL As Integer

43850 intEL = Erl
43860 strES = Err.Description
43870 LogError "frmViewResults", "LoadCoag", intEL, strES

End Sub

Private Sub LoadBiochemistry(ByVal Cn As Integer)

      Dim s As String
      Dim Value As Single
      Dim valu As String
      Dim BRs As New BIEResults
      Dim BR As BIEResult

43880 On Error GoTo LoadBiochemistry_Error

43890 Set BRs = BRs.Load("Bio", lblSampleID, "Results", gDONTCARE, gDONTCARE, "", Cn)

43900 gBio.Visible = False
43910 gBio.Rows = 2
43920 gBio.AddItem ""
43930 gBio.RemoveItem 1
43940 cmdPrint(0).Enabled = False

43950 For Each BR In BRs
43960   cmdPrint(0).Enabled = True
43970   If BR.Valid Then
43980     BioValBy = BR.Operator
43990     If IsNumeric(BR.Result) Then
44000       Value = Val(BR.Result)
44010     Select Case BR.Printformat
            Case 0: valu = Format(Value, "0")
44020       Case 1: valu = Format(Value, "0.0")
44030       Case 2: valu = Format(Value, "0.00")
44040       Case 3: valu = Format(Value, "0.000")
44050       Case Else: valu = Format(Value, "0.000")
44060     End Select
44070     Else
44080       valu = BR.Result
44090     End If
44100   Else
44110     valu = "NV"
44120   End If
44130   s = BR.LongName & vbTab & valu
44140   gBio.AddItem s
44150   If BR.Valid Then
44160     s = QuickInterpBio(BR)
44170     Select Case Trim$(s)
            Case "Low":
44180         gBio.row = gBio.Rows - 1
44190         gBio.Col = 1
44200         gBio.CellBackColor = &HFFFF80
44210 Case "High":
44220         gBio.row = gBio.Rows - 1
44230         gBio.Col = 1
44240         gBio.CellBackColor = vbRed
44250         gBio.CellForeColor = vbYellow
44260 Case Else:
44270         gBio.row = gBio.Rows - 1
44280         gBio.Col = 1
44290         gBio.CellBackColor = 0
44300     End Select
44310   End If
44320 Next

44330 If gBio.Rows > 2 Then
44340   gBio.RemoveItem 1
44350 End If

44360 gBio.Visible = True

44370 Exit Sub

LoadBiochemistry_Error:

      Dim strES As String
      Dim intEL As Integer

44380 intEL = Erl
44390 strES = Err.Description
44400 LogError "frmViewResults", "LoadBiochemistry", intEL, strES

End Sub

Private Sub Form_Deactivate()

44410 Timer1.Enabled = False

End Sub


Private Sub Form_Load()

      Dim n As Integer

44420 Activated = False

44430 gBio.Font.Bold = True

44440 pBar.max = LogOffDelaySecs
44450 pBar = 0

44460 If sysOptDeptBga(0) Then
44470   Me.width = 14625
44480 Else
44490   Me.width = 12120
44500 End If

44510 For n = 0 To 3
44520   cmdPrint(n).Caption = "&Print"
44530   cmdPrint(n).Picture = frmMain.ImageList1.ListImages("Printer").Picture
44540 Next

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

44550 pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

44560 Activated = False

End Sub


Private Sub gCoag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

44570 On Error GoTo gCoag_MouseMove_Error

44580 gCoag.ToolTipText = ""
44590 If gCoag.MouseRow > 0 And gCoag.MouseCol = 3 Then
44600   If Trim$(gCoag.TextMatrix(gCoag.MouseRow, 3)) <> "" Then
44610     gCoag.ToolTipText = gCoag.TextMatrix(gCoag.MouseRow, 3)
44620   End If
44630 End If

44640 Exit Sub

gCoag_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer

44650 intEL = Erl
44660 strES = Err.Description
44670 LogError "frmViewResults", "gCoag_MouseMove", intEL, strES

End Sub

Private Sub gDem_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

44680 If Not IsDate(gDem.TextMatrix(Row1, 7)) Then
44690   Cmp = 0
44700   Exit Sub
44710 End If

44720 If Not IsDate(gDem.TextMatrix(Row2, 7)) Then
44730   Cmp = 0
44740   Exit Sub
44750 End If

44760 d1 = Format(gDem.TextMatrix(Row1, 7), "dd/mmm/yyyy")
44770 d2 = Format(gDem.TextMatrix(Row2, 7), "dd/mmm/yyyy")

44780 Cmp = Sgn(DateDiff("s", d1, d2))

End Sub


Private Sub imgEarliest_Click()

44790 CurrentRecordNumber = 1

44800 FillDemographics
44810 LoadAllResults

44820 imgEarliest.Visible = False
44830 imgPrevious.Visible = False
44840 If gDem.Rows > 2 Then
44850   imgNext.Visible = True
44860   imgLatest.Visible = True
44870 End If

44880 pBar = 0

End Sub

Private Sub imgHaemGraphs_Click()

44890   frmHaemGraphs.SampleID = lblSampleID
44900   frmHaemGraphs.Show 1

End Sub

Private Sub imgLatest_Click()

44910 CurrentRecordNumber = 1

44920 FillDemographics
44930 LoadAllResults

44940 imgNext.Visible = False
44950 imgLatest.Visible = False

44960 If gDem.Rows > 2 Then
44970   imgPrevious.Visible = True
44980   imgEarliest.Visible = True
44990 End If

45000 pBar = 0

End Sub

Private Sub imgNext_Click()

45010 If CurrentRecordNumber > 1 Then
45020   CurrentRecordNumber = CurrentRecordNumber - 1
45030 End If

45040 FillDemographics
45050 LoadAllResults

45060 If CurrentRecordNumber > 1 Then
45070   imgNext.Visible = True
45080   imgLatest.Visible = True
45090 Else
45100   imgNext.Visible = False
45110   imgLatest.Visible = False
45120 End If

45130 If gDem.Rows > 2 Then
45140   imgPrevious.Visible = True
45150   imgEarliest.Visible = True
45160 End If

45170 pBar = 0

End Sub

Private Sub imgPrevious_Click()

45180 If CurrentRecordNumber < gDem.Rows - 1 Then
45190   CurrentRecordNumber = CurrentRecordNumber + 1
45200 End If

45210 FillDemographics
45220 LoadAllResults

45230 If CurrentRecordNumber < gDem.Rows - 1 Then
45240   imgEarliest.Visible = True
45250   imgPrevious.Visible = True
45260 Else
45270   imgEarliest.Visible = False
45280   imgPrevious.Visible = False
45290 End If

45300 If gDem.Rows > 2 Then
45310   imgNext.Visible = True
45320   imgLatest.Visible = True
45330 End If

45340 pBar = 0

End Sub


Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
45350 pBar = pBar + 1
        
45360 If pBar = pBar.max Then
45370   Unload Me
45380 End If

End Sub


