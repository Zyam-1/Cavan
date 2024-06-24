VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewResultsWE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire --- View Results"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13800
   ControlBox      =   0   'False
   HelpContextID   =   10021
   Icon            =   "frmViewResultsWE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   13800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdViewAcreditationStatus 
      Caption         =   "View Test Accreditation Status"
      Height          =   1020
      Left            =   12510
      Picture         =   "frmViewResultsWE.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   5490
      Width           =   1200
   End
   Begin VB.CommandButton cmdViewScan 
      Caption         =   "&View Scan"
      Height          =   1020
      Left            =   12510
      Picture         =   "frmViewResultsWE.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   4410
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdSignOffCoag 
      Caption         =   "Sign OFF"
      Height          =   885
      Left            =   12480
      Picture         =   "frmViewResultsWE.frx":69C2
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   8475
      Width           =   1245
   End
   Begin VB.CommandButton cmdSignOff 
      Caption         =   "Sign OFF"
      Height          =   800
      Left            =   4035
      Picture         =   "frmViewResultsWE.frx":728C
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   3990
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTabBio 
      Height          =   4395
      Left            =   5175
      TabIndex        =   114
      Top             =   1560
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   7752
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Biochemistry"
      TabPicture(0)   =   "frmViewResultsWE.frx":7B56
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gBio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRedCross(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGreenTick(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "External"
      TabPicture(1)   =   "frmViewResultsWE.frx":7B72
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRedCross(1)"
      Tab(1).Control(1)=   "cmdGreenTick(1)"
      Tab(1).Control(2)=   "gBiomnis"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdRedCross 
         Height          =   285
         Index           =   1
         Left            =   -68400
         Picture         =   "frmViewResultsWE.frx":7B8E
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdGreenTick 
         Height          =   285
         Index           =   1
         Left            =   -68715
         Picture         =   "frmViewResultsWE.frx":7E64
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdGreenTick 
         Height          =   285
         Index           =   0
         Left            =   6360
         Picture         =   "frmViewResultsWE.frx":813A
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdRedCross 
         Height          =   285
         Index           =   0
         Left            =   6660
         Picture         =   "frmViewResultsWE.frx":8410
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   0
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid gBio 
         Height          =   3945
         Left            =   15
         TabIndex        =   115
         Top             =   420
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   12
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
         FormatString    =   "<Parameter                  |<Result       |<Ranges    |<Units    |<C|<Source         |<S |<Code |<Valid"
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
      Begin MSFlexGridLib.MSFlexGrid gBiomnis 
         Height          =   3915
         Left            =   -74970
         TabIndex        =   116
         Top             =   420
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   6906
         _Version        =   393216
         Cols            =   12
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
         FormatString    =   "<Parameter                  |<Result                 |<Ranges    |<Units    |<C|<Source       |<S |<Code|<Valid"
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
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3855
      Left            =   14520
      TabIndex        =   113
      Top             =   2550
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Chart |<Date Of Birth |<Name                           "
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
   Begin VB.ListBox lstExt 
      Height          =   1815
      Left            =   17580
      TabIndex        =   110
      Top             =   7320
      Width           =   1515
   End
   Begin VB.CommandButton cmdComments 
      BackColor       =   &H0000FFFF&
      Caption         =   "Comments Available - Click to View"
      Height          =   840
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   5805
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdSemen 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Semen Rept."
      Height          =   855
      Left            =   10020
      Picture         =   "frmViewResultsWE.frx":86E6
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   420
      Width           =   615
   End
   Begin VB.CommandButton cmdBloodGroup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BT Result"
      Height          =   855
      Left            =   12030
      Picture         =   "frmViewResultsWE.frx":8B28
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   420
      Width           =   615
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Request Lab Services"
      Height          =   1005
      Left            =   17460
      Picture         =   "frmViewResultsWE.frx":8E32
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   600
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdMicro 
      BackColor       =   &H0000FF00&
      Caption         =   "Micro Report"
      Height          =   855
      Left            =   10680
      Picture         =   "frmViewResultsWE.frx":9274
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   420
      Width           =   615
   End
   Begin VB.CommandButton cmdExternal 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ext. Rept"
      Height          =   855
      Left            =   11370
      Picture         =   "frmViewResultsWE.frx":96B6
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   420
      Width           =   615
   End
   Begin VB.Frame fraBGA 
      Caption         =   "Blood Gas"
      ForeColor       =   &H80000017&
      Height          =   5505
      Left            =   14820
      TabIndex        =   79
      Top             =   1410
      Width           =   2055
      Begin VB.TextBox txtTotCo2 
         Height          =   330
         Left            =   795
         TabIndex        =   86
         Top             =   2715
         Width           =   1050
      End
      Begin VB.TextBox txtBE 
         Height          =   330
         Left            =   795
         TabIndex        =   85
         Top             =   1905
         Width           =   1050
      End
      Begin VB.TextBox txtO2Sat 
         Height          =   330
         Left            =   795
         TabIndex        =   84
         Top             =   2310
         Width           =   1050
      End
      Begin VB.TextBox txtHco3 
         Height          =   330
         Left            =   795
         TabIndex        =   83
         Top             =   1500
         Width           =   1050
      End
      Begin VB.TextBox txtPo2 
         Height          =   330
         Left            =   795
         TabIndex        =   82
         Top             =   1095
         Width           =   1050
      End
      Begin VB.TextBox txtPco2 
         Height          =   330
         Left            =   795
         TabIndex        =   81
         Top             =   690
         Width           =   1050
      End
      Begin VB.TextBox txtPh 
         Height          =   330
         Left            =   795
         TabIndex        =   80
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Tot CO2"
         Height          =   195
         Left            =   135
         TabIndex        =   94
         Top             =   2790
         Width           =   600
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "O2SAT"
         Height          =   195
         Left            =   210
         TabIndex        =   93
         Top             =   2430
         Width           =   525
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "BE"
         Height          =   195
         Left            =   525
         TabIndex        =   92
         Top             =   1980
         Width           =   210
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
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "PO2"
         Height          =   195
         Left            =   420
         TabIndex        =   90
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PCO2"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   89
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Ph"
         Height          =   195
         Left            =   540
         TabIndex        =   88
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblBGAComment 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   60
         TabIndex        =   87
         Top             =   3090
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   3
      Left            =   14940
      Picture         =   "frmViewResultsWE.frx":9AF8
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   6690
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   4035
      Picture         =   "frmViewResultsWE.frx":A162
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   7965
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   885
      Index           =   2
      Left            =   12480
      Picture         =   "frmViewResultsWE.frx":A7CC
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Print Coagulation."
      Top             =   7530
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   800
      Index           =   1
      Left            =   4035
      Picture         =   "frmViewResultsWE.frx":AE36
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Print Haematology."
      Top             =   3135
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   885
      Index           =   0
      Left            =   12480
      Picture         =   "frmViewResultsWE.frx":B4A0
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Print Biochemistry. "
      Top             =   2490
      Width           =   1245
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Coagulation Cumulative"
      Height          =   885
      Index           =   2
      Left            =   12480
      Picture         =   "frmViewResultsWE.frx":BB0A
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   6585
      Width           =   1245
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Haematology Cumulative"
      Height          =   855
      Index           =   1
      Left            =   4035
      Picture         =   "frmViewResultsWE.frx":BF4C
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Biochemistry Cumulative"
      Height          =   885
      Index           =   0
      Left            =   12480
      Picture         =   "frmViewResultsWE.frx":C38E
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   1545
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Height          =   5820
      Left            =   60
      TabIndex        =   15
      Top             =   2115
      Width           =   3915
      Begin VB.TextBox tnrbcP 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   3780
         Width           =   825
      End
      Begin VB.TextBox tnrbcA 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   3780
         Width           =   825
      End
      Begin VB.TextBox tRetP 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   120
         Top             =   5130
         Width           =   825
      End
      Begin VB.TextBox tPlt 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   38
         Top             =   4065
         Width           =   825
      End
      Begin VB.TextBox tPdw 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4350
         Width           =   825
      End
      Begin VB.TextBox tMPV 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4350
         Width           =   825
      End
      Begin VB.TextBox tPLCR 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4065
         Width           =   825
      End
      Begin VB.TextBox tMCV 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   34
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox tRDWSD 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   33
         Top             =   570
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox tRDWCV 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox tMCH 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   31
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox tHct 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   30
         Top             =   1170
         Width           =   825
      End
      Begin VB.TextBox tHgb 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   29
         Top             =   570
         Width           =   825
      End
      Begin VB.TextBox tRBC 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   28
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox tMCHC 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   27
         Top             =   1170
         Width           =   825
      End
      Begin VB.TextBox tLymP 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   26
         Top             =   2160
         Width           =   825
      End
      Begin VB.TextBox tMonoA 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   25
         Top             =   2445
         Width           =   825
      End
      Begin VB.TextBox tLymA 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   24
         Top             =   2160
         Width           =   825
      End
      Begin VB.TextBox tBasP 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   23
         Top             =   3300
         Width           =   825
      End
      Begin VB.TextBox tMonoP 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   22
         Top             =   2445
         Width           =   825
      End
      Begin VB.TextBox tNeutP 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   21
         Top             =   2730
         Width           =   825
      End
      Begin VB.TextBox tEosA 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   20
         Top             =   3015
         Width           =   825
      End
      Begin VB.TextBox tWBC 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   19
         Top             =   1830
         Width           =   825
      End
      Begin VB.TextBox tNeutA 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   18
         Top             =   2730
         Width           =   825
      End
      Begin VB.TextBox tEosP 
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   17
         Top             =   3015
         Width           =   825
      End
      Begin VB.TextBox tBasA 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3300
         Width           =   825
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "mm/hr"
         Height          =   195
         Left            =   1320
         TabIndex        =   130
         Top             =   4875
         Width           =   450
      End
      Begin VB.Label lblSickledex 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   780
         TabIndex        =   129
         Top             =   5460
         Width           =   855
      End
      Begin VB.Label lblMalaria 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3015
         TabIndex        =   128
         Top             =   5430
         Width           =   825
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Malaria"
         Height          =   195
         Left            =   2430
         TabIndex        =   127
         Top             =   5520
         Width           =   510
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Sickledex"
         Height          =   195
         Left            =   60
         TabIndex        =   126
         Top             =   5505
         Width           =   690
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "nRBC %"
         Height          =   195
         Index           =   6
         Left            =   2355
         TabIndex        =   125
         Top             =   3825
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "nRBC"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   124
         Top             =   3825
         Width           =   420
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Retics %"
         Height          =   195
         Index           =   5
         Left            =   2325
         TabIndex        =   121
         Top             =   5220
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bas %"
         Height          =   195
         Index           =   4
         Left            =   2505
         TabIndex        =   102
         Top             =   3360
         Width           =   435
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Eos %"
         Height          =   195
         Index           =   3
         Left            =   2505
         TabIndex        =   101
         Top             =   3090
         Width           =   435
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Neut %"
         Height          =   195
         Index           =   2
         Left            =   2430
         TabIndex        =   100
         Top             =   2790
         Width           =   510
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mono %"
         Height          =   195
         Index           =   1
         Left            =   2370
         TabIndex        =   99
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lymph %"
         Height          =   195
         Index           =   0
         Left            =   2310
         TabIndex        =   98
         Top             =   2220
         Width           =   630
      End
      Begin VB.Image imgHaemGraphs 
         Height          =   480
         Left            =   3030
         Picture         =   "frmViewResultsWE.frx":C7D0
         ToolTipText     =   "Graphs for this Sample"
         Top             =   1560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblNotValid 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Results not yet available."
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   780
         TabIndex        =   63
         ToolTipText     =   "Sample in Progress"
         Top             =   1500
         Width           =   2745
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "ESR"
         Height          =   195
         Left            =   420
         TabIndex        =   62
         Top             =   4875
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Retics"
         Height          =   195
         Left            =   285
         TabIndex        =   61
         Top             =   5175
         Width           =   450
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Infectious Mono Screen"
         Height          =   375
         Left            =   1965
         TabIndex        =   60
         Top             =   4740
         Width           =   975
      End
      Begin VB.Label lesr 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   780
         TabIndex        =   59
         Top             =   4830
         Width           =   525
      End
      Begin VB.Label lmonospot 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3015
         TabIndex        =   58
         Top             =   4830
         Width           =   825
      End
      Begin VB.Label lretics 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   780
         TabIndex        =   57
         Top             =   5130
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Pdw"
         Height          =   195
         Left            =   2625
         TabIndex        =   56
         Top             =   4395
         Width           =   315
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MPV"
         Height          =   195
         Left            =   405
         TabIndex        =   55
         Top             =   4395
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Plt"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   54
         Top             =   4110
         Width           =   180
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "PLCR"
         Height          =   195
         Left            =   2520
         TabIndex        =   53
         Top             =   4110
         Width           =   420
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RDW CV"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   2280
         TabIndex        =   52
         Top             =   285
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MCHC"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   2475
         TabIndex        =   51
         Top             =   1215
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MCH"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   2580
         TabIndex        =   50
         Top             =   915
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hct"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   49
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MCV"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   48
         Top             =   915
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RBC"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   47
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RDW SD"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   2265
         TabIndex        =   46
         Top             =   615
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Hgb"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   405
         TabIndex        =   45
         Top             =   585
         Width           =   345
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Neut"
         Height          =   195
         Left            =   375
         TabIndex        =   44
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Mono"
         Height          =   195
         Left            =   315
         TabIndex        =   43
         Top             =   2490
         Width           =   420
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Lymph"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   2220
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "WBC"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   41
         Top             =   1890
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Eos"
         Height          =   195
         Left            =   465
         TabIndex        =   40
         Top             =   3030
         Width           =   270
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Bas"
         Height          =   195
         Left            =   465
         TabIndex        =   39
         Top             =   3300
         Width           =   270
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13080
      Top             =   3825
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   270
      Width           =   9405
      Begin VB.Label lblAandE 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   107
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "A && E"
         Height          =   195
         Left            =   480
         TabIndex        =   106
         Top             =   540
         Width           =   555
      End
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
         Left            =   150
         TabIndex        =   64
         Top             =   840
         Width           =   8625
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   2445
         TabIndex        =   12
         Top             =   540
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   7740
         TabIndex        =   11
         Top             =   210
         Width           =   285
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2775
         TabIndex        =   10
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8055
         TabIndex        =   9
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   3915
         TabIndex        =   8
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Index           =   1
         Left            =   6060
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblChartTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "NOPAS"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   2325
         TabIndex        =   5
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4515
         TabIndex        =   4
         Top             =   510
         Width           =   4260
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6435
         TabIndex        =   3
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   2
         Top             =   210
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
         Left            =   2775
         TabIndex        =   1
         Top             =   210
         Width           =   3240
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gCoag 
      Height          =   3075
      Left            =   5220
      TabIndex        =   14
      Top             =   6330
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5424
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Parameter                  |<Result       |<Ranges    |<Units    |<Source             |<S |<Code"
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
   Begin MSFlexGridLib.MSFlexGrid grdDemog 
      Height          =   1770
      Left            =   0
      TabIndex        =   77
      Top             =   9585
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   3122
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      FormatString    =   $"frmViewResultsWE.frx":CC12
   End
   Begin VB.CommandButton cmdSignOffBio 
      Caption         =   "Sign OFF"
      Height          =   885
      Left            =   12480
      Picture         =   "frmViewResultsWE.frx":CCA6
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   3435
      Width           =   1245
   End
   Begin VB.Label lblFasting 
      Caption         =   "Fasting"
      Height          =   195
      Left            =   2610
      TabIndex        =   136
      Top             =   8595
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image imgFilm 
      Height          =   780
      Left            =   4215
      Picture         =   "frmViewResultsWE.frx":D570
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblPregnancyComments 
      Caption         =   "All references ranges including Pregnancy reference ranges && Units are available in Hospital/External Laboratory user manuals"
      Height          =   375
      Left            =   90
      TabIndex        =   135
      Top             =   1485
      Width           =   4875
   End
   Begin VB.Image imgGreyTick 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   9660
      Picture         =   "frmViewResultsWE.frx":E04E
      Top             =   1020
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   9660
      Picture         =   "frmViewResultsWE.frx":E324
      Top             =   750
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   9660
      Picture         =   "frmViewResultsWE.frx":E5FA
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   90
      TabIndex        =   112
      Top             =   1890
      Width           =   1110
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Coagulation"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5265
      TabIndex        =   111
      Top             =   6060
      Width           =   840
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   180
      TabIndex        =   105
      Top             =   8355
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   104
      Top             =   8115
      Width           =   735
   End
   Begin VB.Label lblTimeTaken 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   76
      Top             =   8010
      Width           =   1365
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Time Taken"
      Height          =   195
      Left            =   1665
      TabIndex        =   75
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Run Date"
      Height          =   195
      Index           =   0
      Left            =   1845
      TabIndex        =   73
      Top             =   8295
      Width           =   690
   End
   Begin VB.Label lblRunDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   72
      Top             =   8265
      Width           =   1365
   End
   Begin VB.Image imgLatest 
      Height          =   285
      Left            =   2625
      Picture         =   "frmViewResultsWE.frx":E8D0
      Stretch         =   -1  'True
      ToolTipText     =   "View Most Recent Record"
      Top             =   9135
      Width           =   435
   End
   Begin VB.Image imgNext 
      Height          =   285
      Left            =   2175
      Picture         =   "frmViewResultsWE.frx":EBDA
      Stretch         =   -1  'True
      ToolTipText     =   "View Next Record"
      Top             =   9135
      Width           =   435
   End
   Begin VB.Image imgPrevious 
      Height          =   285
      Left            =   1305
      Picture         =   "frmViewResultsWE.frx":EEE4
      Stretch         =   -1  'True
      ToolTipText     =   "View Previous Record"
      Top             =   9135
      Width           =   435
   End
   Begin VB.Image imgEarliest 
      Height          =   285
      Left            =   855
      Picture         =   "frmViewResultsWE.frx":F1EE
      Stretch         =   -1  'True
      ToolTipText     =   "View Earliest Record"
      Top             =   9135
      Width           =   435
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
      Left            =   855
      TabIndex        =   71
      Top             =   8835
      Width           =   2205
   End
End
Attribute VB_Name = "frmViewResultsWE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Activated As Boolean

Private Sub CheckCumulative()

      Dim tb As Recordset
      Dim sqlBase As String
      Dim sql As String
      Dim n As Integer
      Dim TotalCount As Long

10    On Error GoTo CheckCumulative_Error

20    For n = 0 To 2
30        cmdCum(n).Visible = False
40    Next

50    For n = 1 To grd.Rows - 1
60        If n > 1 Then sqlBase = sqlBase & " Union "
70        sqlBase = sqlBase & "SELECT COUNT(*) Tot FROM Demographics D JOIN @TableName@ R " & _
                    "ON D.SampleID = R.SampleID WHERE " & _
                    "PatName = '" & AddTicks(grd.TextMatrix(n, 2)) & "' "
80        If Trim$(grd.TextMatrix(n, 0)) <> "" Then
90            sqlBase = sqlBase & "AND Chart = '" & grd.TextMatrix(n, 0) & "' "
100       Else
110           sqlBase = sqlBase & "AND ( COALESCE(Chart, '') = '' ) "
120       End If
130       If IsDate(grd.TextMatrix(n, 1)) Then
140           sqlBase = sqlBase & "AND DoB = '" & Format$(grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
150       Else
160           sqlBase = sqlBase & "AND ( COALESCE(DoB, '') = '' ) "
170       End If
180   Next n

      'sqlBase = "SELECT COUNT(*) Tot FROM Demographics D JOIN @TableName@ R " & _
       '          "ON D.SampleID = R.SampleID WHERE " & _
       '          "PatName = '" & AddTicks(lblName) & "' "
      'If Trim$(lblChart) <> "" Then
      '    sqlBase = sqlBase & "AND Chart = '" & lblChart & "' "
      'Else
      '    sqlBase = sqlBase & "AND ( COALESCE(Chart, '') = '' ) "
      'End If
      'If IsDate(lblDoB) Then
      '    sqlBase = sqlBase & "AND DoB = '" & Format$(lblDoB, "dd/mmm/yyyy") & "' "
      'Else
      '    sqlBase = sqlBase & "AND ( COALESCE(DoB, '') = '' ) "
      'End If

190   TotalCount = 0
200   sql = Replace(sqlBase, "@TableName@", "BioResults")
210   For n = 0 To intOtherHospitalsInGroup
220       Set tb = New Recordset
230       RecOpenClient n, tb, sql
240       While Not tb.EOF
250           TotalCount = TotalCount + tb!Tot
260           tb.MoveNext
270       Wend
280       If TotalCount > 1 Then
290           cmdCum(0).Visible = True
300           Exit For
310       End If

320   Next

330   TotalCount = 0
340   sql = Replace(sqlBase, "@TableName@", "HaemResults")
350   For n = 0 To intOtherHospitalsInGroup
360       Set tb = New Recordset
370       RecOpenServer n, tb, sql
380       While Not tb.EOF
390           TotalCount = TotalCount + tb!Tot
400           tb.MoveNext
410       Wend
420       If TotalCount > 1 Then
430           cmdCum(1).Visible = True
440           Exit For
450       End If
460   Next

470   TotalCount = 0
480   sql = Replace(sqlBase, "@TableName@", "CoagResults")
490   For n = 0 To intOtherHospitalsInGroup
500       Set tb = New Recordset
510       RecOpenServer n, tb, sql
520       While Not tb.EOF
530           TotalCount = TotalCount + tb!Tot
540           tb.MoveNext
550       Wend
560       If TotalCount > 1 Then
570           cmdCum(2).Visible = True
580           Exit For
590       End If
600   Next

610   Exit Sub

CheckCumulative_Error:

      Dim strES As String
      Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "frmViewResultsWE", "CheckCumulative", intEL, strES, sql

End Sub

Private Sub CheckExternal(ByVal Cn As Integer)

      Dim sql As String
      Dim tb As Recordset
      Dim ExtList As String
      Dim n As Integer

10    On Error GoTo CheckExternal_Error

20    lstExt.Clear
      '30    cmdExternal.Visible = False

30    ExtList = ""
40    For n = 1 To grdDemog.Rows - 1
50        If grdDemog.TextMatrix(n, 0) <> "" Then
60            ExtList = ExtList & "'" & grdDemog.TextMatrix(n, 0) & "', "
70        End If
80    Next
90    If ExtList = "" Then Exit Sub
100   ExtList = Left$(ExtList, Len(ExtList) - 2)
110   If ExtList = "''" Then Exit Sub
'120   SQL = "SELECT DISTINCT SampleID FROM MedibridgeResults " & _
'            "WHERE CAST(SampleID AS VARCHAR(100)) IN (" & ExtList & ") " & _
'            "UNION "
         sql = "SELECT DISTINCT SampleID FROM ExtResults " & _
            "WHERE CAST(SampleID AS VARCHAR(100)) IN (" & ExtList & ") "
130   Set tb = New Recordset
140   RecOpenClient Cn, tb, sql
150   If Not tb.EOF Then
160       cmdExternal.Visible = True
170       Do While Not tb.EOF
180           lstExt.AddItem tb!SampleID
190           tb.MoveNext
200       Loop
210   End If

220   Exit Sub

CheckExternal_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmViewResultsWE", "CheckExternal", intEL, strES, sql


End Sub

Private Sub CheckMicro(ByVal Cn As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim TotalCount As Long


10    On Error GoTo CheckMicro_Error

20    For n = 1 To grd.Rows - 1
30        If n > 1 Then sql = sql & " Union "
40        sql = sql & "Select Count (*) as Tot from Demographics where " & _
                "PatName = '" & AddTicks(grd.TextMatrix(n, 2)) & "' "

50        If lblChart <> "" Then
60            sql = sql & "and Chart = '" & grd.TextMatrix(n, 0) & "' "
70        Else
80            sql = sql & "and ( Chart is null or Chart = '' ) "
90        End If

100       If IsDate(grd.TextMatrix(n, 1)) Then
110           sql = sql & "and DoB = '" & Format(grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
120       Else
130           sql = sql & "and ( DoB is null or DoB = '' ) "
140       End If
150       sql = sql & "and SampleID > '" & sysOptMicroOffsetOLD(Cn) & "' "

160   Next n


      'sql = "Select Count (*) as Tot from Demographics where " & _
       '      "PatName = '" & AddTicks(lblName) & "' "
      '
      'If lblChart <> "" Then
      '    sql = sql & "and Chart = '" & lblChart & "' "
      'Else
      '    sql = sql & "and ( Chart is null or Chart = '' ) "
      'End If
      '
      'If IsDate(lblDoB) Then
      '    sql = sql & "and DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
      'Else
      '    sql = sql & "and ( DoB is null or DoB = '' ) "
      'End If
      'sql = sql & "and SampleID > '" & sysOptMicroOffset(Cn) & "'"

170   TotalCount = 0
180   Set tb = New Recordset
190   RecOpenServer Cn, tb, sql
200   While Not tb.EOF
210       TotalCount = TotalCount + tb!Tot
220       tb.MoveNext
230   Wend
240   If TotalCount > 0 Then
250       cmdMicro.Visible = True
260   End If

270   Exit Sub

CheckMicro_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmViewResultsWE", "CheckMicro", intEL, strES, sql


End Sub

Private Sub CheckTransfusion()

      Dim sqlCount As String
      Dim tb As Recordset
      Dim RecordCount As Integer

10    On Error GoTo CheckTransfusion_Error
        RecordCount = 0
20    cmdBloodGroup.Visible = False


30      sqlCount = "SELECT COUNT (DISTINCT Name) Tot " & _
        "FROM PatientDetails WHERE " & _
        "PATNUM = '" & lblChart & "'"
        'MsgBox (sqlCount)
40      Set tb = New Recordset
50      RecOpenServerBB 0, tb, sqlCount
60      RecordCount = tb!Tot
        'MsgBox (RecordCount)
        If RecordCount > 0 Then
70        cmdBloodGroup.Visible = True
80        End If


90    Exit Sub

CheckTransfusion_Error:

      Dim strES As String
      Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "frmViewResultsWE", "CheckTransfusion", intEL, strES, sqlCount

End Sub

Private Sub ClearBloodGas()

10    txtBE = ""
20    txtHco3 = ""
30    txtO2Sat = ""
40    txtPco2 = ""
50    txtPh = ""
60    txtPo2 = ""
70    txtTotCo2 = ""
80    lblBGAComment = ""

90    fraBGA.Font.Bold = False
100   fraBGA.ForeColor = vbBlack

End Sub


Private Sub FillDemographics()

10    On Error GoTo FillDemographics_Error

20    With grdDemog

30        lblChart = .TextMatrix(.Row, 10)
40        lblSampleID = .TextMatrix(.Row, 0)
50        lblAandE = .TextMatrix(.Row, 9)
60        lblName = .TextMatrix(.Row, 2)
70        lblDoB = .TextMatrix(.Row, 3)
80        lblAge = .TextMatrix(.Row, 4)
90        If IsDate(.TextMatrix(.Row, 6)) Then
100           lblRunDate = Format$(.TextMatrix(.Row, 5), "dd/mm/yyyy")
110       Else
120           lblRunDate = "Not Specified"
130       End If
140       If IsDate(.TextMatrix(.Row, 6)) Then
150           If Format(.TextMatrix(.Row, 6)) <> "00:00" Then
160               lblTimeTaken = Format(.TextMatrix(.Row, 6), "dd/MM/yy hh:mm")
170           Else
180               lblTimeTaken = "Not Specified"
190           End If
200       Else
210           lblTimeTaken = "Not Specified"
220       End If
230       Select Case Left$(UCase$(.TextMatrix(.Row, 7)), 1)
          Case "M": lblSex = "Male"
240       Case "F": lblSex = "Female"
250       Case Else: lblSex = ""
260       End Select
270       lblAddress = .TextMatrix(.Row, 8)

280       If .Row = 1 Then
290           If .Rows = 2 Then
300               lblRecordInfo = "Record 1 of 1"
310               imgEarliest.Visible = False
320               imgPrevious.Visible = False
330               imgNext.Visible = False
340               imgLatest.Visible = False
350           Else
360               lblRecordInfo = "Most Recent Record."
370               imgEarliest.Visible = True
380               imgPrevious.Visible = True
390               imgNext.Visible = False
400               imgLatest.Visible = False
410           End If
420       ElseIf .Row = .Rows - 1 Then
430           lblRecordInfo = "Earliest Record."
440           imgEarliest.Visible = False
450           imgPrevious.Visible = False
460           imgNext.Visible = True
470           imgLatest.Visible = True
480       Else
490           lblRecordInfo = "Record " & .Rows - .Row & " of " & .Rows - 1
500           imgEarliest.Visible = True
510           imgPrevious.Visible = True
520           imgNext.Visible = True
530           imgLatest.Visible = True
540       End If

550       SetViewScans lblSampleID, cmdViewScan
560   End With

570   Exit Sub

FillDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

580   intEL = Erl
590   strES = Err.Description
600   LogError "frmViewResultsWE", "FillDemographics", intEL, strES


End Sub

Private Sub LoadAllResults()

      Dim Cn As Integer

10    On Error GoTo LoadAllResults_Error



20    cmdComments.Visible = False
30    imgFilm.Visible = False
40    cmdSemen.Visible = False
50    cmdMicro.Visible = False
60    cmdExternal.Visible = False

70    If sysOptDeptMicro(Cn) Then CheckMicro Cn
80    If sysOptDeptExt(Cn) Then CheckExternal Cn

90    If lblSampleID = "" Then Exit Sub
100   Cn = Val(grdDemog.TextMatrix(grdDemog.Row, 1))

'MsgBox Cn
110   LoadBiochemistry Cn
      '60    LoadImmunology Cn
120   LoadCoag Cn
130   LoadHaem Cn
140   LoadComments

150   If sysOptDeptBga(Cn) Then
160       LoadBloodGas Cn
170   End If


180   LogAsViewed "A", lblSampleID, lblChart
190   Exit Sub

LoadAllResults_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmViewResultsWE", "LoadAllResults", intEL, strES

End Sub
Private Sub LoadComments()

      Dim OBS As Observations
      Dim OB As Observation

10    On Error GoTo LoadComments_Error

20    lblDemogComment = ""

30    Set OBS = New Observations
40    Set OBS = OBS.Load(lblSampleID, "Demographic", "Biochemistry", "Haematology", "Coagulation", "Film", "Immunology", "Biomnis", "NVRL", "MATLAB", "Beaumont")

50    If Not OBS Is Nothing Then
60        cmdComments.Visible = True
70        imgFilm.Visible = False
80        For Each OB In OBS
90            If UCase(OB.Discipline) = "FILM" Then
100               imgFilm.Visible = True
110           End If
120       Next

130   End If




140   If CheckAutoComments(lblSampleID, 2) <> "" Then
150       cmdComments.Visible = True
160   End If
170   If CheckAutoComments(lblSampleID, 3) <> "" Then
180       cmdComments.Visible = True
190   End If

200   Exit Sub

LoadComments_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmViewResultsWE", "LoadComments", intEL, strES

End Sub

Private Sub LoadInitialDemographics()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer
      Dim S As String
      Dim ChartList As String
      Dim PatNameList As String
      Dim DoBList As String
      Dim WhereClause As String
      Dim SampleIDIndex As Integer
10    On Error GoTo LoadInitialDemographics_Error

20    SampleIDIndex = 1

30    For n = 1 To grd.Rows - 1
40        If n > 1 Then sql = sql & " Union "
          'Debug.Print sysOptMicroOffset(0)
50        sql = sql & "SELECT * FROM Demographics WHERE (sampleid < 2000000 OR sampleid > 2999999)  and "
          'sql = sql & "SELECT * FROM Demographics WHERE  "
60        sql = sql & "Chart = '" & grd.TextMatrix(n, 0) & "' "
70        sql = sql & "AND PatName = '" & AddTicks(grd.TextMatrix(n, 2)) & "' "
80        If IsDate(grd.TextMatrix(n, 1)) Then
90            sql = sql & "AND DoB = '" & Format(grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
100       Else
110           sql = sql & "AND (DoB = '' or DoB is Null) "
120       End If
130       sql = sql & "AND RunDate > '" & Format(Now - Val(frmMain.txtLookBack), "dd/MMM/yyyy") & "' "
140   Next n
      '130   sql = sql & "ORDER BY RunDate DESC,SampleID DESC"
150   sql = sql & "ORDER BY SampleDate DESC"



      'For n = 1 To grd.Rows - 1
      '    If n > 1 Then WhereClause = WhereClause & " OR "
      '    WhereClause = WhereClause & "("
      '    If grd.TextMatrix(n, 0) <> "" Then
      '        WhereClause = WhereClause & "Chart = '" & grd.TextMatrix(n, 0) & "' AND "
      '    End If
      '    WhereClause = WhereClause & "PatName = '" & grd.TextMatrix(n, 2) & "' AND "
      '    If IsDate(grd.TextMatrix(n, 1)) Then
      '        WhereClause = WhereClause & "DoB = '" & Format(grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
      '    Else
      '        WhereClause = WhereClause & "(DoB = '' or DoB is Null) "
      '    End If
      '    WhereClause = WhereClause & ") "
      '
      '
      '    ChartList = ChartList & "'" & grd.TextMatrix(n, 0) & "',"
      '    DoBList = DoBList & "'" & Format(grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "',"
      '    PatNameList = PatNameList & "'" & AddTicks(grd.TextMatrix(n, 2)) & "',"
      'Next n
      '
      'ChartList = Left$(ChartList, Len(ChartList) - 1)
      'DoBList = Left$(DoBList, Len(DoBList) - 1)
      'PatNameList = Left$(PatNameList, Len(PatNameList) - 1)
      '
      'sql = "SELECT * FROM Demographics WHERE " & WhereClause & "ORDER BY RunDate DESC, SampleID DESC"


      'If grd.Rows = 2 Then
      '    sql = "Select * from Demographics where " & _
           '          "Chart = '" & grd.TextMatrix(1, 0) & "' " & _
           '          "and PatName = '" & grd.TextMatrix(1, 2) & "' "
      '    If IsDate(grd.TextMatrix(1, 1)) Then
      '        sql = sql & "and DoB = '" & Format(grd.TextMatrix(1, 1), "dd/mmm/yyyy") & "' "
      '    Else
      '        sql = sql & "and (DoB = '' or DoB is Null) "
      '    End If
      '    'sql = sql & " AND SampleID < " & sysOptMicroOffset(0)
      '    sql = sql & "ORDER BY RunDate DESC,SampleID DESC"
      'Else
      '    sql = "Select * from Demographics where " & _
           '          "Chart In (" & ChartList & ") " & _
           '          "AND PatName In (" & PatNameList & ") " & _
           '          "AND DoB In (" & DoBList & ") "
      '    'sql = sql & " AND SampleID < " & sysOptMicroOffset(0)
      '    sql = sql & "ORDER BY RunDate DESC,SampleID DESC"
      'End If


160   For n = 0 To intOtherHospitalsInGroup
170       Set tb = New Recordset
180       RecOpenServer n, tb, sql
190       Do While Not tb.EOF
200           S = tb!SampleID & vbTab & _
                  Format$(n) & vbTab & _
                  tb!PatName & vbTab & _
                  tb!DoB & vbTab & _
                  tb!Age & vbTab & _
                  Format(tb!Rundate, "dd/mm/yy hh:mm") & vbTab & _
                  Format(tb!sampleDate, "dd/mm/yy hh:mm") & vbTab & _
                  tb!Sex & vbTab & _
                  tb!Addr0 & " " & tb!Addr1 & "" & vbTab & _
                  tb!AandE & "" & vbTab & _
                  tb!Chart & ""
              'If IsBioExist(tb!sampleid) Or IsHaemExist(tb!sampleid) Or IsCoagExist(tb!sampleid) Then
210           grdDemog.AddItem S
              'End If
220           tb.MoveNext
230       Loop
240   Next

250   With grdDemog
260       If .Rows > 2 Then
270           .RemoveItem 1
280           .col = 5
              '.Sort = 9
290       End If
300       lblRecordInfo = "Record " & .Rows - 1 & " of " & .Rows - 1
310       If .Rows = 2 Then
320           grdDemog.Row = 1
330           lblRecordInfo = "Record 1 of 1"
340           imgEarliest.Visible = False
350           imgPrevious.Visible = False
360           imgNext.Visible = False
370           imgLatest.Visible = False
380       Else
390           For n = 1 To .Rows - 1
400               If .TextMatrix(n, 0) = lblSampleID Then
410                   grdDemog.Row = n
420                   If n = 1 Then
430                       lblRecordInfo = "Most Recent Record."
440                       imgNext.Visible = False
450                       imgLatest.Visible = False

460                   ElseIf n = .Rows - 1 Then
470                       lblRecordInfo = "Earliest Record."
480                       imgEarliest.Visible = False
490                       imgPrevious.Visible = False
500                       imgNext.Visible = True
510                       imgLatest.Visible = True
520                   Else
530                       lblRecordInfo = "Record " & n & " of " & .Rows - 1
540                   End If
550                   SampleIDIndex = n
560                   Exit For
570               End If
580           Next

590       End If
600       lblChart = IIf(.TextMatrix(SampleIDIndex, 10) = "", lblChart, .TextMatrix(SampleIDIndex, 10))
610       lblSampleID = IIf(.TextMatrix(SampleIDIndex, 0) = "", "", .TextMatrix(SampleIDIndex, 0))
620       lblName = IIf(.TextMatrix(SampleIDIndex, 2) = "", lblName, .TextMatrix(SampleIDIndex, 2))
630       lblDoB = IIf(.TextMatrix(SampleIDIndex, 3) = "", lblDoB, .TextMatrix(SampleIDIndex, 3))
640       lblAge = .TextMatrix(SampleIDIndex, 4)
650       If IsDate(.TextMatrix(.Row, 5)) Then
660           lblRunDate = Format$(.TextMatrix(SampleIDIndex, 5), "dd/mm/yyyy")
670       Else
680           lblRunDate = "Not specified"
690       End If
700       lblTimeTaken = .TextMatrix(SampleIDIndex, 6)
710       Select Case Left$(UCase$(.TextMatrix(SampleIDIndex, 7)), 1)
          Case "M": lblSex = "Male"
720       Case "F": lblSex = "Female"
730       Case Else: lblSex = ""
740       End Select
750       lblAddress = .TextMatrix(SampleIDIndex, 8)
760       lblAandE = .TextMatrix(SampleIDIndex, 9)
770       SetViewScans lblSampleID, cmdViewScan
780   End With

790   Exit Sub

LoadInitialDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

800   intEL = Erl
810   strES = Err.Description
820   LogError "frmViewResultsWE", "LoadInitialDemographics", intEL, strES, sql


End Sub
Private Function IsHaemExist(strSampleID As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10       On Error GoTo IsHaemExist_Error

20    IsHaemExist = False

30    sql = "SELECT SampleID FROM HaemResults WHERE " & _
            "SampleID = '" & strSampleID & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      IsHaemExist = True
80    End If

90       Exit Function

IsHaemExist_Error:
      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmViewResultsWE", "IsHaemExist", intEL, strES

End Function
Private Function IsBioExist(strSampleID As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10       On Error GoTo IsBioExist_Error

20    IsBioExist = False

30    sql = "SELECT SampleID FROM bioResults WHERE " & _
            "SampleID = '" & strSampleID & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      IsBioExist = True
80    End If

90       Exit Function

IsBioExist_Error:
      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmViewResultsWE", "IsBioExist", intEL, strES

End Function

Private Function IsCoagExist(strSampleID As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10       On Error GoTo IsCoagExist_Error

20    IsCoagExist = False

30    sql = "SELECT SampleID FROM CoagResults WHERE " & _
            "SampleID = '" & strSampleID & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      IsCoagExist = True
80    End If

90       Exit Function

IsCoagExist_Error:
      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmViewResultsWE", "IsCoagExist", intEL, strES

End Function
Private Sub cmdBloodGroup_Click()
      Dim tb As Recordset
      Dim sqlDistinctPatient As String
      Dim sqlCount As String
      Dim tbLatest As Recordset
      Dim SQLLatest As String
      Dim f As Form
      Dim Cn As Integer
      Dim RecordCount As Long



10    On Error GoTo cmdBloodGroup_Click_Error

20    If Trim$(lblChart) = "" Then
30        Exit Sub
40    End If

      'For latest patient details
50    SQLLatest = "Select top 1 PATNUM, NAME, DOB, ADDR1,datetime " & _
                  "from PatientDetails where " & _
                  "PATNUM = '" & lblChart & "' order by SampleDate Desc"
      'For total record
60    sqlCount = "SELECT COUNT (DISTINCT Name) Tot " & _
                 "FROM PatientDetails WHERE " & _
                 "PATNUM = '" & lblChart & "'"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sqlCount
90    RecordCount = tb!Tot

      'for update the grid
100   sqlDistinctPatient = "SELECT PATNUM, NAME, DOB from PatientDetails where " & _
                           "PATNUM = '" & lblChart & "' GROUP BY PATNUM, NAME, DOB ORDER BY MAX(DateTime) DESC "
110   Set tb = New Recordset
120   RecOpenClientBB Cn, tb, sqlDistinctPatient
130   Set f = New frmConflictBT
140   With f.grd
150       .Rows = 2
160       .AddItem ""
170       .RemoveItem 1
180       Do Until tb.EOF
190           .AddItem tb!Patnum & vbTab & _
                       Format(tb!DoB, "dd/MMM/yyyy") & vbTab & _
                       tb!Name & ""
200           tb.MoveNext
210       Loop
220       If .Rows > 2 And .TextMatrix(1, 0) = "" Then
230           .RemoveItem 1
240       End If
250   End With
      ' for most recent Record
260   With f
270       .CountWarning = RecordCount
280       Set tbLatest = New Recordset
290       RecOpenServerBB 0, tbLatest, SQLLatest
300       If Not tbLatest.EOF Then
310           .RecentPatName = Trim$(tbLatest!Name & "")
320           .RecentChart = Trim$(tbLatest!Patnum & "")
330           .RecentDoB = Format$(tbLatest!DoB, "dd/MMM/yyyy")
340           .RecentDate = Format$(tbLatest!DateTime, "dd/MMM/yyyy")
350       End If
360       .Show 1
370   End With



380   Exit Sub

cmdBloodGroup_Click_Error:
      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "frmViewResultsWE", "cmdBloodGroup_Click", intEL, strES, sqlDistinctPatient & " " & SQLLatest & " " & sqlCount

End Sub

Private Sub cmdCancel_Click()

10    On Error GoTo cmdCancel_Click_Error

      Dim MicroSignOffRequired As Boolean
      Dim BloodSignOffRequired As Boolean
      Dim UserRoleCode As String

20    MicroSignOffRequired = False
30    BloodSignOffRequired = False

40    MicroSignOffRequired = CheckForUnsignedMicro(lblChart)

50    UserRoleCode = ListCodeFor("RL", UserRoleName)
60    If UCase(UserRoleCode) = "RLCON" Or UCase(UserRoleCode) = "RLNCHD" Or UCase(UserRoleCode) = "RLANP" Then

70        BloodSignOffRequired = SignOffRequired(gBio, 6)         'biochemistry
80        If BloodSignOffRequired = False Then
90            BloodSignOffRequired = SignOffRequired(gBiomnis, 6)         'external
100       End If
110       If BloodSignOffRequired = False Then
120           BloodSignOffRequired = SignOffRequired(gCoag, 5)          'coagulation
130       End If
140       If BloodSignOffRequired = False Then
150           BloodSignOffRequired = SignOffRequiredHaem(lblChart)
160       End If
170   End If



180   If MicroSignOffRequired Or BloodSignOffRequired Then

190       If MicroSignOffRequired Then
200           If iMsg(vbCrLf & "There are microbiology results that need signing off. " & vbCrLf & " " & vbCrLf & " Do you wish to sign off", vbYesNo, "Please confirm!") = vbYes Then
210               With frmMicroReport
220                   .ReportDept = "MICRO"
230                   .lblChart = lblChart
240                   .lblName = lblName
250                   .lblDoB = lblDoB
260                   .Show 1
270               End With
280           Else

290               AddActivity lblSampleID, "No to microbiology results that need signing off. ", "Ignored", "", lblChart, "", ""
300               If BloodSignOffRequired = False Then Unload Me
310           End If
320       End If
330       If BloodSignOffRequired Then
340           If iMsg(vbCrLf & "There are blood results that need signing off. " & vbCrLf & " " & vbCrLf & " Do you wish to sign off", vbYesNo, "Please confirm!") = vbNo Then
350               AddActivity lblSampleID, "No to blood results that need signing off. ", "Ignored", "", lblChart, "", ""
360               Unload Me
370           End If
380       End If


390   Else
400       Unload Me
410   End If
420   Exit Sub

cmdCancel_Click_Error:
      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmViewResultsWE", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdComments_Click()
 AddActivity lblSampleID, "Ward Enquiry Comment", "VIEWED", "", lblChart, "", ""
10    With frmComments
20        .SampleID = lblSampleID
30        .Show 1
40    End With

End Sub

Private Sub cmdCum_Click(Index As Integer)

10    Select Case Index
      Case 0:
20        With fFullBioWE
30            .lblChart = lblChart
40            .lblDoB = lblDoB
50            .lblName = lblName
60            .lblSex = lblSex
70            .Show 1
80        End With
90    Case 1:
100       With fCumHaemWE
110           .lblChart = lblChart
120           .lblDoB = lblDoB
130           .lblName = lblName
140           .lblSex = lblSex
150           .Show 1
160       End With
170   Case 2:
180       With fFullCoagWE
190           .lblChart = lblChart
200           .lblDoB = lblDoB
210           .lblName = lblName
220           .Show 1
230       End With
240   End Select

End Sub

Private Sub cmdExternal_Click()

      Dim medibridgepathtoviewer As String
      Dim SID As String
      Dim col As Collection
      Dim yL As Integer
      Dim yG As Integer
      Dim S As String
      Dim f As Form
      Dim n As Integer

10    On Error GoTo cmdExternal_Click_Error

      'CheckExternal 0
20    If lstExt.ListCount = 0 Then Exit Sub

30    If lstExt.ListCount = 0 Then
40        SID = lstExt.List(0)
50    Else
60        Set col = New Collection
70        For yL = lstExt.ListCount - 1 To 0 Step -1
80            S = lstExt.List(yL)
90            For yG = 1 To grdDemog.Rows - 1
100               If grdDemog.TextMatrix(yG, 0) = S Then
110                   S = S & " " & grdDemog.TextMatrix(yG, 6)
120                   col.Add S
130                   Exit For
140               End If
150           Next
160       Next
170       Set f = New frmChooseExternal
180       With f
190           Set .DateSID = col
200           .Show 1
210           SID = .SampleID
220       End With
230       Unload f
240       Set f = Nothing
250   End If

260   If SID <> "" Then
          '270       medibridgepathtoviewer = GetOptionSetting("MedibridgePathToViewer", "", "")
          '280       If medibridgepathtoviewer <> "" Then
          '290           Shell medibridgepathtoviewer & " /SampleID=" & SID & _
           '                    " /UserName=""" & UserName & """" & _
           '                    " /Password=""" & UserPass & """" & _
           '                    " /Department=Medibridge", vbNormalFocus
          '300       End If

270       For n = 1 To grdDemog.Rows - 1
280           If grdDemog.TextMatrix(n, 0) = SID Then
290               grdDemog.Row = n
300               FillDemographics
310               LoadAllResults


320               imgPrevious.Visible = True
330               imgEarliest.Visible = True
340               imgNext.Visible = True
350               imgLatest.Visible = True


360               If grdDemog.Row = 1 Then
370                   imgNext.Visible = False
380                   imgLatest.Visible = False
390               ElseIf grdDemog.Row = grdDemog.Rows - 1 Then
400                   imgEarliest.Visible = False
410                   imgPrevious.Visible = False
420               End If

430               If SSTabBio.Tab = 1 And gBiomnis.TextMatrix(1, 0) <> "" Then
440                   AddActivity lblSampleID, "Ward Enquiry External Results", "VIEWED", "", lblChart, "", ""
450               End If

460               Exit For
470           End If
480       Next
490   End If

500   Exit Sub

cmdExternal_Click_Error:

      Dim strES As String
      Dim intEL As Integer

510   intEL = Erl
520   strES = Err.Description
530   LogError "frmViewResultsWE", "cmdExternal_Click", intEL, strES


End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdGreenTick_Click
' Author    : masood
' Date      : 11/1/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdGreenTick_Click(Index As Integer)
      Dim y As Integer

10    If SSTabBio.Tab = 0 Then
20        For y = 1 To gBio.Rows - 1
30            gBio.Row = y
40            If gBio.CellPicture <> imgGreyTick Then
50                Set gBio.CellPicture = imgGreenTick.Picture
60            End If
70        Next
80    ElseIf SSTabBio.Tab = 1 Then
90        For y = 1 To gBiomnis.Rows - 1
100           gBiomnis.Row = y
110           If gBiomnis.CellPicture <> imgGreyTick Then
120               Set gBiomnis.CellPicture = imgGreenTick.Picture
130           End If
140       Next

150   End If
End Sub



Private Sub cmdMicro_Click()

10    With frmMicroReport
20        .ReportDept = "MICRO"
30        .lblChart = lblChart
40        .lblName = lblName
50        .lblDoB = lblDoB
60        .Show 1
70    End With

End Sub

Private Sub cmdOrder_Click()

10    With frmOrderComms

20        .lblSampleID = lblSampleID
30        .lblName = lblName
40        .lblDoB = lblDoB
50        .lblAge = lblAge
60        .lblChart = lblChart
70        .lblSex = lblSex
80        .lblAddress = lblAddress
90        .lblDemogComment = lblDemogComment

100       .Show 1

110   End With

End Sub

Private Sub cmdPrint_Click(Index As Integer)

      Dim sql As String
      Dim tb As Recordset
      Dim Cn As Integer
      Dim Dept As String
      Dim PrintDept As String
      Dim InhibitDept As String

10    On Error GoTo cmdPrint_Click_Error

20    If DateDiff("d", GetOptionSetting("WardEnqV7Date", "01/May/2011", ""), lblRunDate) > 0 Then
30        With frmReportViewer
40            .SampleID = lblSampleID
50            .InhibitChoosePrinter = True
60            .PrintToPrinter = WardEnqForcedPrinter
70            .Show 1
80        End With
90        Exit Sub
100   End If

110   Select Case Index
      Case 0: Dept = "B": PrintDept = "I": InhibitDept = "Bio"
120   Case 1: Dept = "H": PrintDept = "J": InhibitDept = "Haem"
130   Case 2: Dept = "C": PrintDept = "K": InhibitDept = "Coa"
140   End Select

150   Cn = Val(grdDemog.TextMatrix(grdDemog.Row, 1))

160   If iMsg("Report will be printed on" & vbCrLf & _
              WardEnqForcedPrinter & "." & vbCrLf & _
              "OK?", vbQuestion + vbYesNo) = vbYes Then

170       cmdPrint(Index).Caption = "&Printing..."

180       sql = "DELETE FROM PrintInhibit WHERE " & _
                "SampleID = '" & lblSampleID & "' " & _
                "AND Discipline = '" & InhibitDept & "'"
190       Cnxn(0).Execute sql

200       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = '" & Dept & "' " & _
                "AND SampleID = '" & lblSampleID & "'"
210       Set tb = New Recordset
220       RecOpenClient 0, tb, sql
230       If tb.EOF Then
240           tb.AddNew
250       End If
260       tb!SampleID = lblSampleID
270       tb!Department = Dept
280       If Trim$(UserName) <> "" Then
290           tb!Initiator = UserName
300       Else
310           tb!Initiator = "Ward"
320       End If
330       tb!UsePrinter = WardEnqForcedPrinter
340       tb!ThisIsCopy = 1
350       tb.Update

360       LogAsViewed PrintDept, lblSampleID, lblChart
370       cmdPrint(Index).Caption = "&Print"

380   End If

390   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "frmViewResultsWE", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdRedCross_Click(Index As Integer)
      Dim y As Integer
10    If SSTabBio.Tab = 0 Then
20        For y = 1 To gBio.Rows - 1
30            gBio.Row = y
40            If gBio.CellPicture <> imgGreyTick Then
50                Set gBio.CellPicture = imgRedCross.Picture
60            End If
70        Next
80    ElseIf SSTabBio.Tab = 1 Then
90        For y = 1 To gBiomnis.Rows - 1
100           gBiomnis.Row = y
110           If gBiomnis.CellPicture <> imgGreyTick Then
120               Set gBiomnis.CellPicture = imgRedCross.Picture
130           End If
140       Next
150   End If
End Sub


Private Sub cmdSemen_Click()
10    With frmMicroReport
20        .ReportDept = "SEMEN"
30        .lblChart = lblChart
40        .lblName = lblName
50        .lblDoB = lblDoB
60        .Show 1
70    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSignOffBio_Click
' Author    : Masood
' Date      : 11/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSignOffBio_Click()
10    On Error GoTo cmdSignOffBio_Click_Error

      Dim sql As String
      Dim i As Integer



20    With gBio
30        For i = 1 To .Rows - 1

40            .Row = i
50            .col = 6
60            If .CellPicture = imgGreenTick.Picture Then
70                sql = " UPDATE BioResults SET SignOff = 1 , "
80                sql = sql & " SignOffBy = '" & UserName & "'"
90                sql = sql & " , SignOffDateTime ='" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "'"
100               sql = sql & " WHERE SampleID = '" & lblSampleID & "'"
110               sql = sql & " AND CODE = '" & .TextMatrix(i, 7) & "'"
120               Cnxn(0).Execute sql
130               Set .CellPicture = imgGreyTick.Picture
140           End If
150       Next i
160   End With


170   With gBiomnis
180       For i = 1 To .Rows - 1

190           .Row = i
200           .col = 6
210           If .CellPicture = imgGreenTick.Picture Then
220               sql = " UPDATE BioResults SET SignOff = 1 , "
230               sql = sql & " SignOffBy = '" & UserName & "'"
240               sql = sql & " , SignOffDateTime ='" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "'"
250               sql = sql & " WHERE SampleID = '" & lblSampleID & "'"
260               sql = sql & " AND CODE = '" & .TextMatrix(i, 7) & "'"
270               Cnxn(0).Execute sql
280               Set .CellPicture = imgGreyTick.Picture
290           End If
300       Next i
310   End With
320   Call SignOffEnable(gBiomnis, 6, cmdSignOffBio)
330   Call SignOffEnable(gBio, 6, cmdSignOffBio)

340   Exit Sub


cmdSignOffBio_Click_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "frmViewResultsWE", "cmdSignOffBio_Click", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSignOff_Click
' Author    : Masood
' Date      : 10/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSignOff_Click()
10    On Error GoTo cmdSignOff_Click_Error
      Dim sql As String

20    sql = " UPDATE HaemResults SET SignOff = 1 , "
30    sql = sql & " SignOffBy = '" & UserName & "'"
40    sql = sql & " , SignOffDateTime ='" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "'"
50    sql = sql & " WHERE SampleID = '" & lblSampleID & "'"
60    Cnxn(0).Execute sql

70    sql = " UPDATE HaemResults50 SET SignOff = 1 , "
80    sql = sql & " SignOffBy = '" & UserName & "'"
90    sql = sql & " , SignOffDateTime ='" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "'"
100   sql = sql & " WHERE SampleID = '" & lblSampleID & "'"
110   Cnxn(0).Execute sql

120   cmdSignOff.Enabled = False



130   Exit Sub


cmdSignOff_Click_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmViewResultsWE", "cmdSignOff_Click", intEL, strES, sql
End Sub

Private Sub cmdSignOffCoag_Click()
      Dim sql As String
      Dim i As Integer

10    With gCoag
20        For i = 1 To .Rows - 1

30            .Row = i
40            .col = 5
50            If .CellPicture = imgGreenTick.Picture Then
60                sql = " UPDATE CoagResults SET SignOff = 1 , "
70                sql = sql & " SignOffBy = '" & UserName & "'"
80                sql = sql & " , SignOffDateTime ='" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "'"
90                sql = sql & " WHERE SampleID = '" & lblSampleID & "'"
100               sql = sql & " AND CODE = '" & .TextMatrix(i, 6) & "'"
110               Cnxn(0).Execute sql
120               Set .CellPicture = imgGreyTick.Picture
130           End If
140       Next i
150   End With
160   Call SignOffEnable(gCoag, 5, cmdSignOffCoag)
End Sub

Private Sub cmdViewScan_Click()
10    frmViewScan.CallerDepartment = "WardEnq Bio"
20    frmViewScan.SampleID = lblSampleID
30    frmViewScan.txtSampleID = lblSampleID
40    frmViewScan.Show 1
End Sub

Private Sub cmdViewAcreditationStatus_Click()
      Dim PathToDoc As String

10    On Error GoTo cmdViewAcreditationStatus_Click_Error


20    PathToDoc = App.Path & "\Current Scope of Accreditation.pdf"

30    ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5

40    Exit Sub

cmdViewAcreditationStatus_Click_Error:
      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmViewResultsWE", "cmdViewAcreditationStatus_Click", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Activate
' Author    : Masood
' Date      : 11/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : Form_Activate
' Author    : Masood
' Date      : 12/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error


20    If LogOffNow Then
30        Unload Me
40    End If

50    PBar.Max = LogOffDelaySecs
60    PBar = 0
70    SingleUserUpdateLoggedOn UserName

80    Timer1.Enabled = True

90    If Activated Then Exit Sub
100   Activated = True

110   LoadInitialDemographics
      'CheckExternal 0
120   CheckCumulative
130   LoadAllResults
140   CheckTransfusion
150   gBio.ColWidth(7) = 0
160   gCoag.ColWidth(6) = 0
170   gBiomnis.ColWidth(7) = 0
180   gBiomnis.ColWidth(8) = 0
190   gBio.ColWidth(8) = 0
200   Exit Sub


Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmViewResultsWE", "Form_Activate", intEL, strES


End Sub
Private Sub LoadHaem(ByVal Cn As Integer)

    Dim tb As Recordset
    Dim sql As String
    Dim sqlSample As String
    Dim stb As Recordset
    Dim sampleDate As String
   

10  On Error GoTo LoadHaem_Error
    sqlSample = "SELECT Sampledate from demographics WHERE SampleID = '" & Trim(lblSampleID) & "'"
81  Set stb = New Recordset
91  RecOpenClient Cn, stb, sqlSample
    sampleDate = stb!sampleDate

20  If Trim$(lblSampleID) = "" Then Exit Sub

30  ClearHaem
40  imgHaemGraphs.Visible = False
50  lblNotValid.Visible = False
60  'cmdPrint(1).Visible = False

70  sql = "Select * from HaemResults where " & _
          "SampleID = '" & lblSampleID & "'"
80  Set tb = New Recordset
90  RecOpenClient Cn, tb, sql
100 If Not tb.EOF Then
110     If IsNull(tb!Valid) Or tb!Valid = 0 Then
120         lblNotValid.Visible = True
130         imgHaemGraphs.Visible = False
140     Else
150         If Not IsNull(tb!gwb1) Or _
               Not IsNull(tb!gwb2) Or _
               Not IsNull(tb!gRBC) Or _
               Not IsNull(tb!gplt) Or _
               Not IsNull(tb!gplth) Then
160             imgHaemGraphs.Visible = True
170         End If

180         If Not IsNull(tb!rbc) Then
190             Colourise "RBC", tRBC, tb!rbc, lblSex, lblDoB, sampleDate
200         End If

210         If Not IsNull(tb!Hgb) Then
220             Colourise "Hgb", tHgb, tb!Hgb, lblSex, lblDoB, sampleDate
230         End If

240         If Not IsNull(tb!MCV) Then
250             Colourise "MCV", tMCV, tb!MCV, lblSex, lblDoB, sampleDate
260         End If

270         If Not IsNull(tb!Hct) Then
280             Colourise "Hct", tHct, tb!Hct, lblSex, lblDoB, sampleDate
290         End If

300         If Not IsNull(tb!RDWCV) Then
310             Colourise "RDWCV", tRDWCV, tb!RDWCV, lblSex, lblDoB, sampleDate
320         End If

330         If Not IsNull(tb!rdwsd) Then
340             Colourise "RDWSD", tRDWSD, tb!rdwsd, lblSex, lblDoB, sampleDate
350         End If

360         If Not IsNull(tb!mch) Then
370             Colourise "MCH", tMCH, tb!mch, lblSex, lblDoB, sampleDate
380         End If

390         If Not IsNull(tb!mchc) Then
400             Colourise "MCHC", tMCHC, tb!mchc, lblSex, lblDoB, sampleDate
410         End If

420         If Not IsNull(tb!plt) Then
430             Colourise "plt", tPlt, tb!plt, lblSex, lblDoB, sampleDate
440         End If

450         If Not IsNull(tb!mpv) Then
460             Colourise "MPV", tMPV, tb!mpv, lblSex, lblDoB, sampleDate
470         End If

480         If Not IsNull(tb!plcr) Then
490             Colourise "PLCR", tPLCR, tb!plcr, lblSex, lblDoB, sampleDate
500         End If

510         If Not IsNull(tb!pdw) Then
520             Colourise "Pdw", tPdw, tb!pdw, lblSex, lblDoB, sampleDate
530         End If

540         If Not IsNull(tb!WBC) Then
550             Colourise "WBC", tWBC, tb!WBC, lblSex, lblDoB, sampleDate
560         End If

570         If Not IsNull(tb!LymA) Then
580             Colourise "LymA", tLymA, tb!LymA, lblSex, lblDoB, sampleDate
590         End If

600         If Not IsNull(tb!LymP) Then
610             Colourise "LymP", tLymP, tb!LymP, lblSex, lblDoB, sampleDate
620         End If

630         If Not IsNull(tb!MonoA) Then
640             Colourise "MonoA", tMonoA, tb!MonoA, lblSex, lblDoB, sampleDate
650         End If

660         If Not IsNull(tb!MonoP) Then
670             Colourise "MonoP", tMonoP, tb!MonoP, lblSex, lblDoB, sampleDate
680         End If

690         If Not IsNull(tb!NeutA) Then
700             Colourise "NeutA", tNeutA, tb!NeutA, lblSex, lblDoB, sampleDate
710         End If

720         If Not IsNull(tb!NeutP) Then
730             Colourise "NeutP", tNeutP, tb!NeutP, lblSex, lblDoB, sampleDate
740         End If

750         If Not IsNull(tb!EosA) Then
760             Colourise "EosA", tEosA, tb!EosA, lblSex, lblDoB, sampleDate
770         End If

780         If Not IsNull(tb!EosP) Then
790             Colourise "EosP", tEosP, tb!EosP, lblSex, lblDoB, sampleDate
800         End If

810         If Not IsNull(tb!BasA) Then
820             Colourise "BasA", tBasA, tb!BasA, lblSex, lblDoB, sampleDate
830         End If

840         If Not IsNull(tb!BasP) Then
850             Colourise "BasP", tBasP, tb!BasP, lblSex, lblDoB, sampleDate
860         End If

870         lesr = tb!ESR & ""
880         If Trim(lesr) <> "" Then
                'lesr = lesr & " mm/hr "
890         End If

900         tRetP = tb!RetP & ""    ' tb!retics & ""
910         lretics = tb!RetA & ""
920         If Trim(lretics) <> "" Then
                '            lretics = lretics & " x10/l"
930         End If


940         lblSickledex = tb!Sickledex & ""
950         lblMalaria = tb!Malaria & ""

960         If tb!monospot & "" <> "" Then
961             If UCase(tb!monospot & "") = "P" Then
962                 lmonospot = "Positive"
963             ElseIf UCase(tb!monospot & "") = "N" Then
964                 lmonospot = "Negative"
965             End If
966         End If
970         tnrbcP = tb!nrbcP & ""
980         tnrbcA = tb!nrbcA & ""

990     End If

1000    If IsPrintable(lblSampleID, lblRunDate, "Haematology") Then
1010        cmdPrint(1).Visible = True
1020    End If

1030    If tb!Valid = 1 And (IsNull(tb!SignOff) Or tb!SignOff = 0) Then
1040        cmdSignOff.Enabled = True
1050    Else
1060        cmdSignOff.Enabled = False
1070    End If

1080 End If

1090 Screen.MousePointer = 0

1100 Exit Sub

LoadHaem_Error:

    Dim strES As String
    Dim intEL As Integer

1110 intEL = Erl
1120 strES = Err.Description
1130 LogError "frmViewResultsWE", "LoadHaem", intEL, strES, sql

End Sub
Private Sub Colourise(ByVal Analyte As String, _
                      ByVal Destination As TextBox, _
                      ByVal strValue As String, _
                      ByVal Sex As String, _
                      ByVal DoB As String, ByVal sampleDate As String)

      Dim Value As Single

10    Value = Val(strValue)

20    Destination.Text = strValue
30    If Trim$(strValue) = "" Then
40        Destination.BackColor = &HFFFFFF
50        Destination.ForeColor = &H0&
60        Exit Sub
70    End If

80    Select Case InterpH(Value, Analyte, Sex, DoB, sampleDate)
      Case "X":
90        Destination.BackColor = vbBlack
100       Destination.ForeColor = vbWhite
110   Case "H":
120       Destination.BackColor = sysOptHighBack(0)
130       Destination.ForeColor = sysOptHighFore(0)
140   Case "L"
150       Destination.BackColor = sysOptLowBack(0)
160       Destination.ForeColor = sysOptLowFore(0)
170   Case Else
180       Destination.BackColor = &HFFFFFF
190       Destination.ForeColor = &H0&
200   End Select

End Sub

Private Sub ClearHaem()

10    tWBC = ""
20    tWBC.BackColor = &HFFFFFF
30    tWBC.ForeColor = &H0&

40    tRBC = ""
50    tRBC.BackColor = &HFFFFFF
60    tRBC.ForeColor = &H0&

70    tHgb = ""
80    tHgb.BackColor = &HFFFFFF
90    tHgb.ForeColor = &H0&

100   tMCV = ""
110   tMCV.BackColor = &HFFFFFF
120   tMCV.ForeColor = &H0&

130   tHct = ""
140   tHct.BackColor = &HFFFFFF
150   tHct.ForeColor = &H0&

160   tRDWCV = ""
170   tRDWCV.BackColor = &HFFFFFF
180   tRDWCV.ForeColor = &H0&

190   tRDWSD = ""
200   tRDWSD.BackColor = &HFFFFFF
210   tRDWSD.ForeColor = &H0&

220   tMCH = ""
230   tMCH.BackColor = &HFFFFFF
240   tMCH.ForeColor = &H0&

250   tMCHC = ""
260   tMCHC.BackColor = &HFFFFFF
270   tMCHC.ForeColor = &H0&

280   tPlt = ""
290   tPlt.BackColor = &HFFFFFF
300   tPlt.ForeColor = &H0&

310   tMPV = ""
320   tMPV.BackColor = &HFFFFFF
330   tMPV.ForeColor = &H0&

340   tPLCR = ""
350   tPLCR.BackColor = &HFFFFFF
360   tPLCR.ForeColor = &H0&

370   tPdw = ""
380   tPdw.BackColor = &HFFFFFF
390   tPdw.ForeColor = &H0&

400   tLymA = ""
410   tLymA.BackColor = &HFFFFFF
420   tLymA.ForeColor = &H0&

430   tLymP = ""
440   tLymP.BackColor = &HFFFFFF
450   tLymP.ForeColor = &H0&

460   tMonoA = ""
470   tMonoA.BackColor = &HFFFFFF
480   tMonoA.ForeColor = &H0&

490   tMonoP = ""
500   tMonoP.BackColor = &HFFFFFF
510   tMonoP.ForeColor = &H0&

520   tNeutA = ""
530   tNeutA.BackColor = &HFFFFFF
540   tNeutA.ForeColor = &H0&

550   tNeutP = ""
560   tNeutP.BackColor = &HFFFFFF
570   tNeutP.ForeColor = &H0&

580   tEosA = ""
590   tEosA.BackColor = &HFFFFFF
600   tEosA.ForeColor = &H0&

610   tEosP = ""
620   tEosP.BackColor = &HFFFFFF
630   tEosP.ForeColor = &H0&

640   tBasA = ""
650   tBasA.BackColor = &HFFFFFF
660   tBasA.ForeColor = &H0&

670   tBasP = ""
680   tBasP.BackColor = &HFFFFFF
690   tBasP.ForeColor = &H0&

700   lesr = ""
710   lretics = ""
720   lmonospot = ""
730   cmdSignOff.Enabled = False
End Sub

Private Sub LoadCoag(ByVal Cn As Integer)

      Dim Cxs As New CoagResults
      Dim Cx As CoagResult
      Dim S As String
      Dim FormatStr As String
      Dim i As Integer
10    On Error GoTo LoadCoag_Error

20    gCoag.Visible = False
30    gCoag.Rows = 2
40    gCoag.AddItem ""
50    gCoag.RemoveItem 1

60    'cmdPrint(2).Visible = False

70    Set Cxs = Cxs.Load(lblSampleID, gDONTCARE, gDONTCARE, "Results", Cn)

80    If Cxs.Count <> 0 Then
90        For Each Cx In Cxs
100           S = Cx.TestName & vbTab
110           Select Case Cx.DP
              Case 0: FormatStr = "###0"
120           Case 1: FormatStr = "##0.0"
130           Case 2: FormatStr = "#0.00"
140           Case 3: FormatStr = "0.000"
150           End Select
160           If Cx.Valid Then
170               S = S & Format(Cx.Result, FormatStr) & vbTab & IIf(Cx.PrintRefRange = "1", Cx.Low & "-" & Cx.High, "") & vbTab & Cx.Units & vbTab & Cx.Analyser & vbTab & vbTab & Cx.Code
180           Else
190               S = S & "Not Valid"
200           End If
210           gCoag.AddItem S

220           With gCoag
230               .Row = .Rows - 1
240               .col = 5

250               If Cx.SignOff = False And Cx.Valid = True Then
260                   Set .CellPicture = imgRedCross.Picture
270               ElseIf Cx.SignOff = True And Cx.Valid = True Then
280                   Set .CellPicture = imgGreyTick.Picture
290               End If
300           End With


310           If Cx.Valid Then
320               If Cx.Low = 0 And (Cx.High = 999 Or Cx.High = 0 Or Cx.High = 9999) Then
330               Else
340                   If Val(Cx.Result) <> 0 Then
350                       If Val(Cx.Result) < Val(Cx.Low) Then
360                           gCoag.Row = gCoag.Rows - 1
370                           gCoag.col = 1
380                           gCoag.CellBackColor = sysOptLowBack(0)
390                           gCoag.CellForeColor = sysOptLowFore(0)
400                       ElseIf Val(Cx.Result) > Val(Cx.High) Then
410                           gCoag.Row = gCoag.Rows - 1
420                           gCoag.col = 1
430                           gCoag.CellBackColor = sysOptHighBack(0)
440                           gCoag.CellForeColor = sysOptHighFore(0)
450                       End If
460                   End If
470               End If
480           End If
490       Next
500   End If

510   LoadOutstandingCoag

520   If gCoag.Rows > 2 Then
530       gCoag.RemoveItem 1
540   End If

550   gCoag.Visible = True
560   For i = 0 To gCoag.Rows - 1
570       If Trim(gCoag.TextMatrix(i, 2)) = "0-9999" Or Trim(gCoag.TextMatrix(i, 2)) = "0-0" Then
580           gCoag.TextMatrix(i, 2) = ""
590       End If
600   Next
610   If IsPrintable(lblSampleID, lblRunDate, "Coagulation") Then
620       cmdPrint(2).Visible = True
630   End If
640   Call SignOffEnable(gCoag, 5, cmdSignOffCoag)
650   Exit Sub

LoadCoag_Error:

      Dim strES As String
      Dim intEL As Integer

660   intEL = Erl
670   strES = Err.Description
680   LogError "frmViewResultsWE", "LoadCoag", intEL, strES

End Sub

Private Sub LoadOutstandingCoag()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadOutstandingCoag_Error

20    sql = "Select Distinct D.TestName from CoagRequests as C, " & _
            "CoagTestDefinitions as D where " & _
            "C.SampleID = '" & lblSampleID & "' " & _
            "and D.Code = C.Code"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60        gCoag.AddItem tb!TestName & vbTab & "In Progress"
70        tb.MoveNext
80    Loop

90    Exit Sub

LoadOutstandingCoag_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmViewResultsWE", "LoadOutstandingCoag", intEL, strES, sql


End Sub

Private Sub LoadBiochemistry(ByVal Cn As Integer)

          Dim S As String
          Dim l As Single
          Dim h As Single
          Dim l_Range As String
          Dim Value As Single
          Dim Valu As String
          Dim BRs As New BIEResults
          Dim Br As BIEResult
          'Dim CodeForEGFR As String
          Dim MaskFlag As String
          Dim i As Integer
          Dim sql As String
10        On Error GoTo LoadBiochemistry_Error

          '20    CodeForEGFR = GetOptionSetting("BioCodeForEGFR", "5555", "")
          'MsgBox "bio"
20        Set BRs = BRs.Load("Bio", lblSampleID, "Results", gDONTCARE, gDONTCARE, "", Cn)
30        lblFasting.Visible = False
40        gBio.Visible = False
50        gBio.Rows = 2
60        gBio.AddItem ""
70        gBio.RemoveItem 1

80        gBiomnis.Visible = False
90        gBiomnis.Rows = 2
100       gBiomnis.AddItem ""
110       gBiomnis.RemoveItem 1
120       gBio.Rows = 1
130       gBio.Row = 0
140       gBiomnis.Rows = 1
150       gBiomnis.Row = 0
160       l_Range = ""
          'cmdPrint(0).Visible = False

          '      gBio.Rows = 1
          '      gBio.Row = 0
          '      gBio.Visible = True
          '
          '      gBiomnis.Rows = 1
          '      gBiomnis.Row = 0
170       DoEvents
180       DoEvents


190       For Each Br In BRs
200           l_Range = ""
210           If Br.Code = "2000" And Br.Analyser = "Calc" Then
220               sql = "DELETE FROM Bioresults WHERE SampleID = '" & Br.SampleID & "'  AND Code = '2000' AND Analyser = '2000'"
230               Cnxn(0).Execute sql
240               GoTo continue
250           End If
260           lblFasting.Visible = Br.Fasting
270           If Br.ShortName <> "H" And Br.ShortName <> "I" And Br.ShortName <> "L" Then
                  '110           If Br.Code = CodeForEGFR And Br.Valid Then
                  '120               cmdComments.Visible = True
                  '130           End If

280               MaskFlag = MaskInhibit(Br, BRs)

290               If MaskFlag <> "" Then
300                   Valu = "*****"
                      '+++ Junaid 24-08-2023
310                   If Br.Code <> "2859" Then
320                       l = Val(Br.Low)
330                       h = Val(Br.High)
340                       If l = 0 And (h = 999 Or h = 0 Or h = 9999) Then
350                           l_Range = l_Range & ""
360                       ElseIf l = 0 Then
370                           l_Range = l_Range & "< " & Format$(h)
380                       ElseIf h = 999 Or h = 9999 Then
390                           l_Range = l_Range & "> " & Format$(l)
400                       Else
410                           l_Range = l_Range & Format$(l) & "-" & Format$(h)
420                       End If
430                       S = Br.LongName & vbTab & Valu & vbTab & IIf(Br.Code = "2499", "", l_Range) & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser & vbTab & vbTab & Br.Code & vbTab & Br.Valid
                          '190               S = Br.LongName & vbTab & Valu & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser & vbTab & vbTab & Br.Code & vbTab & Br.Valid
                          '--- Junaid
440                   End If

                      '   BIOMNIS MATLAB NVRL  BEAUMONT
450                   If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
460                       gBiomnis.AddItem S
470                       With gBiomnis
480                           .Row = .Rows - 1
490                           .col = 6

500                           If Br.SignOff = False And Br.Valid = True Then
510                               Set .CellPicture = imgRedCross.Picture
520                           ElseIf Br.SignOff = True And Br.Valid = True Then
530                               Set .CellPicture = imgGreyTick.Picture
540                           End If
550                       End With
560                   Else
570                       gBio.AddItem S

580                       With gBio
590                           .Row = .Rows - 1
600                           .col = 6

610                           If Br.SignOff = False And Br.Valid = True Then
620                               Set .CellPicture = imgRedCross.Picture
630                           ElseIf Br.SignOff = True And Br.Valid = True Then
640                               Set .CellPicture = imgGreyTick.Picture
650                           End If
660                       End With


670                   End If




680                   If Br.Comment <> "" Then
690                       If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
700                           gBiomnis.Row = gBiomnis.Rows - 1
710                           gBiomnis.col = 4
720                           gBiomnis.CellBackColor = vbRed
730                           gBiomnis.CellForeColor = vbRed
740                       Else
750                           gBio.Row = gBio.Rows - 1
760                           gBio.col = 4
770                           gBio.CellBackColor = vbRed
780                           gBio.CellForeColor = vbRed
790                       End If
800                   End If
810               Else
820                   If Br.Valid Then
830                       If IsNumeric(Br.Result) Then
840                           Value = Val(Br.Result)
850                           Select Case Br.Printformat
                                  Case 0: Valu = Format(Value, "0")
860                               Case 1: Valu = Format(Value, "0.0")
870                               Case 2: Valu = Format(Value, "0.00")
880                               Case 3: Valu = Format(Value, "0.000")
890                               Case Else: Valu = Format(Value, "0.000")
900                           End Select
910                       Else
920                           Valu = Br.Result
930                       End If
940                   Else
950                       Valu = "Not Valid"
960                   End If

970                   If Br.Valid And Br.LongName = "CRP" And Val(Br.Result) < 5 And IsNumeric(Br.Result) Then
980                       Valu = "<5"
990                   End If

1000                  If Br.Code <> sysOptBioCodeForGentamicin(0) And Br.Code <> sysOptBioCodeForTobramicin(0) Then
                          '+++ Junaid 24-08-2023

1010                      l = Val(Br.Low)
1020                      h = Val(Br.High)
1030                      If l = 0 And (h = 999 Or h = 0 Or h = 9999) Then
1040                          l_Range = l_Range & ""
1050                      ElseIf l = 0 Then
1060                          l_Range = l_Range & "< " & Format$(h)
1070                      ElseIf h = 999 Or h = 9999 Then
1080                          l_Range = l_Range & "> " & Format$(l)
1090                      Else
1100                          l_Range = l_Range & Format$(l) & "-" & Format$(h)
1110                      End If

1120                      S = Br.LongName & vbTab & Valu & vbTab & IIf(Trim(Br.Code) = "2859", "", l_Range) & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser & vbTab & vbTab & Br.Code & vbTab & Br.Valid
                          '760                   S = Br.LongName & vbTab & Valu & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser & vbTab & vbTab & Br.Code & vbTab & Br.Valid
                          '--- Junaid

                          '740                       S = Br.LongName & vbTab & Valu & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser & vbTab & vbTab & Br.Code
1130                      If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Or UCase(Trim(Br.Analyser)) = "REJ" Then
1140                          gBiomnis.AddItem S

1150                          With gBiomnis
1160                              .Row = .Rows - 1
1170                              .col = 6

1180                              If Br.SignOff = False And Br.Valid = True Then
1190                                  Set .CellPicture = imgRedCross.Picture
1200                              ElseIf Br.SignOff = True And Br.Valid = True Then
1210                                  Set .CellPicture = imgGreyTick.Picture
1220                              End If
1230                          End With

1240                      Else
1250                          gBio.AddItem S


1260                          With gBio
1270                              .Row = .Rows - 1
1280                              .col = 6

1290                              If Br.SignOff = False And Br.Valid = True Then
1300                                  Set .CellPicture = imgRedCross.Picture
1310                              ElseIf Br.SignOff = True And Br.Valid = True Then
1320                                  Set .CellPicture = imgGreyTick.Picture
1330                              End If
1340                          End With
1350                      End If

                          '                    gBio.Row = gBio.Rows - 1
                          '                    gBio.col = 6
                          '
                          '                    If Br.SignOff = False And Br.Valid = True Then
                          '                        Set gBio.CellPicture = imgRedCross.Picture
                          '                    ElseIf Br.SignOff = True And Br.Valid = True Then
                          '                        Set gBio.CellPicture = imgGreyTick.Picture
                          '                    End If

1360                      If Br.Valid Then
1370                          S = QuickInterpBio(Br)
1380                          If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
1390                              Select Case Trim$(S)
                                      Case "Low":
1400                                      gBiomnis.Row = gBiomnis.Rows - 1
1410                                      gBiomnis.col = 1
1420                                      gBiomnis.CellBackColor = sysOptLowBack(0)
1430                                      gBiomnis.CellForeColor = sysOptLowFore(0)
1440                                  Case "High":
1450                                      gBiomnis.Row = gBiomnis.Rows - 1
1460                                      gBiomnis.col = 1
1470                                      gBiomnis.CellBackColor = sysOptHighBack(0)
1480                                      gBiomnis.CellForeColor = sysOptHighFore(0)
1490                                  Case Else:
1500                                      gBiomnis.Row = gBiomnis.Rows - 1
1510                                      gBiomnis.col = 1
1520                                      gBiomnis.CellBackColor = 0
1530                                      gBiomnis.CellForeColor = vbBlack
1540                              End Select
1550                          Else
1560                              Select Case Trim$(S)
                                      Case "Low":
1570                                      gBio.Row = gBio.Rows - 1
1580                                      gBio.col = 1
1590                                      gBio.CellBackColor = sysOptLowBack(0)
1600                                      gBio.CellForeColor = sysOptLowFore(0)
1610                                  Case "High":
1620                                      gBio.Row = gBio.Rows - 1
1630                                      gBio.col = 1
1640                                      gBio.CellBackColor = sysOptHighBack(0)
1650                                      gBio.CellForeColor = sysOptHighFore(0)
1660                                  Case Else:
1670                                      gBio.Row = gBio.Rows - 1
1680                                      gBio.col = 1
1690                                      gBio.CellBackColor = 0
1700                                      gBio.CellForeColor = vbBlack
1710                              End Select
1720                          End If

1730                      End If
1740                      If Br.Comment <> "" Then
1750                          If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
1760                              gBiomnis.Row = gBiomnis.Rows - 1
1770                              gBiomnis.col = 4
1780                              gBiomnis.CellBackColor = vbRed
1790                              gBiomnis.CellForeColor = vbRed
1800                          Else
1810                              gBio.Row = gBio.Rows - 1
1820                              gBio.col = 4
1830                              gBio.CellBackColor = vbRed
1840                              gBio.CellForeColor = vbRed
1850                          End If
1860                      End If
1870                  Else
1880                      If Br.Code = sysOptBioCodeForGentamicin(0) Then
1890                          S = CheckGentTobra(lblSampleID, lblName, "Gentamicin", Br.Code, Br.Result)
1900                      ElseIf Br.Code = sysOptBioCodeForTobramicin(0) Then
1910                          S = CheckGentTobra(lblSampleID, lblName, "Tobramicin", Br.Code, Br.Result)
1920                      End If
1930                      If S = "" Then
1940                          If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
1950                              gBiomnis.AddItem Br.LongName & vbTab & Br.Result & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
1960                          Else
1970                              gBio.AddItem Br.LongName & vbTab & Br.Result & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
1980                          End If

1990                          If Br.Comment <> "" Then
2000                              If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
2010                                  gBiomnis.Row = gBiomnis.Rows - 1
2020                                  gBiomnis.col = 4
2030                                  gBiomnis.CellBackColor = vbRed
2040                                  gBiomnis.CellForeColor = vbRed
2050                              Else
2060                                  gBio.Row = gBio.Rows - 1
2070                                  gBio.col = 4
2080                                  gBio.CellBackColor = vbRed
2090                                  gBio.CellForeColor = vbRed
2100                              End If

2110                          End If
2120                      Else
2130                          If InStr(S, "Peak") <> 0 Then
2140                              If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
2150                                  gBiomnis.AddItem Br.LongName & " Trough" & vbTab & Br.Result & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2160                              Else
2170                                  gBio.AddItem Br.LongName & " Trough" & vbTab & Br.Result & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2180                              End If
2190                          Else
2200                              If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
2210                                  gBiomnis.AddItem Br.LongName & " Peak" & vbTab & Br.Result & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2220                              Else
2230                                  gBio.AddItem Br.LongName & " Peak" & vbTab & Br.Result & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2240                              End If
2250                          End If
2260                          If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
2270                              gBiomnis.AddItem S & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2280                          Else
2290                              gBio.AddItem S & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser

2300                          End If
2310                          If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
2320                              gBiomnis.AddItem S & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2330                          Else
2340                              gBio.AddItem S & vbTab & Br.Low & "-" & Br.High & vbTab & Br.Units & vbTab & Br.Comment & vbTab & Br.Analyser
2350                          End If
2360                          If Br.Comment <> "" Then
2370                              If UCase(Trim(Br.Analyser)) = "BIOMNIS" Or UCase(Trim(Br.Analyser)) = "MATER" Or UCase(Trim(Br.Analyser)) = "NVRL" Or UCase(Trim(Br.Analyser)) = "BEAUMONT" Then
2380                                  gBiomnis.Row = gBiomnis.Rows - 1
2390                                  gBiomnis.col = 3
2400                                  gBiomnis.CellBackColor = vbRed
2410                                  gBiomnis.CellForeColor = vbRed
2420                              Else
2430                                  gBio.Row = gBio.Rows - 1
2440                                  gBio.col = 3
2450                                  gBio.CellBackColor = vbRed
2460                                  gBio.CellForeColor = vbRed
2470                              End If
2480                          End If
2490                      End If
2500                  End If
2510              End If
2520          End If
2530          DoEvents
2540          DoEvents
continue:
2550      Next
          'farhan--------------------------->
2560      For i = 0 To gBio.Rows - 1
2570          If Trim(gBio.TextMatrix(i, 2)) = "0-9999" Then
2580              gBio.TextMatrix(i, 2) = ""
2590          End If
2600      Next
2610      For i = 0 To gBiomnis.Rows - 1
2620          If Trim(gBiomnis.TextMatrix(i, 2)) = "0-9999" Then
2630              gBiomnis.TextMatrix(i, 2) = ""
2640          End If
2650      Next
          'farhan<-------------------------
          '+++Junaid 05-10-2023
2660      If gBio.Rows > 1 Then
              '2280  If gBio.Rows > 0 Then
              '---Junaid
              '2290      gBio.RemoveItem 1
2670          SSTabBio.TabCaption(0) = "<< Biochemistry >>"
2680      Else
2690          SSTabBio.TabCaption(0) = "Biochemistry"
2700      End If
          '+++Junaid 05-10-2023
2710      If gBiomnis.Rows > 1 Then
              '2340  If gBiomnis.Rows > 0 Then
              '---Junaid
              '2350      gBiomnis.RemoveItem 1
2720          SSTabBio.TabCaption(1) = "<< External >>"
2730      Else
2740          SSTabBio.TabCaption(1) = "External"
2750      End If

2760      LoadOutstandingBio

2770      gBio.Visible = True
2780      gBiomnis.Visible = True

2790      If IsPrintable(lblSampleID, lblRunDate, "Biochemistry") Then
2800          cmdPrint(0).Visible = True
2810      End If

2820      If SSTabBio.Tab = 0 Then
2830          Call SignOffEnable(gBio, 6, cmdSignOffBio)
2840      ElseIf SSTabBio.Tab = 1 Then
2850          Call SignOffEnable(gBiomnis, 6, cmdSignOffBio)
2860      End If
2870      Exit Sub

LoadBiochemistry_Error:

          Dim strES As String
          Dim intEL As Integer

2880      intEL = Erl
2890      strES = Err.Description
2900      LogError "frmViewResultsWE", "LoadBiochemistry", intEL, strES

End Sub
Private Sub LoadOutstandingBio()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadOutstandingBio_Error

20    If lblSampleID = "" Then Exit Sub

30    sql = "select distinct BT.LongName from BioRequests as BR, BioTestDefinitions as BT where " & _
            "BR.SampleID = '" & lblSampleID & "' " & _
            "and BR.Code = BT.Code and bt.sampletype = BR.SampleType"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    Do While Not tb.EOF
70        gBio.AddItem tb!LongName & vbTab & "In Progress"
80        tb.MoveNext
90    Loop

100   Exit Sub

LoadOutstandingBio_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmViewResultsWE", "LoadOutstandingBio", intEL, strES, sql


End Sub
Private Sub LoadImmunology(ByVal Cn As Integer)

      Dim S As String
      Dim Value As Single
      Dim Valu As String
      Dim BRs As New BIEResults
      Dim Br As BIEResult
      Dim TitleAdded As Boolean
      Dim BioNotPresent As Boolean

10    On Error GoTo LoadImmunology_Error

20    Set BRs = BRs.Load("Imm", lblSampleID, "Results", gDONTCARE, gDONTCARE, "", Cn)

30    TitleAdded = False

40    gBio.Visible = False

50    BioNotPresent = False
60    If gBio.Rows = 2 And gBio.TextMatrix(1, 0) = "" Then
70        BioNotPresent = True
80    End If

90    For Each Br In BRs
100       If Not TitleAdded Then
110           gBio.AddItem ""
120           gBio.AddItem "Immunology"
130           gBio.Row = gBio.Rows - 1
140           gBio.col = 0
150           gBio.CellBackColor = vbYellow
160           gBio.CellForeColor = vbBlue

170           TitleAdded = True
180       End If
190       If Br.Valid Then
200           If IsNumeric(Br.Result) Then
210               Value = Val(Br.Result)
220               Select Case Br.Printformat
                  Case 0: Valu = Format(Value, "0")
230               Case 1: Valu = Format(Value, "0.0")
240               Case 2: Valu = Format(Value, "0.00")
250               Case 3: Valu = Format(Value, "0.000")
260               Case Else: Valu = Format(Value, "0.000")
270               End Select
280           Else
290               Valu = Br.Result
300           End If
310       Else
320           Valu = "NV"
330       End If
340       S = Br.LongName & vbTab & Valu
350       gBio.AddItem S
360       If Br.Valid Then
370           S = QuickInterpBio(Br)
380           Select Case Trim$(S)
              Case "Low":
390               gBio.Row = gBio.Rows - 1
400               gBio.col = 1
410               gBio.CellBackColor = sysOptLowBack(0)
420               gBio.CellForeColor = sysOptLowFore(0)
430           Case "High":
440               gBio.Row = gBio.Rows - 1
450               gBio.col = 1
460               gBio.CellBackColor = sysOptHighBack(0)
470               gBio.CellForeColor = sysOptHighFore(0)
480           Case Else:
490               gBio.Row = gBio.Rows - 1
500               gBio.col = 1
510               gBio.CellBackColor = 0
520               gBio.CellForeColor = vbBlack
530           End Select
540       End If
550   Next

560   If BioNotPresent And gBio.Rows > 3 Then
570       gBio.RemoveItem 1
580       gBio.RemoveItem 1
590   End If
600   gBio.Visible = True

610   Exit Sub

LoadImmunology_Error:

      Dim strES As String
      Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "frmViewResultsWE", "LoadImmunology", intEL, strES


End Sub

Private Sub LoadBloodGas(ByVal Cn As Integer)

      Dim Bx As BGAResult
      Dim Bxs As New BGAResults

10    On Error GoTo LoadBloodGas_Error

20    ClearBloodGas

30    Set Bx = Bxs.LoadResults(Cn, lblSampleID)

40    If Not Bx Is Nothing Then

50        fraBGA.Font.Bold = True
60        fraBGA.ForeColor = vbRed

70        With Bx
80            txtPh = .pH
90            txtPo2 = .PO2
100           txtPco2 = .PCO2
110           txtHco3 = .HCO3
120           txtBE = .BE
130           txtO2Sat = .O2SAT
140           txtTotCo2 = .TotCO2
150           lblTimeTaken = Format(.RunDateTime, "hh:mm")
160           lblRunDate = Format$(.Rundate, "dd/mm/yyyy")
170       End With

180   End If

190   Exit Sub

LoadBloodGas_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmViewResultsWE", "LoadBloodGas", intEL, strES

End Sub


Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub


Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    Activated = False

      '
      'CheckDisciplineActive "Bio"
      'CheckDisciplineActive "Imm"
      'CheckDisciplineActive "End"

30    gBio.Font.Bold = True

40    PBar.Max = LogOffDelaySecs
50    PBar = 0


60    lblChartTitle = "Chart"

70    If IsIDE Then
80        cmdOrder.Visible = True
90    Else
100       cmdOrder.Visible = sysOptOrderComms(0)
110   End If

120   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmViewResultsWE", "Form_Load", intEL, strES

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub


Private Sub gBio_Click()

10    On Error GoTo gBio_Click_Error

20    If gBio.ColSel = 4 Then
30        If gBio.CellBackColor = vbRed Then
40            iMsg gBio.TextMatrix(gBio.Row, 4), vbOKOnly, gBio.TextMatrix(gBio.Row, 0) & " Comment", vbWhite
50        End If
60    ElseIf gBio.ColSel = 6 Then
70        With gBio
80            .Row = .RowSel
90            .col = 6
100           If .CellPicture <> imgGreyTick.Picture Then
110               If .CellPicture = imgRedCross.Picture Then
120                   Set .CellPicture = imgGreenTick.Picture
130               Else
140                   Set .CellPicture = imgRedCross.Picture
150               End If
160           End If
170       End With

190   End If

200   Exit Sub

gBio_Click_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmViewResultsWE", "gBio_Click", intEL, strES

End Sub

Private Sub gBio_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    On Error GoTo gBio_MouseMove_Error

20    If (gBio.MouseCol = 0 Or gBio.MouseCol = 1) And gBio.MouseRow > 0 Then
30        gBio.ToolTipText = gBio.TextMatrix(gBio.MouseRow, gBio.MouseCol)
40    Else
50        gBio.ToolTipText = ""
60    End If

70    Exit Sub

gBio_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmViewResultsWE", "gBio_MouseMove", intEL, strES

End Sub

Private Sub gBiomnis_Click()

10    On Error GoTo gBiomnis_Click_Error

20    If gBiomnis.ColSel = 4 Then
30        iMsg gBiomnis.TextMatrix(gBiomnis.Row, 4), vbOKOnly, gBiomnis.TextMatrix(gBiomnis.Row, 0) & " Comment", vbWhite
40    ElseIf gBiomnis.ColSel = 6 Then
50        With gBiomnis
60            .Row = .RowSel
70            .col = 6
80            If .CellPicture <> imgGreyTick.Picture Then
90                If .CellPicture = imgRedCross.Picture Then
100                   Set .CellPicture = imgGreenTick.Picture
110               Else
120                   Set .CellPicture = imgRedCross.Picture
130               End If
140           End If
150       End With

160   End If

170   Exit Sub

gBiomnis_Click_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmViewResultsWE", "gBiomnis_Click", intEL, strES

End Sub

Private Sub gBiomnis_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10    On Error GoTo gBiomnis_MouseMove_Error

20    If (gBiomnis.MouseCol = 0 Or gBiomnis.MouseCol = 1) And gBiomnis.MouseRow > 0 Then
30        gBiomnis.ToolTipText = gBiomnis.TextMatrix(gBiomnis.MouseRow, gBiomnis.MouseCol)
40    Else
50        gBiomnis.ToolTipText = ""
60    End If

70    Exit Sub

gBiomnis_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmViewResultsWE", "gBiomnis_MouseMove", intEL, strES

End Sub

Private Sub gCoag_Click()

10    On Error GoTo gCoag_Click_Error

20    If gCoag.MouseCol = 3 Then
30        iMsg gCoag.TextMatrix(gCoag.Row, 3), vbOKOnly, gCoag.TextMatrix(gCoag.Row, 0) & " Comment", vbWhite
40    ElseIf gCoag.ColSel = 5 Then
50        With gCoag
60            .Row = .RowSel
70            .col = 5
80            If .CellPicture <> imgGreyTick.Picture Then
90                If .CellPicture = imgRedCross.Picture Then
100                   Set .CellPicture = imgGreenTick.Picture
110               Else
120                   Set .CellPicture = imgRedCross.Picture
130               End If
140           End If
150       End With

160   End If

170   Exit Sub

gCoag_Click_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmViewResultsWE", "gCoag_Click", intEL, strES

End Sub

Private Sub gCoag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    On Error GoTo gCoag_MouseMove_Error

20    gCoag.ToolTipText = ""
30    If gCoag.MouseRow > 0 And gCoag.MouseCol = 3 Then
40        If Trim$(gCoag.TextMatrix(gCoag.MouseRow, 3)) <> "" Then
50            gCoag.ToolTipText = gCoag.TextMatrix(gCoag.MouseRow, 3)
60        End If
70    End If

80    Exit Sub

gCoag_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmViewResultsWE", "gCoag_MouseMove", intEL, strES

End Sub


Private Sub grdDemog_Compare(ByVal Row1 As Long, ByVal Row2 As Long, cmp As Integer)

      Dim d1 As Date
      Dim d2 As Date
      Dim Column As Integer

10    With grdDemog
20        Column = .col
30        cmp = 0
40        If IsDate(.TextMatrix(Row1, Column)) Then
50            d1 = Format(.TextMatrix(Row1, Column), "dd/mmm/yyyy")
60            If IsDate(.TextMatrix(Row2, Column)) Then
70                d2 = Format(.TextMatrix(Row2, Column), "dd/mmm/yyyy")
80                cmp = Sgn(DateDiff("d", d1, d2))
90            End If
100       End If
110   End With

End Sub


Private Sub imgEarliest_Click()
    On Error Resume Next
10    grdDemog.Row = grdDemog.Rows - 1

20    FillDemographics
30    LoadAllResults

40    If SSTabBio.Tab = 1 And gBiomnis.TextMatrix(1, 0) <> "" Then
50        AddActivity lblSampleID, "Ward Enquiry External Results", "VIEWED", "", lblChart, "", ""
60    End If

70    imgEarliest.Visible = False
80    imgPrevious.Visible = False
90    imgNext.Visible = True
100   imgLatest.Visible = True

110   PBar = 0

End Sub

Private Sub imgHaemGraphs_Click()

10    If sysOptHaemAn1(0) = "MaxM" Then
20        frmMaxMGraphs.SampleID = lblSampleID
30        frmMaxMGraphs.Show 1
40    Else
50        frmHaemGraphs.SampleID = lblSampleID
60        frmHaemGraphs.Show 1
70    End If

End Sub

Private Sub imgLatest_Click()
    On Error Resume Next
10    grdDemog.Row = 1

20    FillDemographics
30    LoadAllResults

40    If SSTabBio.Tab = 1 And gBiomnis.TextMatrix(1, 0) <> "" Then
50        AddActivity lblSampleID, "Ward Enquiry External Results", "VIEWED", "", lblChart, "", ""
60    End If

70    imgNext.Visible = False
80    imgLatest.Visible = False
90    imgPrevious.Visible = True
100   imgEarliest.Visible = True

110   PBar = 0

End Sub

Private Sub imgNext_Click()
    On Error Resume Next
10    grdDemog.Row = grdDemog.Row - 1

20    FillDemographics
30    LoadAllResults

40    If grdDemog.Row > 1 Then
50        imgNext.Visible = True
60        imgLatest.Visible = True
70    Else
80        imgNext.Visible = False
90        imgLatest.Visible = False
100   End If

110   If SSTabBio.Tab = 1 And gBiomnis.TextMatrix(1, 0) <> "" Then
120       AddActivity lblSampleID, "Ward Enquiry External Results", "VIEWED", "", lblChart, "", ""
130   End If


140   imgPrevious.Visible = True
150   imgEarliest.Visible = True

160   PBar = 0

End Sub

Private Sub imgPrevious_Click()
    On Error Resume Next
10    grdDemog.Row = grdDemog.Row + 1

20    FillDemographics
30    LoadAllResults

40    If grdDemog.Row < grdDemog.Rows - 1 Then
50        imgEarliest.Visible = True
60        imgPrevious.Visible = True
70    Else
80        imgEarliest.Visible = False
90        imgPrevious.Visible = False
100   End If


110   If SSTabBio.Tab = 1 And gBiomnis.TextMatrix(1, 0) <> "" Then
120       AddActivity lblSampleID, "Ward Enquiry External Results", "VIEWED", "", lblChart, "", ""
130   End If

140   imgNext.Visible = True
150   imgLatest.Visible = True

160   PBar = 0

End Sub



'---------------------------------------------------------------------------------------
' Procedure : SSTabBio_Click
' Author    : Masood
' Date      : 26/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SSTabBio_Click(PreviousTab As Integer)

10    On Error GoTo SSTabBio_Click_Error


20    If SSTabBio.Tab = 0 Then
30        Call SignOffEnable(gBio, 6, cmdSignOffBio)
40    ElseIf SSTabBio.Tab = 1 Then
50        Call SignOffEnable(gBiomnis, 6, cmdSignOffBio)
60        If gBiomnis.TextMatrix(1, 0) <> "" Then
70            AddActivity lblSampleID, "Ward Enquiry External Results", "VIEWED", "", lblChart, "", ""
80        End If
90    End If


100   Exit Sub


SSTabBio_Click_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmViewResultsWE", "SSTabBio_Click", intEL, strES
End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1

20    If PBar = PBar.Max Then
30        LogOffNow = True
40        Unload Me
50    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : bSignOffEnable
' Author    : Masood
' Date      : 26/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SignOffEnable(ByVal f As MSFlexGrid, Colum As Integer, ByVal btnSignOF As CommandButton)

10    On Error GoTo bSignOffEnable_Error

      'cmdSignOffCoag.Enabled = False
20    btnSignOF.Enabled = False
      Dim i As Integer
30    With f
40        For i = 0 To .Rows - 1
50            .Row = i
60            .col = Colum
70            If .CellPicture = imgRedCross.Picture Then
                  'cmdSignOffCoag.Enabled = True
80                btnSignOF.Enabled = True
90                Exit Sub
100           End If
110       Next i
120   End With


130   Exit Sub


bSignOffEnable_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmViewResultsWE", "bSignOffEnable", intEL, strES
End Sub


Private Function SignOffRequired(ByVal f As MSFlexGrid, Colum As Integer) As Boolean
10    On Error GoTo SignOffRequired_Error

      Dim i As Integer
20    With f
30        For i = 0 To .Rows - 1
40            .Row = i
50            .col = Colum
60            If .CellPicture = imgRedCross.Picture Then
                  
70                SignOffRequired = True
80                Exit Function
90            End If
100       Next i
110   End With

120   Exit Function
SignOffRequired_Error:
         
130   LogError "frmViewResultsWE", "SignOffRequired", Erl, Err.Description


End Function

Private Function SignOffRequiredHaem(chartNumber As String) As Boolean
10    On Error GoTo SignOffRequiredHaem_Error

      Dim tb As Recordset
      Dim sql As String

20    sql = "Select Count(sampleId) as CNT From HaemResults Where SampleID in (select sampleid from demographics where Chart ='" & chartNumber & "') and SignOff is NULL"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb!Cnt > 0 Then
60        SignOffRequiredHaem = True
70    End If


80    Exit Function
SignOffRequiredHaem_Error:
         
90    LogError "frmViewResultsWE", "SignOffRequiredHaem", Erl, Err.Description, sql
End Function


