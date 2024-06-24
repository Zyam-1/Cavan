VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditAll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frmEditAll.frx":0000
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   308
      Top             =   10170
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   240
      TabIndex        =   307
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame fraWait 
      BackColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   15000
      TabIndex        =   186
      Top             =   5000
      Width           =   3975
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         TabIndex        =   294
         Top             =   2040
         Width           =   1920
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   1020
         Picture         =   "frmEditAll.frx":00CD
         Top             =   120
         Width           =   1920
      End
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
      Height          =   285
      Left            =   15240
      MaxLength       =   8
      TabIndex        =   282
      Top             =   780
      Width           =   1125
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
      Height          =   2145
      Left            =   14340
      TabIndex        =   281
      Top             =   7260
      Width           =   5280
   End
   Begin VB.Timer tmrUpDown 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   -30
      Top             =   2400
   End
   Begin VB.CommandButton cmdTag 
      Caption         =   "Tag"
      Height          =   315
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   268
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdViewReports 
      BackColor       =   &H00FFFF00&
      Caption         =   "View Reports"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   263
      Top             =   2010
      Width           =   1275
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":1594
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Log as Phoned"
      Top             =   6780
      Width           =   1245
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":19D6
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   11310
      Top             =   -120
   End
   Begin VB.CommandButton bPrintHold 
      Caption         =   "Print && Hold"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":1CE0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4395
      Width           =   1275
   End
   Begin VB.Frame fr 
      Caption         =   "Frame1"
      Height          =   1815
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   210
      Width           =   14055
      Begin VB.CommandButton bsearchDob 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Left            =   12780
         TabIndex        =   279
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
         Left            =   8700
         TabIndex        =   3
         Top             =   870
         Width           =   3495
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
         Left            =   6180
         TabIndex        =   2
         Top             =   870
         Width           =   2505
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
         Height          =   450
         Left            =   4740
         MaxLength       =   8
         TabIndex        =   1
         Top             =   870
         Width           =   1425
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   12780
         MaxLength       =   10
         TabIndex        =   4
         Top             =   570
         Width           =   1125
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   12780
         MaxLength       =   4
         TabIndex        =   5
         Top             =   930
         Width           =   1125
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   12780
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1290
         Width           =   1125
      End
      Begin VB.Frame fr 
         Caption         =   "EnableCopyFrom"
         Height          =   1905
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   4725
         Begin VB.ComboBox cmbOtherSamples 
            Height          =   315
            Left            =   180
            TabIndex        =   297
            Top             =   1200
            Width           =   4485
         End
         Begin VB.CommandButton cmdPatientNotePad 
            Height          =   500
            Left            =   4230
            Picture         =   "frmEditAll.frx":234A
            Style           =   1  'Graphical
            TabIndex        =   280
            Top             =   270
            Width           =   465
         End
         Begin VB.CommandButton cmdscan 
            Caption         =   "&Scan"
            Height          =   870
            Left            =   3300
            Picture         =   "frmEditAll.frx":2C14
            Style           =   1  'Graphical
            TabIndex        =   276
            Top             =   270
            Width           =   885
         End
         Begin VB.CommandButton cmdViewScan 
            Caption         =   "&View Scan"
            Height          =   870
            Left            =   2400
            Picture         =   "frmEditAll.frx":3C96
            Style           =   1  'Graphical
            TabIndex        =   275
            Top             =   270
            Width           =   885
         End
         Begin VB.CommandButton cmdUpDown 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1920
            TabIndex        =   271
            Top             =   780
            Width           =   285
         End
         Begin VB.CommandButton cmdUpDown 
            Caption         =   "+"
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   270
            Top             =   510
            Width           =   285
         End
         Begin VB.ComboBox cMRU 
            Height          =   315
            Left            =   570
            TabIndex        =   19
            Text            =   "cMRU"
            Top             =   1500
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
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Other Samples"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   298
            Top             =   990
            Width           =   1035
         End
         Begin VB.Label lblResultOrRequest 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Results"
            Height          =   285
            Left            =   600
            TabIndex        =   24
            Top             =   210
            Width           =   885
         End
         Begin VB.Image imgLast 
            Height          =   300
            Left            =   2070
            Picture         =   "frmEditAll.frx":9484
            Stretch         =   -1  'True
            ToolTipText     =   "Find Last Record"
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "MRU"
            Height          =   195
            Index           =   50
            Left            =   150
            TabIndex        =   20
            Top             =   1560
            Width           =   375
         End
         Begin VB.Image iRelevant 
            Height          =   480
            Index           =   1
            Left            =   1500
            Picture         =   "frmEditAll.frx":98C6
            Top             =   150
            Width           =   480
         End
         Begin VB.Image iRelevant 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   150
            Picture         =   "frmEditAll.frx":9BD0
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sample ID"
            Height          =   195
            Index           =   58
            Left            =   720
            TabIndex        =   16
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Searc&h"
         Height          =   345
         Left            =   11520
         TabIndex        =   7
         Top             =   480
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   285
         Left            =   13200
         TabIndex        =   8
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Chart #"
         Height          =   255
         Left            =   4770
         TabIndex        =   305
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblSurNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "SurName                                         ForeName"
         Height          =   195
         Left            =   6210
         TabIndex        =   29
         Top             =   660
         Width           =   3240
      End
      Begin VB.Label lblSampleDate 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10230
         TabIndex        =   27
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SampleDate"
         Height          =   195
         Index           =   51
         Left            =   9330
         TabIndex        =   26
         Top             =   150
         Width           =   870
      End
      Begin VB.Label lblUrgent 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   6540
         TabIndex        =   25
         Top             =   120
         Width           =   2100
      End
      Begin VB.Label lAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4740
         TabIndex        =   22
         Top             =   1350
         Width           =   7455
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
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "D.o.B             Age              Sex"
         Height          =   945
         Index           =   56
         Left            =   12330
         TabIndex        =   17
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":9EDA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2805
      Width           =   1275
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "&History"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":A1E4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":A626
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "bprint"
      Top             =   5175
      Width           =   1275
   End
   Begin VB.CommandButton cmdFAX 
      Caption         =   "&Fax Results"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":AC90
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5985
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   12870
      Picture         =   "frmEditAll.frx":B0D2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9210
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   150
      TabIndex        =   31
      Top             =   2025
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmEditAll.frx":B73C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDemogValid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRequestID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSaveDemographics"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCopyFromPrevious"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOrderExt(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "bOrderTests"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fr(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSaveHoldDemographics"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fr(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fr(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdValidationList"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdValidateDemographics"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdUnLock"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fmeNote"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "flxQuestions"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Haematology"
      TabPicture(1)   =   "frmEditAll.frx":B758
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(61)"
      Tab(1).Control(1)=   "lblAnalyser"
      Tab(1).Control(2)=   "Label1(38)"
      Tab(1).Control(3)=   "Label1(39)"
      Tab(1).Control(4)=   "Label1(37)"
      Tab(1).Control(5)=   "Label1(36)"
      Tab(1).Control(6)=   "lHaemErrors"
      Tab(1).Control(7)=   "lblHaemPrinted"
      Tab(1).Control(8)=   "gOutstandingHaem"
      Tab(1).Control(9)=   "cmdSaveHaem"
      Tab(1).Control(10)=   "pbCD"
      Tab(1).Control(11)=   "cmbFilmComment"
      Tab(1).Control(12)=   "cmbHaemComment"
      Tab(1).Control(13)=   "txtFilmComment"
      Tab(1).Control(14)=   "cmdDiff"
      Tab(1).Control(15)=   "cFilm"
      Tab(1).Control(16)=   "bHaemGraphs"
      Tab(1).Control(17)=   "txtHaemComment"
      Tab(1).Control(18)=   "Panel3D1(2)"
      Tab(1).Control(19)=   "Panel3D1(3)"
      Tab(1).Control(20)=   "Panel3D1(0)"
      Tab(1).Control(21)=   "Panel3D1(1)"
      Tab(1).Control(22)=   "cmdValidateHaem"
      Tab(1).Control(23)=   "bViewHaemRepeat"
      Tab(1).Control(24)=   "Panel3D1(4)"
      Tab(1).Control(25)=   "cmdOrderHeam"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Biochemistry"
      TabPicture(2)   =   "frmEditAll.frx":B774
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSplitView"
      Tab(2).Control(1)=   "lblSplit(5)"
      Tab(2).Control(2)=   "lblSplit(4)"
      Tab(2).Control(3)=   "lblSplit(3)"
      Tab(2).Control(4)=   "lblSplit(2)"
      Tab(2).Control(5)=   "lblSplit(1)"
      Tab(2).Control(6)=   "lblSplit(0)"
      Tab(2).Control(7)=   "lblAss"
      Tab(2).Control(8)=   "lRandom"
      Tab(2).Control(9)=   "lblSplit(6)"
      Tab(2).Control(10)=   "lblDelta(2)"
      Tab(2).Control(11)=   "lblGreaterThan"
      Tab(2).Control(12)=   "gBio"
      Tab(2).Control(13)=   "grdOutstanding"
      Tab(2).Control(14)=   "cmdSaveBio"
      Tab(2).Control(15)=   "lstAdd"
      Tab(2).Control(16)=   "fraSelectPrint(1)"
      Tab(2).Control(17)=   "cmbNewResult"
      Tab(2).Control(18)=   "bViewBioRepeat"
      Tab(2).Control(19)=   "fr(5)"
      Tab(2).Control(20)=   "cmdValidateBio"
      Tab(2).Control(21)=   "cmdAddBio"
      Tab(2).Control(22)=   "cmdOrderBio"
      Tab(2).Control(23)=   "cmbAdd"
      Tab(2).Control(24)=   "cmbUnits"
      Tab(2).Control(25)=   "cmbSampleType"
      Tab(2).Control(26)=   "fr(7)"
      Tab(2).Control(27)=   "txtAutoComment(2)"
      Tab(2).ControlCount=   28
      TabCaption(3)   =   "Coagulation"
      TabPicture(3)   =   "frmEditAll.frx":B790
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblCoagAnalyser"
      Tab(3).Control(1)=   "Label1(62)"
      Tab(3).Control(2)=   "lblPrevCoag"
      Tab(3).Control(3)=   "Label1(46)"
      Tab(3).Control(4)=   "Label3"
      Tab(3).Control(5)=   "lblDelta(0)"
      Tab(3).Control(6)=   "grdPrev"
      Tab(3).Control(7)=   "gOutstandingCoag"
      Tab(3).Control(8)=   "gCoag"
      Tab(3).Control(9)=   "fraSelectPrint(0)"
      Tab(3).Control(10)=   "cmbBioComment(2)"
      Tab(3).Control(11)=   "bPrintAll"
      Tab(3).Control(12)=   "fr(8)"
      Tab(3).Control(13)=   "txtCoagComment"
      Tab(3).Control(14)=   "cmdValidateCoag"
      Tab(3).Control(15)=   "bViewCoagRepeat"
      Tab(3).Control(16)=   "bAddCoag"
      Tab(3).Control(17)=   "cmdOrderCoag"
      Tab(3).Control(18)=   "tResult"
      Tab(3).Control(19)=   "cParameter"
      Tab(3).Control(20)=   "cmdSaveCoag"
      Tab(3).Control(21)=   "txtAutoComment(3)"
      Tab(3).ControlCount=   22
      TabCaption(4)   =   "Immunology"
      TabPicture(4)   =   "frmEditAll.frx":B7AC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Blood Gas"
      TabPicture(5)   =   "frmEditAll.frx":B7C8
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "External"
      TabPicture(6)   =   "frmEditAll.frx":B7E4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "g"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "cmdOrderExt(0)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "etc(8)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "etc(7)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "etc(6)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "etc(5)"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "etc(1)"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "etc(2)"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "etc(3)"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "etc(4)"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "etc(0)"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "cmdDeleteExt"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "cmdSaveExt"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).Control(13)=   "cmdMedibridge"
      Tab(6).Control(13).Enabled=   0   'False
      Tab(6).Control(14)=   "btnPrintDoc"
      Tab(6).Control(14).Enabled=   0   'False
      Tab(6).ControlCount=   15
      Begin VB.CommandButton btnPrintDoc 
         Caption         =   "Print Document"
         Height          =   825
         Left            =   -64140
         TabIndex        =   304
         Top             =   4920
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid flxQuestions 
         Height          =   1545
         Left            =   5910
         TabIndex        =   299
         Top             =   3390
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   2725
         _Version        =   393216
         Cols            =   4
      End
      Begin VB.Frame fmeNote 
         Caption         =   "Note"
         Height          =   1245
         Left            =   6000
         TabIndex        =   295
         Top             =   6600
         Width           =   4995
         Begin VB.TextBox txtNote 
            Height          =   975
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   296
            Top             =   210
            Width           =   4875
         End
      End
      Begin VB.CommandButton cmdOrderHeam 
         Caption         =   "Order Tests"
         Height          =   735
         Left            =   -66585
         Picture         =   "frmEditAll.frx":B800
         Style           =   1  'Graphical
         TabIndex        =   292
         Tag             =   "bOrder"
         Top             =   6615
         Width           =   1155
      End
      Begin VB.Frame Frame 
         Caption         =   "Reject Sample"
         Height          =   1275
         Left            =   405
         TabIndex        =   288
         Top             =   6585
         Width           =   5460
         Begin VB.CheckBox ChkExtReject 
            Caption         =   "Reject External Sample"
            Height          =   240
            Left            =   2790
            TabIndex        =   293
            Top             =   765
            Width           =   2490
         End
         Begin VB.CheckBox chkHaemReject 
            Caption         =   "Reject Haematology Sample"
            Height          =   240
            Left            =   2790
            TabIndex        =   291
            Top             =   315
            Width           =   2490
         End
         Begin VB.CheckBox chkCoagReject 
            Caption         =   "Reject Coagulation Sample"
            Height          =   240
            Left            =   225
            TabIndex        =   290
            Top             =   765
            Width           =   2490
         End
         Begin VB.CheckBox chkBioReject 
            Caption         =   "Reject Biochemistry Sample"
            Height          =   240
            Left            =   225
            TabIndex        =   289
            Top             =   315
            Width           =   2490
         End
      End
      Begin VB.CommandButton cmdUnLock 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Unlock Demographics"
         Height          =   1215
         Left            =   11100
         Picture         =   "frmEditAll.frx":BB0A
         Style           =   1  'Graphical
         TabIndex        =   269
         Top             =   5160
         Width           =   1245
      End
      Begin VB.CommandButton cmdValidateDemographics 
         Caption         =   "Un-Validate"
         Height          =   285
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   267
         Top             =   5910
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdValidationList 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Validation List"
         Height          =   1245
         Left            =   11100
         Picture         =   "frmEditAll.frx":C9D4
         Style           =   1  'Graphical
         TabIndex        =   265
         Top             =   6570
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdMedibridge 
         Caption         =   "View Medibridge"
         Height          =   1635
         Left            =   -64140
         Picture         =   "frmEditAll.frx":CD78
         Style           =   1  'Graphical
         TabIndex        =   264
         Top             =   930
         Width           =   1335
      End
      Begin VB.TextBox txtAutoComment 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   1725
         Index           =   3
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   262
         ToolTipText     =   "Auto-generated comments."
         Top             =   6030
         Width           =   5565
      End
      Begin VB.TextBox txtAutoComment 
         BackColor       =   &H00E0E0E0&
         Height          =   675
         Index           =   2
         Left            =   -70200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   260
         ToolTipText     =   "Auto-generated comments."
         Top             =   6210
         Width           =   3645
      End
      Begin VB.CommandButton cmdSaveExt 
         Caption         =   "&Save"
         Height          =   735
         Left            =   -64140
         Picture         =   "frmEditAll.frx":1181A
         Style           =   1  'Graphical
         TabIndex        =   256
         Top             =   5940
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveCoag 
         Caption         =   "&Save"
         Height          =   945
         Left            =   -66180
         Picture         =   "frmEditAll.frx":1319C
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   6810
         Width           =   1065
      End
      Begin VB.Frame fr 
         Caption         =   "Specimen Condition"
         Height          =   915
         Index           =   7
         Left            =   -67170
         TabIndex        =   221
         Top             =   5280
         Width           =   2505
         Begin VB.CheckBox chkOld 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Left            =   150
            TabIndex        =   222
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Haemolysed"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   228
            Top             =   510
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Icteric"
            Height          =   195
            Index           =   20
            Left            =   1710
            TabIndex        =   227
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lipaemic"
            Height          =   195
            Index           =   21
            Left            =   1710
            TabIndex        =   226
            Top             =   510
            Width           =   630
         End
         Begin VB.Label lblLipaemic 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1380
            TabIndex        =   225
            Top             =   480
            Width           =   285
         End
         Begin VB.Label lblIcteric 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1380
            TabIndex        =   224
            Top             =   210
            Width           =   285
         End
         Begin VB.Label lblHaemolysed 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1080
            TabIndex        =   223
            Top             =   480
            Width           =   285
         End
      End
      Begin VB.ComboBox cmbSampleType 
         Height          =   315
         Left            =   -69960
         TabIndex        =   220
         Text            =   "cmbSampleType"
         Top             =   5730
         Width           =   1605
      End
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   -71340
         TabIndex        =   219
         Text            =   "cmbUnits"
         Top             =   5730
         Width           =   1305
      End
      Begin VB.ComboBox cmbAdd 
         Height          =   315
         Left            =   -74850
         TabIndex        =   218
         Text            =   "cmbAdd"
         Top             =   5730
         Width           =   1635
      End
      Begin VB.Frame fr 
         Height          =   1155
         Index           =   2
         Left            =   10950
         TabIndex        =   215
         Top             =   990
         Width           =   1365
         Begin VB.CheckBox chkFasting 
            Alignment       =   1  'Right Justify
            Caption         =   "Fasting"
            Height          =   225
            Left            =   480
            TabIndex        =   303
            Top             =   900
            Width           =   825
         End
         Begin VB.CheckBox chkUrgent 
            Alignment       =   1  'Right Justify
            Caption         =   "Urgent"
            Height          =   225
            Left            =   510
            TabIndex        =   302
            Top             =   660
            Width           =   795
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   217
            Top             =   180
            Width           =   885
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   216
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Date"
         Height          =   2055
         Index           =   1
         Left            =   5910
         TabIndex        =   206
         Top             =   990
         Width           =   4905
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   1800
            TabIndex        =   207
            Top             =   1260
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   227999745
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   210
            TabIndex        =   208
            Top             =   510
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   227999745
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   1590
            TabIndex        =   209
            ToolTipText     =   "Time of Sample"
            Top             =   510
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
         Begin MSComCtl2.DTPicker dtRecDate 
            Height          =   315
            Left            =   2820
            TabIndex        =   210
            Top             =   510
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   227999745
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tRecTime 
            Height          =   315
            Left            =   4170
            TabIndex        =   211
            ToolTipText     =   "Time of Sample"
            Top             =   510
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sample                                              Received"
            Height          =   195
            Index           =   55
            Left            =   210
            TabIndex        =   214
            Top             =   300
            Width           =   3285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Run"
            Height          =   195
            Index           =   45
            Left            =   1410
            TabIndex        =   213
            Top             =   1320
            Width           =   300
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   1770
            Picture         =   "frmEditAll.frx":14B1E
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1590
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   2670
            Picture         =   "frmEditAll.frx":14F60
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1590
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   210
            Picture         =   "frmEditAll.frx":153A2
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   840
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   1080
            Picture         =   "frmEditAll.frx":157E4
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   840
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   2280
            Picture         =   "frmEditAll.frx":15C26
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1590
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   690
            Picture         =   "frmEditAll.frx":16068
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   840
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   3300
            Picture         =   "frmEditAll.frx":164AA
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   840
            Width           =   360
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   3690
            Picture         =   "frmEditAll.frx":168EC
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   840
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   2820
            Picture         =   "frmEditAll.frx":16D2E
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   840
            Width           =   480
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
            Left            =   3720
            TabIndex        =   212
            Top             =   1320
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdOrderBio 
         Caption         =   "Order Tests"
         Height          =   915
         Left            =   -65970
         Picture         =   "frmEditAll.frx":17170
         Style           =   1  'Graphical
         TabIndex        =   205
         Tag             =   "bOrder"
         Top             =   6870
         Width           =   765
      End
      Begin VB.CommandButton cmdAddBio 
         Caption         =   "&Add Result"
         Height          =   645
         Left            =   -68280
         Picture         =   "frmEditAll.frx":1747A
         Style           =   1  'Graphical
         TabIndex        =   204
         Tag             =   "bAdd"
         Top             =   5580
         Width           =   1035
      End
      Begin VB.ComboBox cParameter 
         Height          =   315
         Left            =   -73320
         TabIndex        =   203
         Text            =   "cParameter"
         Top             =   5190
         Width           =   1545
      End
      Begin VB.TextBox tResult 
         Height          =   315
         Left            =   -71700
         TabIndex        =   202
         Top             =   5190
         Width           =   1395
      End
      Begin VB.CommandButton cmdOrderCoag 
         Caption         =   "Order Tests"
         Height          =   945
         Left            =   -64650
         Picture         =   "frmEditAll.frx":184FC
         Style           =   1  'Graphical
         TabIndex        =   201
         Tag             =   "bOrder"
         Top             =   2280
         Width           =   1065
      End
      Begin VB.PictureBox Panel3D1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFF80&
         Height          =   1425
         Index           =   4
         Left            =   -74430
         ScaleHeight     =   1365
         ScaleWidth      =   7395
         TabIndex        =   197
         Top             =   5490
         Width           =   7455
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   90
            ScaleHeight     =   1185
            ScaleWidth      =   6975
            TabIndex        =   199
            Top             =   90
            Width           =   7005
            Begin VB.PictureBox pdelta 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   4695
               Left            =   0
               ScaleHeight     =   4695
               ScaleWidth      =   6945
               TabIndex        =   200
               Top             =   0
               Width           =   6945
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   1215
            LargeChange     =   500
            Left            =   7080
            Max             =   2500
            SmallChange     =   25
            TabIndex        =   198
            Top             =   90
            Width           =   255
         End
      End
      Begin VB.CommandButton bAddCoag 
         Caption         =   "&Add Result"
         Height          =   945
         Left            =   -68970
         Picture         =   "frmEditAll.frx":18806
         Style           =   1  'Graphical
         TabIndex        =   196
         Tag             =   "bAdd"
         Top             =   6840
         Width           =   1065
      End
      Begin VB.CommandButton bViewCoagRepeat 
         Caption         =   "View Repeat"
         Height          =   945
         Left            =   -66090
         Picture         =   "frmEditAll.frx":18C48
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   2280
         Width           =   1065
      End
      Begin VB.CommandButton bViewHaemRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   735
         Left            =   -65280
         Picture         =   "frmEditAll.frx":18DD2
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton cmdValidateHaem 
         Caption         =   "&Validate"
         Height          =   735
         Left            =   -65280
         Picture         =   "frmEditAll.frx":18F5C
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   5565
         Width           =   1155
      End
      Begin VB.CommandButton cmdValidateBio 
         Caption         =   "&Validate"
         Height          =   915
         Left            =   -64290
         Picture         =   "frmEditAll.frx":19266
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   6870
         Width           =   765
      End
      Begin VB.CommandButton cmdValidateCoag 
         Caption         =   "&Validate"
         Height          =   945
         Left            =   -64830
         Picture         =   "frmEditAll.frx":19570
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   6810
         Width           =   1065
      End
      Begin VB.CommandButton cmdSaveHoldDemographics 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   765
         Left            =   7320
         Picture         =   "frmEditAll.frx":1987A
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   5430
         Width           =   1155
      End
      Begin VB.Frame fr 
         Caption         =   "Specimen Comments"
         Height          =   1725
         Index           =   5
         Left            =   -74880
         TabIndex        =   184
         Top             =   6120
         Width           =   4635
         Begin VB.TextBox txtBioComment 
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
            Height          =   795
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   188
            Top             =   540
            Width           =   4515
         End
         Begin VB.ComboBox cmbBioComment 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   187
            Text            =   "cmbBioComment"
            Top             =   240
            Width           =   4515
         End
         Begin VB.ComboBox cmbBioComment 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   185
            Text            =   "cmbBioComment"
            Top             =   1380
            Width           =   4515
         End
      End
      Begin VB.Frame fr 
         Height          =   5595
         Index           =   0
         Left            =   390
         TabIndex        =   163
         Top             =   990
         Width           =   5475
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   3
            Left            =   1050
            MaxLength       =   20
            TabIndex        =   301
            Top             =   5100
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   2
            Left            =   1050
            MaxLength       =   20
            TabIndex        =   300
            Top             =   4770
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "...."
            Height          =   285
            Left            =   5010
            Style           =   1  'Graphical
            TabIndex        =   284
            Top             =   4440
            Width           =   390
         End
         Begin VB.ComboBox cmbHospital 
            Height          =   315
            Left            =   1050
            TabIndex        =   273
            Text            =   "cmbHospital"
            Top             =   1680
            Width           =   3915
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
            Left            =   5010
            Style           =   1  'Graphical
            TabIndex        =   261
            ToolTipText     =   "Copy To"
            Top             =   2490
            Width           =   375
         End
         Begin VB.TextBox txtDemographicComment 
            Height          =   915
            Left            =   1050
            MultiLine       =   -1  'True
            TabIndex        =   172
            Top             =   3450
            Width           =   3915
         End
         Begin VB.ComboBox cClDetails 
            Height          =   315
            Left            =   1050
            Sorted          =   -1  'True
            TabIndex        =   171
            Top             =   4410
            Width           =   3915
         End
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   1050
            TabIndex        =   168
            Text            =   "cmbWard"
            Top             =   2040
            Width           =   3915
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   0
            Left            =   750
            MaxLength       =   20
            TabIndex        =   167
            Top             =   1110
            Width           =   4215
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   1
            Left            =   750
            MaxLength       =   20
            TabIndex        =   166
            Top             =   1380
            Width           =   4215
         End
         Begin VB.ComboBox cmbClinician 
            Height          =   315
            Left            =   1050
            TabIndex        =   169
            Text            =   "cmbClinician"
            Top             =   2400
            Width           =   3915
         End
         Begin VB.ComboBox cmbGP 
            Height          =   315
            Left            =   1050
            TabIndex        =   170
            Text            =   "cmbGP"
            Top             =   2760
            Width           =   3915
         End
         Begin VB.ComboBox cmbDemogComment 
            Height          =   315
            Left            =   1050
            TabIndex        =   165
            Text            =   "cmbDemogComment"
            Top             =   3120
            Width           =   3915
         End
         Begin VB.TextBox txtExtSampleID 
            Height          =   285
            Left            =   3480
            MaxLength       =   10
            TabIndex        =   164
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label lblSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4260
            TabIndex        =   183
            Top             =   810
            Width           =   705
         End
         Begin VB.Label lblAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3180
            TabIndex        =   182
            Top             =   810
            Width           =   585
         End
         Begin VB.Label lblDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   181
            Top             =   810
            Width           =   1515
         End
         Begin VB.Label lblName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   750
            TabIndex        =   180
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label lblChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   179
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Cl Details"
            Height          =   945
            Index           =   35
            Left            =   60
            TabIndex        =   178
            Top             =   4410
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Chart #                                        Ext. SampleID"
            Height          =   195
            Index           =   26
            Left            =   90
            TabIndex        =   177
            Top             =   210
            Width           =   3330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Index           =   27
            Left            =   210
            TabIndex        =   176
            Top             =   510
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "D.o.B                                                Age                  Sex"
            Height          =   195
            Index           =   28
            Left            =   210
            TabIndex        =   175
            Top             =   840
            Width           =   3930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   174
            Top             =   1140
            Width           =   570
         End
         Begin VB.Label Label1 
            Caption         =   "Hospital                      Ward     Clinician                         GP       Comments"
            Height          =   1725
            Index           =   30
            Left            =   360
            TabIndex        =   173
            Top             =   1680
            Width           =   660
         End
      End
      Begin VB.PictureBox Panel3D1 
         Height          =   2265
         Index           =   1
         Left            =   -70650
         ScaleHeight     =   2205
         ScaleWidth      =   3615
         TabIndex        =   140
         Top             =   660
         Width           =   3675
         Begin VB.TextBox tRDWSD 
            Height          =   285
            Left            =   2430
            MaxLength       =   5
            TabIndex        =   285
            Top             =   1560
            Width           =   825
         End
         Begin VB.CommandButton bHgb 
            Caption         =   "Hgb"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   151
            Top             =   630
            Width           =   555
         End
         Begin VB.TextBox tMCHC 
            Height          =   285
            Left            =   2430
            MaxLength       =   5
            TabIndex        =   150
            Top             =   930
            Width           =   825
         End
         Begin VB.TextBox tRBC 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   149
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox tHgb 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   148
            Top             =   630
            Width           =   825
         End
         Begin VB.TextBox tHct 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   147
            Top             =   930
            Width           =   825
         End
         Begin VB.TextBox tMCH 
            Height          =   285
            Left            =   2430
            MaxLength       =   5
            TabIndex        =   146
            Top             =   630
            Width           =   825
         End
         Begin VB.TextBox tRDWCV 
            Height          =   285
            Left            =   2430
            MaxLength       =   5
            TabIndex        =   145
            Top             =   1230
            Width           =   825
         End
         Begin VB.TextBox tMCV 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   144
            Top             =   1230
            Width           =   825
         End
         Begin VB.TextBox tnrbcA 
            Height          =   285
            Left            =   750
            TabIndex        =   143
            Top             =   1890
            Width           =   825
         End
         Begin VB.TextBox tnrbcP 
            Height          =   285
            Left            =   1590
            TabIndex        =   142
            Top             =   1890
            Width           =   825
         End
         Begin VB.TextBox txtIRF 
            Height          =   285
            Left            =   2430
            TabIndex        =   141
            Top             =   330
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "RDWSD"
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
            Index           =   1
            Left            =   1665
            TabIndex        =   286
            Top             =   1590
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "RBC"
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
            Index           =   23
            Left            =   360
            TabIndex        =   162
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MCV"
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
            Index           =   53
            Left            =   330
            TabIndex        =   161
            Top             =   1260
            Width           =   405
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hct"
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
            Index           =   12
            Left            =   420
            TabIndex        =   160
            Top             =   960
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MCH"
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
            Index           =   14
            Left            =   2010
            TabIndex        =   159
            Top             =   660
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MCHC"
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
            Index           =   15
            Left            =   1890
            TabIndex        =   158
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "RDWCV"
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
            Index           =   42
            Left            =   1680
            TabIndex        =   157
            Top             =   1260
            Width           =   705
         End
         Begin VB.Label ipflag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Suspect"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   156
            Top             =   90
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label ipflag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abnormal"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   2730
            TabIndex        =   155
            Top             =   90
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "nRBC"
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
            Index           =   13
            Left            =   240
            TabIndex        =   154
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   16
            Left            =   2430
            TabIndex        =   153
            Top             =   1920
            Width           =   150
         End
         Begin VB.Label lblIRF 
            Caption         =   "IRF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2010
            TabIndex        =   152
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.PictureBox Panel3D1 
         Height          =   3195
         Index           =   0
         Left            =   -74550
         ScaleHeight     =   3135
         ScaleWidth      =   3615
         TabIndex        =   112
         Top             =   660
         Width           =   3675
         Begin VB.TextBox tBasA 
            Height          =   285
            Left            =   750
            TabIndex        =   124
            Top             =   2130
            Width           =   825
         End
         Begin VB.TextBox tEosP 
            Height          =   285
            Left            =   2430
            MaxLength       =   6
            TabIndex        =   123
            Top             =   1800
            Width           =   825
         End
         Begin VB.TextBox tNeutA 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   122
            Top             =   750
            Width           =   825
         End
         Begin VB.CommandButton bClearDiff 
            Caption         =   "Clear &Diff"
            Height          =   315
            Left            =   2160
            TabIndex        =   121
            Top             =   330
            Width           =   1095
         End
         Begin VB.TextBox tWBC 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   120
            Top             =   420
            Width           =   825
         End
         Begin VB.TextBox tEosA 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   119
            Top             =   1770
            Width           =   825
         End
         Begin VB.TextBox tNeutP 
            Height          =   285
            Left            =   2430
            MaxLength       =   6
            TabIndex        =   118
            Top             =   750
            Width           =   825
         End
         Begin VB.TextBox tMonoP 
            Height          =   285
            Left            =   2430
            MaxLength       =   6
            TabIndex        =   117
            Top             =   1455
            Width           =   825
         End
         Begin VB.TextBox tBasP 
            Height          =   285
            Left            =   2430
            MaxLength       =   6
            TabIndex        =   116
            Top             =   2130
            Width           =   825
         End
         Begin VB.TextBox tLymA 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   115
            Top             =   1110
            Width           =   825
         End
         Begin VB.TextBox tMonoA 
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   114
            Top             =   1440
            Width           =   825
         End
         Begin VB.TextBox tLymP 
            Height          =   285
            Left            =   2430
            MaxLength       =   6
            TabIndex        =   113
            Top             =   1110
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bas"
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
            Height          =   195
            Index           =   10
            Left            =   1830
            TabIndex        =   139
            Top             =   2160
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Eos"
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
            Height          =   195
            Index           =   9
            Left            =   1830
            TabIndex        =   138
            Top             =   1830
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   22
            Left            =   3240
            TabIndex        =   137
            Top             =   780
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "WBC"
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
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   136
            Top             =   480
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Lymph"
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
            Height          =   195
            Index           =   8
            Left            =   1665
            TabIndex        =   135
            Top             =   1170
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Mono"
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
            Height          =   195
            Index           =   40
            Left            =   1740
            TabIndex        =   134
            Top             =   1500
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Neut"
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
            Height          =   195
            Index           =   41
            Left            =   1770
            TabIndex        =   133
            Top             =   780
            Width           =   420
         End
         Begin VB.Label ipflag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abnormal"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   0
            Left            =   2730
            TabIndex        =   132
            Top             =   90
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label ipflag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Suspect"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   131
            Top             =   90
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "WIC"
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
            Height          =   195
            Index           =   11
            Left            =   300
            TabIndex        =   130
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "WOC"
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
            Height          =   195
            Index           =   24
            Left            =   1920
            TabIndex        =   129
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lWIC 
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
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   750
            TabIndex        =   128
            Top             =   2490
            Width           =   810
         End
         Begin VB.Label lWOC 
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
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   2430
            TabIndex        =   127
            Top             =   2490
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "WVF"
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
            Height          =   195
            Index           =   25
            Left            =   1950
            TabIndex        =   126
            Top             =   2880
            Width           =   420
         End
         Begin VB.Label lblWVF 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2430
            TabIndex        =   125
            Top             =   2820
            Width           =   825
         End
      End
      Begin VB.PictureBox Panel3D1 
         Height          =   855
         Index           =   3
         Left            =   -70650
         ScaleHeight     =   795
         ScaleWidth      =   3615
         TabIndex        =   105
         Top             =   2970
         Width           =   3675
         Begin VB.TextBox tMPV 
            Height          =   285
            Left            =   2340
            TabIndex        =   107
            Top             =   360
            Width           =   825
         End
         Begin VB.TextBox tPlt 
            Height          =   285
            Left            =   720
            MaxLength       =   5
            TabIndex        =   106
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Plt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   195
            Index           =   17
            Left            =   285
            TabIndex        =   111
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MPV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   195
            Index           =   18
            Left            =   1800
            TabIndex        =   110
            Top             =   420
            Width           =   405
         End
         Begin VB.Label ipflag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abnormal"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   4
            Left            =   2730
            TabIndex        =   109
            Top             =   90
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label ipflag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Suspect"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   108
            Top             =   90
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.PictureBox Panel3D1 
         Height          =   3825
         Index           =   2
         Left            =   -66180
         ScaleHeight     =   3765
         ScaleWidth      =   2085
         TabIndex        =   88
         Top             =   660
         Width           =   2145
         Begin VB.CommandButton cmdPrintMalaria 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   1500
            Picture         =   "frmEditAll.frx":19EE4
            Style           =   1  'Graphical
            TabIndex        =   274
            ToolTipText     =   "Print Sickledex Only"
            Top             =   2220
            Width           =   555
         End
         Begin VB.CommandButton cmdPrintSickledex 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   1500
            Picture         =   "frmEditAll.frx":1A54E
            Style           =   1  'Graphical
            TabIndex        =   272
            ToolTipText     =   "Print Sickledex Only"
            Top             =   2880
            Width           =   555
         End
         Begin VB.CommandButton bprintesr 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   1500
            Picture         =   "frmEditAll.frx":1ABB8
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Print ESR Only"
            Top             =   270
            Width           =   555
         End
         Begin VB.TextBox tRetP 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   99
            Top             =   990
            Width           =   675
         End
         Begin VB.TextBox tESR 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            MaxLength       =   7
            TabIndex        =   98
            Top             =   360
            Width           =   1035
         End
         Begin VB.CheckBox cESR 
            Caption         =   "ESR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   150
            Width           =   765
         End
         Begin VB.CheckBox cRetics 
            Caption         =   "Retics"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   96
            Top             =   780
            Width           =   885
         End
         Begin VB.TextBox tMonospot 
            Height          =   285
            Left            =   240
            TabIndex        =   95
            Top             =   1620
            Width           =   1065
         End
         Begin VB.CheckBox cMonospot 
            Caption         =   "Infectious Mono screen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   240
            TabIndex        =   94
            Top             =   1275
            Width           =   1515
         End
         Begin VB.TextBox tRetA 
            Height          =   285
            Left            =   240
            TabIndex        =   93
            Top             =   990
            Width           =   705
         End
         Begin VB.CheckBox chkMalaria 
            Caption         =   "Malaria Screen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   92
            Top             =   2040
            Width           =   1635
         End
         Begin VB.CheckBox chkSickledex 
            Caption         =   "Sickledex Screen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   91
            Top             =   2670
            Width           =   1815
         End
         Begin VB.CheckBox chkRA 
            Caption         =   "RA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   90
            Top             =   3240
            Width           =   585
         End
         Begin VB.CommandButton cmdPrintMonospot 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   1500
            Picture         =   "frmEditAll.frx":1B222
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Print Monospot Only"
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   49
            Left            =   1710
            TabIndex        =   104
            Top             =   1020
            Width           =   210
         End
         Begin VB.Label lblMalaria 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   240
            TabIndex        =   103
            Top             =   2250
            Width           =   1095
         End
         Begin VB.Label lblSickledex 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   240
            TabIndex        =   102
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lblRA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   3450
            Width           =   1095
         End
      End
      Begin VB.TextBox txtHaemComment 
         Height          =   555
         Left            =   -74100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Top             =   4860
         Width           =   7125
      End
      Begin VB.CommandButton bViewBioRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Left            =   -63450
         Picture         =   "frmEditAll.frx":1B88C
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   6870
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtCoagComment 
         BackColor       =   &H80000018&
         Height          =   1185
         Left            =   -69030
         MultiLine       =   -1  'True
         TabIndex        =   85
         Top             =   3570
         Width           =   5445
      End
      Begin VB.Frame fr 
         Caption         =   "Warfarin"
         Height          =   855
         Index           =   8
         Left            =   -66120
         TabIndex        =   82
         Top             =   780
         Width           =   2535
         Begin VB.TextBox tWarfarin 
            Height          =   285
            Left            =   270
            MaxLength       =   5
            TabIndex        =   84
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton bPrintINR 
            Caption         =   "Print INR"
            Height          =   285
            Left            =   1290
            TabIndex        =   83
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton bHaemGraphs 
         Height          =   735
         Left            =   -66540
         Picture         =   "frmEditAll.frx":1BA16
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   4665
         Width           =   1155
      End
      Begin VB.CommandButton bOrderTests 
         Caption         =   "Order Tests"
         Height          =   1215
         Left            =   10860
         Picture         =   "frmEditAll.frx":1BE58
         Style           =   1  'Graphical
         TabIndex        =   80
         Tag             =   "bOrder"
         Top             =   2160
         Width           =   765
      End
      Begin VB.CheckBox cFilm 
         Caption         =   "Film"
         Height          =   195
         Left            =   -73680
         TabIndex        =   79
         Top             =   3990
         Width           =   585
      End
      Begin VB.CommandButton bPrintAll 
         Caption         =   "Print All"
         Height          =   945
         Left            =   -67560
         Picture         =   "frmEditAll.frx":1C162
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   6810
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdDeleteExt 
         Caption         =   "Delete"
         Height          =   735
         Left            =   -64140
         Picture         =   "frmEditAll.frx":1C7CC
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2730
         Width           =   1335
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -74730
         TabIndex        =   76
         Top             =   4350
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   -74730
         TabIndex        =   75
         Top             =   5400
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   -74730
         TabIndex        =   74
         Top             =   5130
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   -74730
         TabIndex        =   73
         Top             =   4860
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   -74730
         TabIndex        =   72
         Top             =   4620
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   -74730
         TabIndex        =   71
         Top             =   5670
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   -74730
         TabIndex        =   70
         Top             =   5910
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   -74730
         TabIndex        =   69
         Top             =   6150
         Width           =   10275
      End
      Begin VB.TextBox etc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   -74730
         TabIndex        =   68
         Top             =   6420
         Width           =   10275
      End
      Begin VB.CommandButton cmdOrderExt 
         Caption         =   "Order External Tests"
         Height          =   1035
         Index           =   0
         Left            =   -64140
         Picture         =   "frmEditAll.frx":1CC0E
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3660
         Width           =   1335
      End
      Begin VB.CommandButton cmdDiff 
         Caption         =   "Man. Diff"
         Height          =   315
         Left            =   -74550
         TabIndex        =   66
         Top             =   3960
         Width           =   825
      End
      Begin VB.CommandButton cmdOrderExt 
         Caption         =   "Order External Tests"
         Height          =   1215
         Index           =   1
         Left            =   11670
         Picture         =   "frmEditAll.frx":1CF18
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   2145
         Width           =   765
      End
      Begin VB.TextBox txtFilmComment 
         Height          =   525
         Left            =   -74100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   4320
         Width           =   7125
      End
      Begin VB.ComboBox cmbHaemComment 
         Height          =   315
         Left            =   -69930
         TabIndex        =   63
         Text            =   "cmbHaemComment"
         Top             =   3960
         Width           =   2955
      End
      Begin VB.ComboBox cmbFilmComment 
         Height          =   315
         Left            =   -72960
         TabIndex        =   62
         Text            =   "cmbFilmComment"
         Top             =   3960
         Width           =   2085
      End
      Begin VB.CommandButton cmdCopyFromPrevious 
         BackColor       =   &H00FF80FF&
         Caption         =   "Copy all Details from Sample # 123456789"
         Height          =   585
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   510
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cmbNewResult 
         Height          =   315
         ItemData        =   "frmEditAll.frx":1D222
         Left            =   -73140
         List            =   "frmEditAll.frx":1D224
         TabIndex        =   60
         Text            =   "cmbNewResult"
         ToolTipText     =   "Type Result or select from list."
         Top             =   5730
         Width           =   1725
      End
      Begin VB.PictureBox pbCD 
         BackColor       =   &H00C0FFFF&
         Height          =   2355
         Left            =   -63990
         ScaleHeight     =   2295
         ScaleWidth      =   1395
         TabIndex        =   45
         Top             =   4995
         Visible         =   0   'False
         Width           =   1455
         Begin VB.TextBox txtCD3A 
            Height          =   285
            Left            =   30
            TabIndex        =   52
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox txtCD3P 
            Height          =   285
            Left            =   630
            TabIndex        =   51
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox txtCD4A 
            Height          =   285
            Left            =   30
            TabIndex        =   50
            Top             =   840
            Width           =   585
         End
         Begin VB.TextBox txtCD4P 
            Height          =   285
            Left            =   630
            TabIndex        =   49
            Top             =   840
            Width           =   585
         End
         Begin VB.TextBox txtCD8A 
            Height          =   285
            Left            =   30
            TabIndex        =   48
            Top             =   1440
            Width           =   585
         End
         Begin VB.TextBox txtCD8P 
            Height          =   285
            Left            =   630
            TabIndex        =   47
            Top             =   1440
            Width           =   585
         End
         Begin VB.TextBox txtCD48 
            Height          =   285
            Left            =   30
            TabIndex        =   46
            Top             =   1980
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "CD3"
            Height          =   195
            Index           =   63
            Left            =   90
            TabIndex        =   59
            Top             =   60
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "CD4"
            Height          =   195
            Index           =   64
            Left            =   90
            TabIndex        =   58
            Top             =   645
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "CD8"
            Height          =   195
            Index           =   65
            Left            =   90
            TabIndex        =   57
            Top             =   1215
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "4/8"
            Height          =   195
            Index           =   66
            Left            =   90
            TabIndex        =   56
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "%"
            Height          =   195
            Index           =   67
            Left            =   1230
            TabIndex        =   55
            Top             =   270
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "%"
            Height          =   195
            Index           =   68
            Left            =   1230
            TabIndex        =   54
            Top             =   1470
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "%"
            Height          =   195
            Index           =   69
            Left            =   1230
            TabIndex        =   53
            Top             =   870
            Width           =   120
         End
      End
      Begin VB.ComboBox cmbBioComment 
         Height          =   315
         Index           =   2
         Left            =   -69030
         TabIndex        =   44
         Text            =   "cmbBioComment"
         Top             =   4800
         Width           =   5385
      End
      Begin VB.Frame fraSelectPrint 
         Height          =   435
         Index           =   1
         Left            =   -66660
         TabIndex        =   40
         Top             =   315
         Width           =   2085
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   0
            Left            =   1740
            Picture         =   "frmEditAll.frx":1D226
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   0
            Left            =   1410
            Picture         =   "frmEditAll.frx":1D4FC
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Select for Printing"
            Height          =   195
            Index           =   44
            Left            =   45
            TabIndex        =   43
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame fraSelectPrint 
         Height          =   435
         Index           =   0
         Left            =   -71280
         TabIndex        =   36
         Top             =   420
         Width           =   2085
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   1
            Left            =   1410
            Picture         =   "frmEditAll.frx":1D7D2
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   1
            Left            =   1740
            Picture         =   "frmEditAll.frx":1DAA8
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Select for Printing"
            Height          =   195
            Index           =   43
            Left            =   45
            TabIndex        =   39
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.ListBox lstAdd 
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   6030
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveDemographics 
         Caption         =   "&Save"
         Height          =   765
         Left            =   8610
         Picture         =   "frmEditAll.frx":1DD7E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5430
         Width           =   1155
      End
      Begin VB.CommandButton cmdSaveHaem 
         Caption         =   "&Save"
         Height          =   735
         Left            =   -66540
         Picture         =   "frmEditAll.frx":1F700
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5565
         Width           =   1155
      End
      Begin VB.CommandButton cmdSaveBio 
         Caption         =   "&Save"
         Height          =   915
         Left            =   -65130
         Picture         =   "frmEditAll.frx":21082
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6870
         Width           =   765
      End
      Begin MSFlexGridLib.MSFlexGrid grdOutstanding 
         Height          =   4500
         Left            =   -63690
         TabIndex        =   189
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   7938
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding |<Code"
      End
      Begin MSFlexGridLib.MSFlexGrid gBio 
         Height          =   4500
         Left            =   -74880
         TabIndex        =   229
         Top             =   840
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7938
         _Version        =   393216
         Cols            =   10
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
         FormatString    =   $"frmEditAll.frx":22A04
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
      Begin MSFlexGridLib.MSFlexGrid gCoag 
         Height          =   4275
         Left            =   -73350
         TabIndex        =   230
         Top             =   870
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7541
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
         FormatString    =   "<Parameter            |<Result    |<Units       |<Flag|^VP |^P "
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
      Begin MSFlexGridLib.MSFlexGrid gOutstandingCoag 
         Height          =   4275
         Left            =   -74790
         TabIndex        =   231
         Top             =   870
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding  "
      End
      Begin MSFlexGridLib.MSFlexGrid grdPrev 
         Height          =   2175
         Left            =   -69030
         TabIndex        =   232
         Top             =   1110
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         FormatString    =   "<Parameter            |<Result                 "
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   3885
         Left            =   -74760
         TabIndex        =   233
         Top             =   450
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6853
         _Version        =   393216
         Cols            =   6
         RowHeightMin    =   800
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         AllowUserResizing=   3
         FormatString    =   $"frmEditAll.frx":22AA0
      End
      Begin MSFlexGridLib.MSFlexGrid gOutstandingHaem 
         Height          =   3870
         Left            =   -63885
         TabIndex        =   287
         Top             =   585
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   6826
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding |<Code"
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
         Height          =   375
         Left            =   390
         TabIndex        =   306
         Top             =   630
         Width           =   4185
      End
      Begin VB.Label lblDelta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   945
         Index           =   0
         Left            =   -69030
         TabIndex        =   277
         ToolTipText     =   "Delta Check"
         Top             =   5550
         Width           =   5445
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   $"frmEditAll.frx":22B5C
         Height          =   615
         Left            =   -74700
         TabIndex        =   278
         Top             =   5310
         Width           =   7830
      End
      Begin VB.Label lblDemogValid 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Demographics Not Valid"
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   5970
         TabIndex        =   266
         Top             =   5430
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGreaterThan 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensure all xx Test Results are reviewed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -66300
         TabIndex        =   259
         Top             =   6240
         Width           =   3780
      End
      Begin VB.Label lblDelta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   915
         Index           =   2
         Left            =   -70200
         TabIndex        =   258
         ToolTipText     =   "Delta Check"
         Top             =   6900
         Width           =   3645
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Split 6"
         Height          =   255
         Index           =   6
         Left            =   -68430
         TabIndex        =   257
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Specimen Comments"
         Height          =   225
         Index           =   46
         Left            =   -69030
         TabIndex        =   254
         Top             =   3360
         Width           =   1515
      End
      Begin VB.Label lblHaemPrinted 
         AutoSize        =   -1  'True
         Caption         =   "Already Printed"
         Height          =   195
         Left            =   -66510
         TabIndex        =   253
         Top             =   6375
         Width           =   1065
      End
      Begin VB.Label lRandom 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Random Sample"
         Height          =   285
         Left            =   -72540
         TabIndex        =   252
         ToolTipText     =   "Click to Toggle"
         Top             =   450
         Width           =   1875
      End
      Begin VB.Label lHaemErrors 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F L A G S "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1845
         Left            =   -66750
         TabIndex        =   251
         ToolTipText     =   "Click to show errors"
         Top             =   660
         Width           =   315
      End
      Begin VB.Label lblPrevCoag 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Chart # for Previous Details"
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   -69000
         TabIndex        =   250
         Top             =   870
         Width           =   2715
      End
      Begin VB.Label lblAss 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Associated  Glucose 1"
         Height          =   825
         Left            =   -64620
         TabIndex        =   249
         Top             =   5370
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Already Validated"
         Height          =   195
         Index           =   36
         Left            =   -65310
         TabIndex        =   248
         Top             =   6375
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Film Comment"
         Height          =   495
         Index           =   37
         Left            =   -74790
         TabIndex        =   247
         Top             =   4320
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Specimen Comment"
         Height          =   435
         Index           =   39
         Left            =   -70710
         TabIndex        =   246
         Top             =   3930
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Specimen Comment"
         Height          =   435
         Index           =   38
         Left            =   -74880
         TabIndex        =   245
         Top             =   4920
         Width           =   840
      End
      Begin VB.Label lblAnalyser 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sapphire - A"
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
         Left            =   -63990
         TabIndex        =   244
         Top             =   4665
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Analyser"
         Height          =   195
         Index           =   61
         Left            =   -63990
         TabIndex        =   243
         Top             =   4455
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Analyser"
         Height          =   195
         Index           =   62
         Left            =   -66090
         TabIndex        =   242
         Top             =   1830
         Width           =   600
      End
      Begin VB.Label lblCoagAnalyser 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   420
         Left            =   -65430
         TabIndex        =   241
         Top             =   1710
         Width           =   1830
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   -74850
         TabIndex        =   240
         Top             =   5280
         Width           =   510
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Split 1"
         Height          =   255
         Index           =   1
         Left            =   -74310
         TabIndex        =   239
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Split 2"
         Height          =   255
         Index           =   2
         Left            =   -73110
         TabIndex        =   238
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Split 3"
         Height          =   255
         Index           =   3
         Left            =   -71940
         TabIndex        =   237
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Split 4"
         Height          =   255
         Index           =   4
         Left            =   -70770
         TabIndex        =   236
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Split 5"
         Height          =   255
         Index           =   5
         Left            =   -69600
         TabIndex        =   235
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label lblSplitView 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Viewing All"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -74850
         TabIndex        =   234
         Top             =   450
         Width           =   2220
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gMDemoLabNoUpd 
      Height          =   3255
      Left            =   14400
      TabIndex        =   283
      Top             =   1980
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5741
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
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmEditAll.frx":22CBB
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmEditAll.frx":22F91
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblDateConflict 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample Date 88/88/8888 Run Date 88/88/8888"
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
      Height          =   885
      Left            =   12840
      TabIndex        =   30
      Top             =   7560
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Menu mnuListsBio 
      Caption         =   "&Lists"
      Begin VB.Menu mnuLIH 
         Caption         =   "LIH Values"
      End
      Begin VB.Menu MnuAssignPanelsForHealthLink 
         Caption         =   "Assign Panels for HealthLink"
      End
      Begin VB.Menu mnuSplits 
         Caption         =   "&Splits"
      End
      Begin VB.Menu mnuNewResult 
         Caption         =   "&New Result"
      End
      Begin VB.Menu mnuTestSequence 
         Caption         =   "Print &Sequence"
      End
      Begin VB.Menu mnuNormalFlag 
         Caption         =   "Normal && &Flag Ranges"
      End
      Begin VB.Menu mnuPlausible 
         Caption         =   "&Plausible Range"
      End
      Begin VB.Menu mnuInUse 
         Caption         =   "&In Use"
      End
      Begin VB.Menu mnuKnownToAnalyser 
         Caption         =   "&Known to Analyser"
      End
      Begin VB.Menu mnuDelta 
         Caption         =   "&Delta Check"
      End
      Begin VB.Menu mnuAutoVal 
         Caption         =   "&Auto Validation"
      End
      Begin VB.Menu mnuMasks 
         Caption         =   "&Masks"
      End
      Begin VB.Menu mnuUnitsPrecision 
         Caption         =   "&Units && Precision"
      End
      Begin VB.Menu mnuReRunTimes 
         Caption         =   "&Re-Run Times"
      End
      Begin VB.Menu mnuFasting 
         Caption         =   "&Fasting Ranges"
      End
      Begin VB.Menu mnuNewAnalyte 
         Caption         =   "&New Analyte"
      End
      Begin VB.Menu mnuAmendAnalyte 
         Caption         =   "&Amend Existing Analyte"
      End
      Begin VB.Menu mnuTestCodeMappingBio 
         Caption         =   "&Test Code Mapping"
      End
   End
   Begin VB.Menu mnuListsHaem 
      Caption         =   "&Lists"
      Begin VB.Menu mnuAutoValHaem 
         Caption         =   "&AutoValidation"
      End
      Begin VB.Menu mnuHaemDefinitions 
         Caption         =   "&Normal Ranges"
      End
      Begin VB.Menu mnuBarCodesH 
         Caption         =   "&Bar Codes"
      End
      Begin VB.Menu mnuTestCodeMappingHaem 
         Caption         =   "&Test Code Mapping"
      End
   End
   Begin VB.Menu mnuListsCoag 
      Caption         =   "&Lists"
      Begin VB.Menu mnuAddCoagTest 
         Caption         =   "&Add Test"
      End
      Begin VB.Menu mnuCoagDefinitions 
         Caption         =   "&Normal Ranges "
      End
      Begin VB.Menu mnuCoagPanels 
         Caption         =   "&Panels"
      End
      Begin VB.Menu mnuTestCodeMappingCoag 
         Caption         =   "&Test Code Mapping"
      End
   End
   Begin VB.Menu mnuNull 
      Caption         =   ""
   End
End
Attribute VB_Name = "frmEditAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_ShowDoc As Boolean
Public m_Notes As String

Private UpDownDirection As Integer
'+++Abubaker 16-11-2023
'Dim m_Counter As Integer
'Dim m_SampleID As String
'---Abubaker 16-11-2023

Private mNewRecord As Boolean

'Private PreviousBioSID As String
Private PreviousHaem As Boolean
Private PreviousCoag As Boolean
Private PreviousBio As Boolean

Private HaemLoaded As Boolean
Private BioLoaded As Boolean
Private CoagLoaded As Boolean
Private ExtLoaded As Boolean

Private Activated As Boolean
Private UrgentTest As Boolean

Private pPrintToPrinter As String

Private HaemValBy As String
Private BioValBy As String
Private CoagValBy As String
Private ExtValBy As String

Private frmOptDeptHaem As Boolean
Private frmOptDeptBio As Boolean
Private frmOptDeptCoag As Boolean
Private frmOptDeptExt As Boolean
Private frmOptUrgent As Boolean
Private frmOptBloodBank As Boolean
Private frmOptAllowClinicianFreeText As Boolean

Private CodeForGlucose As String
Private CodeForChol As String
Private CodeForTrig As String

Private Opts(0 To 13) As udtOptionList

Private LiIcHas As New LIHs

Private m_StartInDepartment As String

Private bds As New BIEDefinitions

Dim PrevDateHaem As String
Dim PrevIDHaem As String
Private MatchingDemoLoaded As Boolean
Private FormLoaded As Boolean

Private m_bSavedDemoFromGPCom As Boolean
Private m_bCancelFromGpCom As Boolean

Private Const fcLine_NO = 0
Private Const fcQus = 1
Private Const fcAns = 2
Private Const fcRID = 3



Private Sub ClearDifferential()

37040     lWIC = ""
37050     lWOC = ""

37060     tLymA = ""
37070     tLymA.BackColor = &HFFFFFF
37080     tLymA.ForeColor = &H0&

37090     tLymP = ""
37100     tLymP.BackColor = &HFFFFFF
37110     tLymP.ForeColor = &H0&

37120     tMonoA = ""
37130     tMonoA.BackColor = &HFFFFFF
37140     tMonoA.ForeColor = &H0&

37150     tMonoP = ""
37160     tMonoP.BackColor = &HFFFFFF
37170     tMonoP.ForeColor = &H0&

37180     tNeutA = ""
37190     tNeutA.BackColor = &HFFFFFF
37200     tNeutA.ForeColor = &H0&

37210     tNeutP = ""
37220     tNeutP.BackColor = &HFFFFFF
37230     tNeutP.ForeColor = &H0&

37240     tEosA = ""
37250     tEosA.BackColor = &HFFFFFF
37260     tEosA.ForeColor = &H0&

37270     tEosP = ""
37280     tEosP.BackColor = &HFFFFFF
37290     tEosP.ForeColor = &H0&

37300     tBasA = ""
37310     tBasA.BackColor = &HFFFFFF
37320     tBasA.ForeColor = &H0&

37330     tBasP = ""
37340     tBasP.BackColor = &HFFFFFF
37350     tBasP.ForeColor = &H0&

End Sub

Private Sub ClearHaemExceptHgb()

37360     ClearDifferential

37370     tnrbcA = ""
37380     tnrbcP = ""

37390     tWBC = ""
37400     tWBC.BackColor = &HFFFFFF
37410     tWBC.ForeColor = &H0&

37420     tRBC = ""
37430     tRBC.BackColor = &HFFFFFF
37440     tRBC.ForeColor = &H0&

37450     tMCV = ""
37460     tMCV.BackColor = &HFFFFFF
37470     tMCV.ForeColor = &H0&

37480     tHct = ""
37490     tHct.BackColor = &HFFFFFF
37500     tHct.ForeColor = &H0&

37510     tRDWCV = ""
37520     tRDWCV.BackColor = &HFFFFFF
37530     tRDWCV.ForeColor = &H0&


37540     tRDWSD = ""
37550     tRDWSD.BackColor = &HFFFFFF
37560     tRDWSD.ForeColor = &H0&


37570     tMCH = ""
37580     tMCH.BackColor = &HFFFFFF
37590     tMCH.ForeColor = &H0&

37600     tMCHC = ""
37610     tMCHC.BackColor = &HFFFFFF
37620     tMCHC.ForeColor = &H0&

37630     tPlt = ""
37640     tPlt.BackColor = &HFFFFFF
37650     tPlt.ForeColor = &H0&

37660     tMPV = ""
37670     tMPV.BackColor = &HFFFFFF
37680     tMPV.ForeColor = &H0&

End Sub
Private Sub CheckTag()

          Dim tb As Recordset
          Dim sql As String

37690     On Error GoTo CheckTag_Error

37700     cmdTag.BackColor = vbButtonFace

37710     sql = "SELECT COUNT(*) Tot FROM MicroTag WHERE Chart = '" & txtChart & "'"
37720     Set tb = New Recordset
37730     RecOpenServer 0, tb, sql
37740     If tb!Tot > 0 Then
37750         cmdTag.BackColor = vbRed
37760     End If

37770     Exit Sub

CheckTag_Error:

          Dim strES As String
          Dim intEL As Integer

37780     intEL = Erl
37790     strES = Err.Description
37800     LogError "frmEditAll", "CheckTag", intEL, strES, sql

End Sub



Private Sub AdjustUCreat(ByRef BR As BIEResult)

          Dim sql As String

37810     On Error GoTo AdjustUCreat_Error

37820     If BR.LongName = "Urinary Creatinine" Then
37830         If IsNumeric(BR.Result) Then
37840             If Val(BR.Result) < 50 Then
37850                 BR.Result = Val(BR.Result) * 1000
37860                 sql = "UPDATE BioResults " & _
                          "SET Result = '" & BR.Result & "' " & _
                          "WHERE SampleID = '" & BR.SampleID & "' " & _
                          "AND Code = '" & BR.Code & "'"
37870                 Cnxn(0).Execute sql
37880             End If
37890         End If
37900     End If

37910     Exit Sub

AdjustUCreat_Error:

          Dim strES As String
          Dim intEL As Integer

37920     intEL = Erl
37930     strES = Err.Description
37940     LogError "frmEditAll", "AdjustUCreat", intEL, strES, sql

End Sub

Private Function MaskInhibit(ByVal BR As BIEResult, ByVal BRs As BIEResults) As String

          Dim Lx As LIH
          Dim RetVal As String
          Dim Result As Single
          Dim BRLIH As BIEResult
          Dim CutOffForThisParameter As Single
          Dim LIHValue As Single
          Dim sql As String
37950     On Error GoTo ErrorHandler
37960     RetVal = ""

37970     Set Lx = LiIcHas.Item("L", BR.Code, "P")
37980     If Not Lx Is Nothing Then
37990         Set BRLIH = BRs.Item("1073")
38000         If Not BRLIH Is Nothing Then
38010             CutOffForThisParameter = Lx.CutOff
38020             If CutOffForThisParameter > 0 Then
38030                 LIHValue = BRLIH.Result
38040                 If LIHValue >= CutOffForThisParameter Then
38050                     RetVal = "XL"
38060                 End If
38070             End If
38080         End If
38090     End If

38100     If RetVal = "" Then
38110         Set Lx = LiIcHas.Item("I", BR.Code, "P")
38120         If Not Lx Is Nothing Then
38130             Set BRLIH = BRs.Item("1072")
38140             If Not BRLIH Is Nothing Then
38150                 CutOffForThisParameter = Lx.CutOff
38160                 If CutOffForThisParameter > 0 Then
38170                     LIHValue = BRLIH.Result
38180                     If LIHValue >= CutOffForThisParameter Then
38190                         RetVal = "XI"
38200                     End If
38210                 End If
38220             End If
38230         End If
38240     End If

38250     If RetVal = "" Then
38260         Set Lx = LiIcHas.Item("H", BR.Code, "P")
38270         If Not Lx Is Nothing Then
38280             Set BRLIH = BRs.Item("1071")
38290             If Not BRLIH Is Nothing Then
38300                 CutOffForThisParameter = Lx.CutOff
38310                 If CutOffForThisParameter > 0 Then
38320                     LIHValue = BRLIH.Result
38330                     If LIHValue >= CutOffForThisParameter Then
38340                         RetVal = "XH"
38350                     End If
38360                 End If
38370             End If
38380         End If
38390     End If

38400     If RetVal <> "" Then
38410         sql = "UPDATE BioResults SET Result = 'XXXXX' WHERE SampleID = '" & txtSampleID & "' AND Code = '" & BR.Code & "'"
38420         Cnxn(0).Execute sql
38430         BR.Result = "XXXXX"
38440     End If

38450     MaskInhibit = RetVal
38460     Exit Function
ErrorHandler:
          '    MsgBox Err.Description
End Function

Private Function MaskResult(ByVal Code As String) As Boolean

          Dim sql As String
          Dim tb As Recordset

38470     On Error GoTo MaskResult_Error

38480     MaskResult = False

38490     sql = "SELECT M.SampleID FROM Masks M " & _
              "JOIN BioTestDefinitions T " & _
              "ON M.[O] = T.[O] " & _
              "WHERE M.SampleID = '" & txtSampleID & "' " & _
              "AND M.[O] = 1 " & _
              "AND T.Code = '" & Code & "' "
38500     Set tb = New Recordset
38510     RecOpenServer 0, tb, sql
38520     If Not tb.EOF Then
38530         MaskResult = True
38540     End If

38550     Exit Function

MaskResult_Error:

          Dim strES As String
          Dim intEL As Integer

38560     intEL = Erl
38570     strES = Err.Description
38580     LogError "frmEditAll", "MaskResult", intEL, strES, sql

End Function

Private Function CheckEGFR(ByVal BRs As BIEResults) As Boolean
          'returns True if eGFR added

          Dim BR As BIEResult
          Dim Code As String
          Dim EGFR As String
          Dim Rundate As String
          Dim RunTime As String
          Dim bNew As BIEResult
          Dim sql As String
          Dim tb As Recordset
          Dim EGFRCode As String
          Dim CodeForCreat As String

38590     On Error GoTo CheckEGFR_Error

38600     CheckEGFR = False

38610     If BRs Is Nothing Then Exit Function

38620     If Not egfrInclude(cmbWard, cmbClinician, cmbGP) Then Exit Function

38630     EGFRCode = GetOptionSetting("BioCodeForEGFR", "5555")
38640     CodeForCreat = GetOptionSetting("BioCodeForCreatinine", "234")
38650     For Each BR In BRs
38660         Code = UCase$(Trim$(BR.Code))
38670         If Code = EGFRCode Then    '"5555"
38680             Exit Function
38690         End If
38700     Next
          '+++ Junaid
38710     Exit Function
          '--- Junaid
38720     For Each BR In BRs
38730         Code = UCase$(Trim$(BR.Code))
38740         If Code = CodeForCreat Then    'And Not BR.Printed Then
38750             EGFR = CalculateEGFR(BR.Result)
38760             If Val(EGFR) <> Val(BR.Result) Then
38770                 If IsDate(BR.Rundate) Then
38780                     Rundate = BR.Rundate
38790                 Else
38800                     Rundate = Format$(BR.RunTime, "dd/mmm/yyyy")
38810                 End If
38820                 RunTime = BR.RunTime
38830                 sql = "SELECT * FROM BioResults WHERE " & _
                          "SampleID = '" & txtSampleID & "' " & _
                          "AND Code = '" & EGFRCode & "'"
38840                 Set tb = New Recordset
38850                 RecOpenClient 0, tb, sql
38860                 If tb.EOF Then
38870                     tb.AddNew
38880                 End If
38890                 tb!SampleID = txtSampleID
38900                 tb!Rundate = Rundate
38910                 tb!RunTime = RunTime
38920                 tb!Code = EGFRCode    '5555
38930                 tb!Result = EGFR
38940                 tb!Units = "ml/min/1.73m2"
38950                 tb!Printed = BR.Printed
38960                 tb!Valid = BR.Valid
38970                 tb!FAXed = 0
38980                 tb!Analyser = ""
38990                 tb!SampleType = "S"
39000                 tb.Update

39010                 Set bNew = New BIEResult
39020                 bNew.SampleID = txtSampleID
39030                 bNew.Code = EGFRCode    '"5555"
39040                 bNew.Rundate = Rundate
39050                 bNew.RunTime = RunTime
39060                 bNew.Result = EGFR
39070                 bNew.Units = "ml/min/1.73m2"
39080                 bNew.Printed = BR.Printed
39090                 bNew.Valid = BR.Valid
39100                 bNew.SampleType = "S"
39110                 bNew.LongName = "eGFR"
39120                 bNew.ShortName = "eGFR"
39130                 bNew.PlausibleLow = 0
39140                 bNew.PlausibleHigh = 9999
39150                 bNew.FlagLow = 0
39160                 bNew.FlagHigh = 9999
39170                 BRs.Add bNew

39180                 CheckEGFR = True
39190                 Exit For
39200             End If
39210         End If
39220     Next

39230     Exit Function

CheckEGFR_Error:

          Dim strES As String
          Dim intEL As Integer

39240     intEL = Erl
39250     strES = Err.Description
39260     LogError "frmEditAll", "CheckEGFR", intEL, strES, sql

End Function

Private Function CalculateEGFR(ByVal Creat As String) As String

          Dim EGFR As Long
          Dim pAge As Long

39270     On Error GoTo CalculateEGFR_Error
          '+++ Junaid
39280     Exit Function
          '--- Junaid
39290     CalculateEGFR = Creat

39300     If txtDoB = "" Then Exit Function

39310     pAge = CalcpAge(txtDoB)

39320     If pAge < 18 Then Exit Function

          ' 186 x (Creat / 88.4)-1.154 x (Age)-0.203 x (0.742 if female)
          'EGFR = 175 * ((Val(Creat) * 0.0113) ^ -1.154) * (Val(pAge) ^ -0.203)

          'eGFR = 175 x (Creat / 88.4)-1.154 x (Age)-0.203 x (0.742 if female)     (new formula given on 20171116)

39330     EGFR = 175 * ((Val(Creat) / 88.4) ^ -1.154) * (Val(pAge) ^ -0.203)


39340     If Left$(lblSex, 1) = "F" Then
39350         EGFR = EGFR * 0.742
39360     End If
39370     If EGFR > 90 Then
39380         CalculateEGFR = ">90"
39390     Else
39400         CalculateEGFR = EGFR
39410     End If

39420     Exit Function

CalculateEGFR_Error:

          Dim strES As String
          Dim intEL As Integer

39430     intEL = Erl
39440     strES = Err.Description
39450     LogError "frmEditAll", "CalculateEGFR", intEL, strES

End Function


Private Sub CheckAssGlucose(ByVal CurrentBRs As BIEResults)

          Dim tb As Recordset
          Dim sql As String
          Dim BR As BIEResult

39460     On Error GoTo CheckAssGlucose_Error

39470     If CurrentBRs.Count = 1 Then
39480         Set BR = CurrentBRs.Item(CodeForGlucose)
39490         If Not BR Is Nothing Then
                  '        CurrentBRs(1).Code = CodeForGlucose Then
                  'check prev or next for general
39500             sql = "Select distinct D.SampleID " & _
                      "from Demographics as D " & _
                      "where D.sampleid in " & _
                      "  (  select SampleID from BioResults where " & _
                      "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                      "     and Code <> '" & CodeForGlucose & "'  ) " & _
                      "and D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                      "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
39510             Set tb = New Recordset
39520             RecOpenServer 0, tb, sql
39530             If Not tb.EOF Then
39540                 lblAss = "Associated Results " & tb!SampleID
39550                 lblAss.Visible = True
39560             End If
39570         Else
39580             sql = "Select distinct D.SampleID " & _
                      "from Demographics as D " & _
                      "where D.sampleid in " & _
                      "  (  select SampleID from BioResults where " & _
                      "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                      "     and Code = '" & CodeForGlucose & "'  ) " & _
                      "and D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                      "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
39590             Set tb = New Recordset
39600             RecOpenServer 0, tb, sql
39610             If Not tb.EOF Then
39620                 lblAss = "Associated Glucose " & tb!SampleID
39630                 lblAss.Visible = True
39640             End If
39650         End If
39660     Else
39670         sql = "Select distinct D.SampleID " & _
                  "from Demographics as D " & _
                  "where D.sampleid in " & _
                  "  (  select SampleID from BioResults where " & _
                  "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                  "     and Code = '" & CodeForGlucose & "'  ) " & _
                  "and D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                  "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
39680         Set tb = New Recordset
39690         RecOpenServer 0, tb, sql
39700         If Not tb.EOF Then
39710             lblAss = "Associated Glucose " & tb!SampleID
39720             lblAss.Visible = True
39730         End If
39740     End If

39750     Exit Sub

CheckAssGlucose_Error:

          Dim strES As String
          Dim intEL As Integer

39760     intEL = Erl
39770     strES = Err.Description
39780     LogError "frmEditAll", "CheckAssGlucose", intEL, strES, sql

End Sub

Private Sub CheckAssTDM(ByVal CurrentBRs As BIEResults)

          Dim tb As Recordset
          Dim tbR As Recordset
          Dim sql As String
          Dim CurrentValue As String
          Dim AssValue As String
          Dim n As Integer
          Dim CodeForGent As String
          Dim CodeForTobra As String
          Dim CodeForVanco As String
          Dim BR As BIEResult
          'Gentamicin and Tobramicin

39790     On Error GoTo CheckAssTDM_Error

39800     lblAss.Caption = ""

39810     CodeForGent = GetOptionSetting("BioCodeForGentamicin", "")
39820     If CodeForGent = "" Then Exit Sub
39830     CodeForTobra = GetOptionSetting("BioCodeForTobramicin", "")
39840     If CodeForTobra = "" Then Exit Sub
39850     CodeForVanco = GetOptionSetting("BioCodeForVancomycin", "2869")
39860     If CodeForVanco = "" Then Exit Sub

39870     If CurrentBRs.Count < 4 Then
39880         For n = 1 To CurrentBRs.Count
39890             Set BR = CurrentBRs.Item(CodeForVanco)
39900             If Not BR Is Nothing Then
39910                 CurrentValue = BR.Result
39920                 sql = "Select distinct D.SampleID " & _
                          "from Demographics as D " & _
                          "where D.SampleID IN " & _
                          "  (  select SampleID from BioResults where " & _
                          "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                          "     and Code = '" & CodeForVanco & "'  ) " & _
                          "and D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                          "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
39930                 Set tb = New Recordset
39940                 RecOpenServer 0, tb, sql
39950                 If Not tb.EOF Then
39960                     sql = "Select Result from BioResults where " & _
                              "SampleID = '" & tb!SampleID & "' " & _
                              "and Code = '" & CodeForVanco & "'"
39970                     Set tbR = New Recordset
39980                     RecOpenServer 0, tbR, sql
39990                     If Not tbR.EOF Then
40000                         AssValue = tbR!Result & ""
40010                         If Val(AssValue) < Val(CurrentValue) Or InStr(AssValue, "<") <> 0 Then
40020                             lblAss = lblAss & "Vancomycin Trough " & Format$(AssValue, "0.0") & vbCrLf
40030                         Else
40040                             lblAss = lblAss & "Vancomycin Peak " & Format$(AssValue, "0.0") & vbCrLf
40050                         End If
40060                         lblAss.Visible = True
40070                     End If
40080                 End If
40090             Else
40100                 Set BR = CurrentBRs.Item(CodeForGent)
40110                 If Not BR Is Nothing Then
40120                     CurrentValue = BR.Result
40130                     sql = "Select distinct D.SampleID " & _
                              "from Demographics as D " & _
                              "where D.sampleid in " & _
                              "  (  select SampleID from BioResults where " & _
                              "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                              "     and Code = '" & CodeForGent & "'  ) " & _
                              "and D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                              "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
40140                     Set tb = New Recordset
40150                     RecOpenServer 0, tb, sql
40160                     If Not tb.EOF Then
40170                         sql = "Select Result from BioResults where " & _
                                  "SampleID = '" & tb!SampleID & "' " & _
                                  "and Code = '" & CodeForGent & "'"
40180                         Set tbR = New Recordset
40190                         RecOpenServer 0, tbR, sql
40200                         If Not tbR.EOF Then
40210                             AssValue = tbR!Result & ""
40220                             If Val(AssValue) < Val(CurrentValue) Or InStr(AssValue, "<") <> 0 Then
40230                                 lblAss = lblAss & "Gentamicin Trough " & Format$(AssValue, "0.0") & vbCrLf
40240                             Else
40250                                 lblAss = lblAss & "Gentamicin Peak " & Format$(AssValue, "0.0") & vbCrLf
40260                             End If
40270                             lblAss.Visible = True
40280                         End If
40290                     End If
40300                 Else
40310                     Set BR = CurrentBRs.Item(CodeForTobra)
40320                     If Not BR Is Nothing Then
40330                         CurrentValue = BR.Result
40340                         sql = "Select distinct D.SampleID " & _
                                  "from Demographics as D " & _
                                  "where D.sampleid in " & _
                                  "  (  select SampleID from BioResults where " & _
                                  "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                                  "     and Code = '" & CodeForTobra & "'  ) " & _
                                  "and D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
                                  "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
40350                         Set tb = New Recordset
40360                         RecOpenServer 0, tb, sql
40370                         If Not tb.EOF Then
40380                             sql = "Select Result from BioResults where " & _
                                      "SampleID = '" & tb!SampleID & "' " & _
                                      "and Code = '" & CodeForTobra & "'"
40390                             Set tbR = New Recordset
40400                             RecOpenServer 0, tbR, sql
40410                             If Not tbR.EOF Then
40420                                 AssValue = tbR!Result & ""
40430                                 If Val(AssValue) < Val(CurrentValue) Then
40440                                     lblAss = lblAss & "Tobramicin Trough " & Format$(AssValue, "0.0") & vbCrLf
40450                                 Else
40460                                     lblAss = lblAss & "Tobramicin Peak " & Format$(AssValue, "0.0") & vbCrLf
40470                                 End If
40480                                 lblAss.Visible = True
40490                             End If
40500                         End If
40510                     End If
40520                 End If
40530             End If
40540         Next
40550     End If

40560     Exit Sub

CheckAssTDM_Error:

          Dim strES As String
          Dim intEL As Integer

40570     intEL = Erl
40580     strES = Err.Description
40590     LogError "frmEditAll", "CheckAssTDM", intEL, strES, sql

End Sub

Private Sub CheckBioNormalStatus(ByRef BR As BIEResult)

    '      Dim Sex As String
    '      Dim DoB As Boolean
    '      Dim NewStatus As String
    '      Dim sql As String
    '
    '      'D DoB
    '      'B Male only
    '      'G Female only
    '      'M Male+DoB
    '      'F Female+DoB
    '      'N None
    '
    '10    On Error GoTo CheckBioNormalStatus_Error
    '
    '20    Sex = UCase$(Left$(txtSex, 1))
    '30    If Sex <> "M" And Sex <> "F" Then
    '40      Sex = ""
    '50    End If
    '
    '60    DoB = False
    '70    If IsDate(txtDoB) Then
    '80      DoB = True
    '90    End If
    '
    '100   If Sex = "M" And DoB Then
    '110     NewStatus = "M"
    '120   ElseIf Sex = "F" And DoB Then
    '130     NewStatus = "F"
    '140   ElseIf Sex = "M" And Not DoB Then
    '150     NewStatus = "B"
    '160   ElseIf Sex = "F" And Not DoB Then
    '170     NewStatus = "G"
    '180   ElseIf DoB Then
    '190     NewStatus = "D"
    '200   Else
    '210     NewStatus = "N"
    '220   End If
    '
    '230   If NewStatus <> BR.NormalUsed Then
    '240     BR.NormalLow = BR.Low
    '250     BR.NormalHigh = BR.High
    '260     BR.NormalUsed = NewStatus
    '270     sql = "UPDATE BioResults SET NormalUsed = '" & NewStatus & "', "
    '280     If NewStatus = "N" Then
    '290       sql = sql & "NormalLow = 0, NormalHigh = 9999 "
    '300     Else
    '310       sql = sql & "NormalLow = '" & BR.Low & "', NormalHigh = '" & BR.High & "' "
    '320     End If
    '330     sql = sql & "WHERE SampleID = '" & txtSampleID & "' " & _
    '              "AND Code = '" & BR.Code & "'"
    '340     Cnxn(0).Execute Sql
    '350   End If
    '
    '360   Exit Sub
    '
    'CheckBioNormalStatus_Error:
    '
    '      Dim strES As String
    '      Dim intEL As Integer
    '
    '370   intEL = Erl
    '380   strES = Err.Description
    '390   LogError "frmEditAll", "CheckBioNormalStatus", intEL, strES, sql
    '

End Sub

Private Sub CheckCalcPSA(ByVal BRs As BIEResults)

          Dim BR As BIEResult
          Dim FPS As Single
          Dim FPSTime As String
          Dim FPSDate As String
          Dim PSA As Single
          Dim Ratio As Single
          Dim Code As String

40600     On Error GoTo CheckCalcPSA_Error

40610     If BRs Is Nothing Then Exit Sub

40620     FPS = 0
40630     PSA = 0
40640     Ratio = 0

40650     For Each BR In BRs
40660         Code = UCase$(Trim$(BR.Code))
40670         If Code = "FPS" Then
40680             FPS = Val(BR.Result)
40690             FPSDate = BR.Rundate
40700             FPSTime = BR.RunTime
40710         ElseIf Code = "PSA" Then
40720             PSA = Val(BR.Result)
40730         ElseIf Code = "FPR" Then
40740             Ratio = Val(BR.Result)
40750         End If
40760     Next

40770     If (FPS * PSA) <> 0 And Ratio = 0 Then
40780         Ratio = FPS / PSA
40790         Set BR = New BIEResult
40800         BR.SampleID = txtSampleID
40810         BR.Code = "FPR"
40820         BR.Rundate = FPSDate
40830         BR.RunTime = FPSTime
40840         BR.Result = Format$(Ratio, "#0.00")
40850         BR.Units = ""
40860         BR.Printed = 0
40870         BR.Valid = 0
40880         BRs.Add BR
40890         BRs.Save "Bio", BRs
40900     End If

40910     Exit Sub

CheckCalcPSA_Error:

          Dim strES As String
          Dim intEL As Integer

40920     intEL = Erl
40930     strES = Err.Description
40940     LogError "frmEditAll", "CheckCalcPSA", intEL, strES

End Sub
Private Sub CheckCorrCalcium(ByVal BRs As BIEResults)

          Dim BRResNew As New BIEResults
          Dim BR As BIEResult
          Dim AlbTime As String
          Dim AlbDate As String
          Dim AlbResult As Single
          Dim CCalResult As Single
          Dim Code As String
          Dim CalciumResult As Single
          Dim CodeForALB As String
          Dim CodeForCa As String
          Dim CodeForCCal As String

40950     On Error GoTo CheckCorrCalcium_Error
          '+++ Junaid
40960     Exit Sub
          '--- Junaid
40970     If BRs Is Nothing Then Exit Sub

          'if albumin <42  then corrcalc = (42-ALB) x 0.02 + Calcium.
          'if Albumin >42  then corrcalc = (Alb-42) x 0.02 - Calcium

40980     CodeForALB = GetOptionSetting("BioCodeForAlb", "1015")
40990     CodeForCa = GetOptionSetting("BioCodeForCa", "1066")
41000     CodeForCCal = GetOptionSetting("BioCodeForCCal", "2000")

41010     AlbResult = 0
41020     CCalResult = 0
41030     CalciumResult = 0

41040     For Each BR In BRs
41050         Code = UCase$(Trim$(BR.Code))
41060         If Code = CodeForALB Then
41070             AlbResult = Val(BR.Result)
41080             AlbDate = BR.Rundate
41090             AlbTime = BR.RunTime
41100         ElseIf Code = CodeForCa Then
41110             CalciumResult = Val(BR.Result)
41120         ElseIf Code = CodeForCCal Then
41130             CCalResult = Val(BR.Result)
41140         End If
41150     Next

          Dim BD As New BIEDefinition
41160     Set BD = bds.ItemByShortName("CCAL")
41170     If Not BD Is Nothing Then
41180         If (AlbResult * CalciumResult) <> 0 And CCalResult = 0 Then
                  '+++Junaid
41190             CCalResult = 0 'CalciumResult + ((42 - AlbResult) * 0.02)
                  '---Junaid
41200             Set BR = New BIEResult
41210             BR.SampleID = txtSampleID
41220             BR.Code = CodeForCCal
41230             BR.Rundate = AlbDate
41240             BR.RunTime = AlbTime
                  '300           BR.Result = Format$(CCalResult, "#0.00")
41250             BR.Units = BD.Units
41260             BR.Printed = 0
41270             BR.Valid = 0
41280             BR.SampleType = "S"
41290             BR.ShortName = "CCAL"
41300             BR.Analyser = "Calc"
41310             BR.Category = "Human"
41320             BR.LongName = "Corrected Calcium"
41330             BR.Printformat = BD.DP
41340             BR.Low = BD.MaleLow
41350             BR.High = BD.MaleHigh
41360             BR.PlausibleHigh = BD.PlausibleHigh
41370             BR.PlausibleLow = BD.PlausibleLow
41380             BR.PrintRefRange = BD.PrintRefRange
41390             BR.FlagLow = BD.FlagMaleLow
41400             BR.FlagHigh = BD.FlagMaleHigh
41410             BRs.Add BR
41420             BRResNew.Add BR
41430             BRResNew.Save "Bio", BRResNew
41440         ElseIf (AlbResult * CalciumResult) <> 0 Then
                  '+++Junaid
41450             CCalResult = 0 'CalciumResult + ((42 - AlbResult) * 0.02)
                  '---Junaid
41460             Set BR = New BIEResult
41470             BR.SampleID = txtSampleID
41480             BR.Code = CodeForCCal
41490             BR.Rundate = AlbDate
41500             BR.RunTime = AlbTime
41510             BR.Result = Format$(CCalResult, "#0.00")
41520             BR.Units = BD.Units
41530             BR.Printed = 0
41540             BR.Valid = 0
41550             BR.SampleType = "S"
41560             BR.ShortName = "CCAL"
41570             BR.Analyser = "Calc"
41580             BR.Category = "Human"
41590             BR.LongName = "Corrected Calcium"
41600             BR.Printformat = BD.DP
41610             BR.Low = BD.MaleLow
41620             BR.High = BD.MaleHigh
41630             BR.PlausibleHigh = BD.PlausibleHigh
41640             BR.PlausibleLow = BD.PlausibleLow
41650             BR.PrintRefRange = BD.PrintRefRange
41660             BR.FlagLow = BD.FlagMaleLow
41670             BR.FlagHigh = BD.FlagMaleHigh
41680             BR.Update "Bio"

41690             BRs(CodeForCCal).Result = CCalResult

41700         End If
41710     End If

41720     Exit Sub

CheckCorrCalcium_Error:

          Dim strES As String
          Dim intEL As Integer

41730     intEL = Erl
41740     strES = Err.Description
41750     LogError "frmEditAll", "CheckCorrCalcium", intEL, strES

End Sub


Private Sub CheckuAlb24H(ByVal BRs As BIEResults)

          Dim BRResNew As New BIEResults
          Dim BR As BIEResult
          Dim uAlbTime As String
          Dim uAlbDate As String
          Dim uAlb As Single
          Dim Code As String
          Dim TUV As Single
          Dim CodeForTUV As String
          Dim CodeForuALB As String
          Dim CodeFor24HuAlb As String
          Dim Result As Single
          Dim BD As BIEDefinition

41760     On Error GoTo CheckuAlb24H_Error
          '+++ Junaid
41770     Exit Sub
          '--- Junaid
41780     If BRs Is Nothing Then Exit Sub

41790     CodeForuALB = GetOptionSetting("BioCodeForuAlb", "2839")
41800     CodeFor24HuAlb = GetOptionSetting("BioCodeFor24HuAlb", "24HuAlb")
41810     CodeForTUV = GetOptionSetting("BioCodeForTUV", "TUV")

41820     uAlb = 0
41830     TUV = 0
41840     Result = 0

41850     For Each BR In BRs
41860         Code = UCase$(Trim$(BR.Code))
41870         If Code = CodeForuALB Then
41880             uAlb = Val(BR.Result)
41890             uAlbDate = BR.Rundate
41900             uAlbTime = BR.RunTime
41910         ElseIf Code = CodeForTUV Then
41920             TUV = Val(BR.Result)
41930         ElseIf Code = CodeFor24HuAlb Then
41940             Result = Val(BR.Result)
41950         End If
41960     Next

41970     bds.Load
41980     Set BD = bds.ItemByCode(CodeFor24HuAlb)
41990     If Not BD Is Nothing Then
42000         If (TUV * uAlb) <> 0 And Result = 0 Then
42010             Result = uAlb * TUV / 1000
42020             Set BR = New BIEResult
42030             BR.SampleID = txtSampleID
42040             BR.Code = CodeFor24HuAlb
42050             BR.Rundate = uAlbDate
42060             BR.RunTime = uAlbTime
42070             BR.Result = Result
42080             BR.Units = BD.Units
42090             BR.Printed = 0
42100             BR.Valid = 0
42110             BR.SampleType = "U"
42120             BR.ShortName = BD.ShortName
42130             BR.Analyser = "Calc"
42140             BR.Category = "Human"
42150             BR.LongName = BD.LongName
42160             BR.Printformat = BD.DP
42170             BR.Low = BD.MaleLow
42180             BR.High = BD.MaleHigh
42190             BR.PlausibleHigh = BD.PlausibleHigh
42200             BR.PlausibleLow = BD.PlausibleLow
42210             BR.PrintRefRange = BD.PrintRefRange
42220             BR.FlagLow = BD.FlagMaleLow
42230             BR.FlagHigh = BD.FlagMaleHigh
42240             BRs.Add BR
42250             BRResNew.Add BR
42260             BRResNew.Save "Bio", BRResNew
42270         End If
42280     End If

42290     Exit Sub

CheckuAlb24H_Error:

          Dim strES As String
          Dim intEL As Integer

42300     intEL = Erl
42310     strES = Err.Description
42320     LogError "frmEditAll", "CheckuAlb24H", intEL, strES

End Sub

Private Sub CheckCholHDL(ByVal BRs As BIEResults)

          Dim BR As BIEResult
          Dim Chol As Single
          Dim HDL As Single
          Dim CholTime As String
          Dim CholDate As String
          Dim Ratio As Single
          Dim Code As String
          Dim BRResNew As New BIEResults
          Dim CodeForCholHDLRatio As String
          Dim CodeForHDL As String
          
42330     On Error GoTo ErrorHandler
42340     If BRs Is Nothing Then Exit Sub

42350     Chol = 0
42360     HDL = 0
42370     Ratio = 0

42380     CodeForCholHDLRatio = GetOptionSetting("BIOCODEFORCHOLHDLRATIO", "")
42390     If CodeForCholHDLRatio = "" Then Exit Sub
42400     CodeForHDL = GetOptionSetting("BIOCODEFORHDL", "")
42410     If CodeForHDL = "" Then Exit Sub

42420     For Each BR In BRs
42430         Code = UCase$(Trim$(BR.Code))
42440         If Code = CodeForChol Then
42450             Chol = Val(BR.Result)
42460             CholDate = BR.Rundate
42470             CholTime = BR.RunTime
42480         ElseIf Code = CodeForHDL Then
42490             HDL = Val(BR.Result)
42500         ElseIf Code = CodeForCholHDLRatio Then
42510             Ratio = Val(BR.Result)
42520         End If
42530     Next

42540     If (Chol * HDL) <> 0 And Ratio = 0 Then
42550         Ratio = Chol / HDL
42560         Set BR = New BIEResult
42570         BR.SampleID = txtSampleID
42580         BR.Code = CodeForCholHDLRatio
42590         BR.ShortName = "C/H R"
42600         BR.Rundate = CholDate
42610         BR.RunTime = CholTime
42620         BR.Result = Format$(Ratio, "#0.00")
42630         BR.SampleType = "S"
42640         BR.Units = "Ratio"
42650         BR.Valid = 0
42660         BR.Printed = 0
42670         BR.Authorised = 0
42680         BR.Printformat = 1

42690         BRs.Add BR
42700         BRResNew.Add BR
42710         BRResNew.Save "Bio", BRResNew
42720     End If
42730     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub CheckIfMustPhone(ByVal Discipline As String, ByVal Parameter As String, ByVal Result As String)

          Dim tb As Recordset
          Dim sql As String
          Dim TempResult As String

42740     On Error GoTo CheckIfMustPhone_Error


42750     TempResult = Result
42760     Do Until InStr(TempResult, ">") = 0
42770         TempResult = Replace(TempResult, ">", "")
42780     Loop
42790     Do Until InStr(TempResult, "<") = 0
42800         TempResult = Replace(TempResult, "<", "")
42810     Loop
42820     Result = Val(TempResult)

42830     sql = "SELECT COUNT(*) Tot FROM PhoneAlertLevel WHERE " & _
              "Discipline = '" & Discipline & "' " & _
              "AND Parameter = '" & Parameter & "' " & _
              "AND ( LessThan > " & Val(Result) & " " & _
              "   OR GreaterThan < " & Val(Result) & " ) "
42840     Set tb = New Recordset
42850     RecOpenServer 0, tb, sql
42860     If tb!Tot = 0 Or Val(Result) = 0 Then
42870         sql = "DELETE FROM PhoneAlert WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Discipline = '" & Discipline & "' " & _
                  "AND Parameter = '" & Parameter & "'"
42880         Cnxn(0).Execute sql
42890         Exit Sub
42900     End If

42910     sql = "IF EXISTS (SELECT * FROM PhoneLog " & _
              "           WHERE SampleID = '" & txtSampleID & "' " & _
              "           AND Discipline LIKE '%" & Left$(Discipline, 1) & "%') " & _
              "  DELETE FROM PhoneAlert WHERE " & _
              "  SampleID = '" & txtSampleID & "' " & _
              "  AND Discipline = '" & Discipline & "' " & _
              "ELSE " & _
              "  IF NOT EXISTS (SELECT SampleID FROM PhoneAlert WHERE " & _
              "                 Discipline = '" & Discipline & "' " & _
              "                 AND Parameter = '" & Parameter & "' " & _
              "                 AND SampleID = '" & txtSampleID & "') " & _
              "    INSERT INTO PhoneAlert " & _
              "    (SampleID, Discipline, Parameter) VALUES " & _
              "    ('" & txtSampleID & "', " & _
              "     '" & Discipline & " ', " & _
              "     '" & Parameter & "')"
42920     Cnxn(0).Execute sql

42930     Exit Sub

CheckIfMustPhone_Error:

          Dim strES As String
          Dim intEL As Integer

42940     intEL = Erl
42950     strES = Err.Description
42960     LogError "frmEditAll", "CheckIfMustPhone", intEL, strES, sql


End Sub

Private Sub CheckIfWardClinicianOK()

          Dim tb As Recordset
          Dim sql As String

42970     On Error GoTo CheckIfWardClinicianOK_Error

42980     If Trim$(cmbWard) <> "" Then
42990         sql = "SELECT Text FROM Wards WHERE " & _
                  "Text = '" & AddTicks(cmbWard) & "'"
43000         Set tb = New Recordset
43010         RecOpenServer 0, tb, sql
43020         If tb.EOF Then
43030             cmbWard = "GP"
43040         End If
43050     End If

          

43060     If Trim$(cmbClinician) <> "" Then
43070         sql = "SELECT Text FROM Clinicians WHERE " & _
                  "Text = '" & AddTicks(cmbClinician) & "'"
43080         Set tb = New Recordset
43090         RecOpenServer 0, tb, sql
43100         If tb.EOF Then
43110             cmbClinician = ""
43120         End If
43130     End If

43140     Exit Sub

CheckIfWardClinicianOK_Error:

          Dim strES As String
          Dim intEL As Integer

43150     intEL = Erl
43160     strES = Err.Description
43170     LogError "frmEditAll", "CheckIfWardClinicianOK", intEL, strES, sql

End Sub

Private Sub CheckRunSampleDates()

          Dim S As String
43180     On Error GoTo ErrorHandler
43190     If DateDiff("d", dtSampleDate, dtRunDate) <> 0 Then
43200         S = "Sample Date " & Format$(dtSampleDate, "dd/mm/yyyy") & " " & _
                  "Run Date " & Format$(dtRunDate, "dd/mm/yyyy")
43210         lblDateConflict.Caption = S
43220         lblDateConflict.Visible = True
43230     End If
43240     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub


Private Sub CheckVLDL(ByRef BRs As BIEResults)

          Dim BR As BIEResult
          Dim HDL As Single
          Dim VLDL As Single
          Dim LDL As Single
          Dim Chol As Single
          Dim Trig As Single
          Dim HDLTime As String
          Dim HDLDate As String
          Dim Code As String
          Dim BRResNew As New BIEResults
          Dim Low As Single
          Dim High As Single
          Dim tb As Recordset
          Dim sql As String
          Dim CodeForHDL As String
          Dim CodeForLDL As String
          Dim CodeForVLDL As String

43250     On Error GoTo CheckVLDL_Error

43260     If BRs Is Nothing Then Exit Sub

43270     CodeForHDL = GetOptionSetting("BIOCODEFORHDL", "")
43280     If CodeForHDL = "" Then Exit Sub
43290     CodeForLDL = GetOptionSetting("BIOCODEFORLDL", "")
43300     If CodeForLDL = "" Then Exit Sub
43310     CodeForVLDL = GetOptionSetting("BIOCODEFORVLDL", "")
43320     If CodeForVLDL = "" Then Exit Sub

43330     HDL = 0
43340     VLDL = 0    'vldl normal range 0.3-0.7 male and female
43350     LDL = 0
43360     Chol = 0
43370     Trig = 0

43380     For Each BR In BRs
43390         Code = UCase$(Trim$(BR.Code))
43400         If Code = CodeForHDL Then
43410             HDL = Val(BR.Result)
43420             HDLDate = BR.Rundate
43430             HDLTime = BR.RunTime
43440         ElseIf Code = CodeForVLDL Then
43450             VLDL = Val(BR.Result)
43460         ElseIf Code = CodeForLDL Then
43470             LDL = Val(BR.Result)
43480         ElseIf Code = CodeForChol Then
43490             Chol = Val(BR.Result)
43500         ElseIf Code = CodeForTrig Then
43510             Trig = Val(BR.Result)
43520         End If
43530     Next

43540     If Chol > 0 And HDL >= 0.1 And Trig <= 4 And Trig > 0 And (LDL = 0 Or VLDL = 0) Then

43550         VLDL = Trig / 2.18
43560         LDL = (Chol - HDL) - VLDL

43570         sql = "SELECT Code, " & _
                  "       COALESCE(FemaleLow, '0') AS Low, " & _
                  "       COALESCE(FemaleHigh, '9999') AS High, " & _
                  "       COALESCE(FlagFemaleLow, '0') AS FlagLow, " & _
                  "       COALESCE(FlagFemaleHigh, '9999') AS FlagHigh, " & _
                  "       COALESCE(PlausibleLow, '0') AS PlausibleLow, " & _
                  "       COALESCE(PlausibleHigh, '9999') AS PlausibleHigh " & _
                  "FROM BioTestDefinitions WHERE " & _
                  "Code = '" & CodeForLDL & "' " & _
                  "OR Code = '" & CodeForVLDL & "' "
43580         Set tb = New Recordset
43590         RecOpenServer 0, tb, sql
43600         Do While Not tb.EOF
43610             Low = Val(tb!Low)
43620             High = Val(tb!High)
43630             Set BR = New BIEResult
43640             BR.SampleID = txtSampleID
43650             If tb!Code = CodeForLDL Then
43660                 BR.Code = CodeForLDL
43670                 BR.ShortName = "LDL"
43680                 BR.LongName = "LDL"
43690                 BR.Result = Format$(LDL, "#0.00")
43700             Else
43710                 BR.Code = CodeForVLDL
43720                 BR.ShortName = "VLDL"
43730                 BR.LongName = "VLDL"
43740                 BR.Result = Format$(VLDL, "#0.00")
43750             End If
43760             BR.Rundate = HDLDate
43770             BR.RunTime = HDLTime
43780             BR.SampleType = "S"
43790             BR.Units = "mmol/l"
43800             BR.Valid = 0
43810             BR.Printed = 0
43820             BR.Authorised = 0
43830             BR.Printformat = 2
43840             BR.Low = Low
43850             BR.High = High
43860             BR.FlagLow = tb!FlagLow
43870             BR.FlagHigh = tb!FlagHigh
43880             BR.PlausibleLow = tb!PlausibleLow
43890             BR.PlausibleHigh = tb!PlausibleHigh
43900             BRs.Add BR
43910             BRResNew.Add BR

43920             tb.MoveNext
43930         Loop

43940         BRResNew.Save "Bio", BRResNew

43950     End If

43960     Exit Sub

CheckVLDL_Error:

          Dim strES As String
          Dim intEL As Integer

43970     intEL = Erl
43980     strES = Err.Description
43990     LogError "frmEditAll", "CheckVLDL", intEL, strES, sql

End Sub
Private Sub CheckDepartments()
44000     On Error GoTo ErrorHandler
44010     If frmOptDeptHaem Then
44020         If AreResultsPresent("Haem", txtSampleID) Then
44030             SSTab1.TabCaption(1) = "<<Haematology>>"
44040         End If
44050     End If

44060     If frmOptDeptBio Then
44070         If AreResultsPresent("Bio", txtSampleID) Then
44080             SSTab1.TabCaption(2) = "<<Biochemistry>>"
44090         End If
44100     End If

44110     If frmOptDeptCoag Then
44120         If AreResultsPresent("Coag", txtSampleID) Then
44130             SSTab1.TabCaption(3) = "<<Coagulation>>"
44140         End If
44150     End If

44160     If frmOptDeptExt Then
44170         If AreResultsPresent("Ext", txtSampleID) Then
44180             SSTab1.TabCaption(6) = "<<External>>"
44190         End If
44200     End If
44210     Exit Sub
ErrorHandler:
          '        MsgBox Err.Description
End Sub

Private Sub CheckIfPhoned()
44220     On Error GoTo ErrorHandler
          Dim PhLog As PhoneLog

44230     PhLog = CheckPhoneLog(txtSampleID)
44240     If PhLog.SampleID <> 0 Then
44250         cmdPhone.BackColor = vbYellow
44260         cmdPhone.Caption = "Results Phoned"
44270         cmdPhone.ToolTipText = "Results Phoned"
44280     Else
44290         cmdPhone.BackColor = &H8000000F
44300         cmdPhone.Caption = "Phone Results"
44310         cmdPhone.ToolTipText = "Phone Results"
44320     End If
44330     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub
Private Sub CheckIfFaxed()

          Dim tb As Recordset
          Dim sql As String

44340     On Error GoTo CheckIfFaxed_Error

44350     sql = "Select * from FaxLog where " & _
              " Cast(SampleID as varchar (100)) = '" & Val(txtSampleID) & "'"
44360     Set tb = Cnxn(0).Execute(sql)
44370     If tb.EOF Then
44380         cmdFAX.BackColor = &H8000000F
44390         cmdFAX.Caption = "&Fax Results"
44400         cmdFAX.ToolTipText = "Fax Results"
44410     Else
44420         cmdFAX.BackColor = vbYellow
44430         cmdFAX.Caption = "Results Faxed"
44440         cmdFAX.ToolTipText = "Right Click to view Fax Log"
44450     End If

44460     Exit Sub

CheckIfFaxed_Error:

          Dim strES As String
          Dim intEL As Integer

44470     intEL = Erl
44480     strES = Err.Description
44490     LogError "frmEditAll", "CheckIfFaxed", intEL, strES, sql

End Sub

Private Function CheckReagentLotNumber(ByVal Analyte As String, _
          ByVal SampleID As Long) _
          As Boolean

          Dim tb As Recordset
          Dim sql As String

44500     On Error GoTo CheckReagentLotNumber_Error

44510     CheckReagentLotNumber = False

44520     sql = "Select * from ReagentLotNumbers where " & _
              "Analyte = '" & Analyte & "' " & _
              "and SampleID = " & SampleID
44530     Set tb = New Recordset
44540     RecOpenServer 0, tb, sql

44550     If Not tb.EOF Then
44560         If Trim$(tb!LotNumber & "") <> "" Then
44570             CheckReagentLotNumber = True
44580         End If
44590     End If

44600     Exit Function

CheckReagentLotNumber_Error:

          Dim strES As String
          Dim intEL As Integer

44610     intEL = Erl
44620     strES = Err.Description
44630     LogError "frmEditAll", "CheckReagentLotNumber", intEL, strES, sql

End Function



Private Sub ClearCoagulation()
44640     On Error GoTo ErrorHandler
44650     With gCoag
44660         .Rows = 2
44670         .AddItem ""
44680         .RemoveItem 1
44690     End With
44700     cParameter = ""
44710     tResult = ""
44720     txtCoagComment = ""
44730     tWarfarin = ""
44740     bViewCoagRepeat.Visible = False

44750     With grdPrev
44760         .Rows = 2
44770         .AddItem ""
44780         .RemoveItem 1
44790     End With
44800     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Function egfrInclude(ByVal Ward As String, _
          ByVal Clinician As String, _
          ByVal GP As String) _
          As Boolean

          Dim tb As Recordset
          Dim sql As String

44810     On Error GoTo egfrInclude_Error

44820     egfrInclude = False

44830     If UCase$(Ward) = "GP" Then
44840         sql = "SELECT COUNT(*) Tot FROM IncludeEGFR WHERE " & _
                  "(SourceType = 'Ward' " & _
                  "AND SourceName = 'GP' " & _
                  "AND Include = 1 ) " & _
                  "OR " & _
                  "(SourceType = 'GP' " & _
                  "AND SourceName = '" & AddTicks(GP) & "' " & _
                  "AND Include = 1 ) "
44850     Else
44860         sql = "SELECT COUNT(*) Tot FROM IncludeEGFR WHERE " & _
                  "(SourceType = 'Ward' " & _
                  "AND SourceName = '" & AddTicks(Ward) & "' " & _
                  "AND Include = 1 ) " & _
                  "OR " & _
                  "(SourceType = 'Clinician' " & _
                  "AND SourceName = '" & AddTicks(Clinician) & "' " & _
                  "AND Include = 1 ) "
44870     End If
44880     Set tb = New Recordset
44890     RecOpenServer 0, tb, sql
44900     If tb!Tot > 0 Then
44910         egfrInclude = True
44920     End If

44930     Exit Function

egfrInclude_Error:

          Dim strES As String
          Dim intEL As Integer

44940     intEL = Erl
44950     strES = Err.Description
44960     LogError "frmEditAll", "egfrInclude", intEL, strES, sql

End Function

Private Sub EnableCopyFrom()

          Dim sql As String
          Dim tb As Recordset
          Dim PrevSID As Long

10        On Error GoTo EnableCopyFrom_Error
          'Uncomment
20        cmdCopyFromPrevious.Visible = False

30        If sysOptAllowCopyDemographics(0) = False Then
40            Exit Sub
50        End If

60        If Trim$(txtSurName) <> "" Or Trim$(txtForeName) <> "" Or Trim$(txtDoB) <> "" Then
70            Exit Sub
80        End If

90        PrevSID = Val(txtSampleID) - 1

100       sql = "Select PatName from Demographics where " & _
              "SampleID = " & PrevSID & " " & _
              "and PatName <> '' " & _
              "and PatName is not null " & _
              "and DoB is not null"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           cmdCopyFromPrevious.Caption = "Copy All Details from Sample # " & _
                  Format$(PrevSID) & _
                  " Name " & tb!PatName
150           'cmdCopyFromPrevious.Visible = True
160       End If

170       Exit Sub

EnableCopyFrom_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditAll", "EnableCopyFrom", intEL, strES, sql

End Sub

Private Sub FillcParameter()

          Dim tb As Recordset
          Dim sql As String

45160     On Error GoTo FillcParameter_Error

45170     cParameter.Clear

45180     sql = "SELECT DISTINCT TestName, MIN(PrintPriority) AS X FROM CoagTestDefinitions " & _
              "GROUP BY TestName " & _
              "ORDER BY X"
45190     Set tb = New Recordset
45200     RecOpenServer 0, tb, sql
45210     Do While Not tb.EOF
45220         cParameter.AddItem tb!TestName & ""
45230         tb.MoveNext
45240     Loop

45250     Exit Sub

FillcParameter_Error:

          Dim strES As String
          Dim intEL As Integer

45260     intEL = Erl
45270     strES = Err.Description
45280     LogError "frmEditAll", "FillcParameter", intEL, strES, sql

End Sub


Private Function GetHaemInfo(ByVal Analyte As String, _
          ByVal Sex As String, _
          ByVal DoB As String) _
          As udtHaem

          Dim sql As String
          Dim tb As Recordset
          Dim DaysOld As Long
          Dim SexSQL As String
          Dim RetVal As udtHaem

45290     On Error GoTo GetHaemInfo_Error

45300     Select Case Left$(UCase$(Sex), 1)
              Case "M"
45310             SexSQL = "COALESCE(MaleLow, '0') as Low, COALESCE(MaleHigh, '9999') as High"
45320         Case "F"
45330             SexSQL = "COALESCE(FemaleLow, '0') as Low, COALESCE(FemaleHigh, '9999') as High"
45340         Case Else
45350             SexSQL = "COALESCE(FemaleLow, '0') as Low, COALESCE(MaleHigh, '9999') as High"
45360     End Select

45370     sql = "Select top 1 " & _
              "COALESCE(PlausibleLow, '0') AS PlausibleLow, " & _
              "COALESCE(PlausibleHigh, '9999') AS PlausibleHigh, " & _
              SexSQL & ", DoDelta, DeltaValue ,DeltaDaysBackLimit " & _
              "from HaemTestDefinitions where Analytename = '" & Analyte & "' "

45380     If IsDate(DoB) Then

45390         DaysOld = Abs(DateDiff("d", Now, DoB))

45400         sql = sql & "AND AgeFromDays <= '" & DaysOld & "' " & _
                  "AND AgeToDays >= '" & DaysOld & "' " & _
                  "ORDER BY AgeFromDays DESC, AgeToDays ASC"
45410     Else
45420         sql = sql & "AND AgeFromDays = '0' " & _
                  "AND AgeToDays >= '43830'"
45430     End If

45440     Set tb = New Recordset
45450     RecOpenClient 0, tb, sql
45460     If Not tb.EOF Then
45470         With RetVal
45480             .DeltaValue = tb!DeltaValue
45490             .DoDelta = tb!DoDelta
45500             .High = tb!High
45510             .Low = tb!Low
45520             .PlausibleHigh = tb!PlausibleHigh
45530             .PlausibleLow = tb!PlausibleLow
45540             .DeltaDaysBackLimit = Val(tb!DeltaDaysBackLimit & "")
45550         End With
45560     Else
45570         RetVal.High = 9999
45580         RetVal.PlausibleHigh = 9999
45590         RetVal.Low = 0
45600         RetVal.PlausibleLow = 0
45610         RetVal.DeltaValue = 9999
45620         RetVal.DoDelta = False
45630     End If

45640     GetHaemInfo = RetVal

45650     Exit Function

GetHaemInfo_Error:

          Dim strES As String
          Dim intEL As Integer

45660     intEL = Erl
45670     strES = Err.Description
45680     LogError "frmEditAll", "GetHaemInfo", intEL, strES, sql

End Function



'---------------------------------------------------------------------------------------
' Procedure : GetCoagInfo
' Author    : XPMUser
' Date      : 13/Aug/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetCoagInfo(ByVal Analyte As String, _
          ByVal Sex As String, _
          ByVal DoB As String) _
          As udtHaem

          Dim sql As String
          Dim tb As Recordset
          Dim DaysOld As Long
          Dim SexSQL As String
          Dim RetVal As udtHaem



45690     On Error GoTo GetCoagInfo_Error


45700     Select Case Left$(UCase$(Sex), 1)
              Case "M"
45710             SexSQL = "COALESCE(MaleLow, 0) as Low, COALESCE(MaleHigh, 9999) as High"
45720         Case "F"
45730             SexSQL = "COALESCE(FemaleLow, 0) as Low, COALESCE(FemaleHigh, 9999) as High"
45740         Case Else
45750             SexSQL = "COALESCE(FemaleLow, 0) as Low, COALESCE(MaleHigh, 9999) as High"
45760     End Select

45770     sql = "Select top 1 " & _
              "COALESCE(PlausibleLow, 0) AS PlausibleLow, " & _
              "COALESCE(PlausibleHigh, 9999) AS PlausibleHigh, " & _
              SexSQL & ", DoDelta, deltalimit ,DeltaDaysBackLimit " & _
              "from CoagTestDefinitions where TestName = '" & Analyte & "' "

45780     If IsDate(DoB) Then

45790         DaysOld = Abs(DateDiff("d", Now, DoB))

45800         sql = sql & "AND AgeFromDays <= '" & DaysOld & "' " & _
                  "AND AgeToDays >= '" & DaysOld & "' " & _
                  "ORDER BY AgeFromDays DESC, AgeToDays ASC"
45810     Else
45820         sql = sql & "AND AgeFromDays = '0' " & _
                  "AND AgeToDays >= '43830'"
45830     End If

45840     Set tb = New Recordset
45850     RecOpenClient 0, tb, sql
45860     If Not tb.EOF Then
45870         With RetVal
45880             .DeltaValue = tb!DeltaLimit
45890             .DoDelta = tb!DoDelta
45900             .High = tb!High
45910             .Low = tb!Low
45920             .PlausibleHigh = tb!PlausibleHigh
45930             .PlausibleLow = tb!PlausibleLow
45940             .DeltaDaysBackLimit = Val(tb!DeltaDaysBackLimit & "")
45950         End With
45960     Else
45970         RetVal.High = 9999
45980         RetVal.PlausibleHigh = 9999
45990         RetVal.Low = 0
46000         RetVal.PlausibleLow = 0
46010         RetVal.DeltaValue = 9999
46020         RetVal.DoDelta = False
46030     End If

46040     GetCoagInfo = RetVal




46050     Exit Function


GetCoagInfo_Error:

          Dim strES As String
          Dim intEL As Integer

46060     intEL = Erl
46070     strES = Err.Description
46080     LogError "frmEditAll", "GetCoagInfo", intEL, strES, sql

End Function
Public Sub LoadBiochemistry()

          Dim DeltaSn As Recordset
          Dim Deltatb As Recordset
          Dim tb As Recordset
          Dim sql As String
          Dim S As String
          Dim Value As Single
          Dim OldValue As Single
          Dim valu As String
          Dim PreviousDate As String
          Dim Res As String
          Dim n As Integer
          Dim DeltaLimit As Single
          Dim SampleType As String
          Dim BRs As New BIEResults
          Dim BRres As BIEResults
          Dim BR As BIEResult
          Dim Fasting As Boolean
          Dim Fx As Fasting
          Dim Flag As String
          Dim Rundate As String
          Dim l As Single
          Dim H As Single
          Dim PSA As Boolean
          Dim LessOrGreater As Single
          Dim FoundK As Boolean
          Dim DeltaDate As String
          Dim DoCheckEGFR As Boolean
          Dim Highlight As Integer
          Dim CodeForUCreat As String
          Dim MaskFlag As String

          Dim DoCheckACR As Boolean
          Dim DoCheckUrine24hr As Boolean
          Dim DoCheckUrineCreatinineRatio As Boolean
          Dim DoCheckCreatinineClearance As Boolean

46090     On Error GoTo LoadBiochemistry_Error

46100     SetViewReports "Biochemistry", txtSampleID

46110     CodeForUCreat = GetOptionSetting("BioCodeForUCreat", "")

46120     lblGreaterThan.Visible = False

          'Abubaker Siddique +++ 09-10-2023 autocomments stop
          'Abubaker Siddique +++ 29-11-2023 reverted

46130     txtAutoComment(2) = ""

46140     txtAutoComment(2) = CheckAutoComments(txtSampleID, 2)
          'Abubaker Siddique ---

46150     DoCheckEGFR = GetOptionSetting("CheckEGFR", 0) = 1
46160     DoCheckACR = GetOptionSetting("CheckACR", 0) = 1    ' Masood 11-02-2016

46170     DoCheckUrine24hr = GetOptionSetting("CheckUrine24hr", 0) = 1   ' Masood 11-02-2016
46180     DoCheckUrineCreatinineRatio = GetOptionSetting("CheckUrineCreatinineRatio", 0) = 1  ' Masood 03-03-2016

46190     DoCheckCreatinineClearance = GetOptionSetting("CheckCreatinineClearance", 0) = 1    ' Masood 03-03-2016

46200     txtSampleID = Format$(Val(txtSampleID))

46210     lblDateConflict.Visible = False

46220     lblLipaemic = ""
46230     lblIcteric = ""
46240     lblHaemolysed = ""
46250     chkOld.Value = 0

46260     Set LiIcHas = New LIHs

46270     Fasting = lRandom = "Fasting Sample"

46280     lblDelta(2) = ""
46290     bViewBioRepeat.Visible = False

46300     SSTab1.TabCaption(2) = "Biochemistry"
46310     PreviousBio = False
46320     Set BRres = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE, , , Trim$(txtSex), Trim$(txtDoB))

46330     lblAss.Visible = False
46340     If Not BRres Is Nothing Then
46350         If sysOptDoAssGlucose(0) Then
46360             CheckAssGlucose BRres
46370         End If
46380         If Not lblAss.Visible Then    'lblAss is only visible if AssGlucose succeeds
46390             CheckAssTDM BRres
46400         End If
        

46410         CheckCalcPSA BRres
              '340       CheckCorrCalcium BRres
              '350       CheckuAlb24H BRres

46420         If DoCheckEGFR Then CheckEGFR BRres
46430         If sysOptCheckCholHDLRatio(0) Then CheckCholHDL BRres

46440         If DoCheckACR Then CheckACR BRres

46450         If DoCheckUrine24hr Then
46460             Call CheckUrine24hr(BRres, "BioCodeForUrineSodium24hr", "Na24", "BioCodeForUrinarySodium", "1133")
46470             Call CheckUrine24hr(BRres, "BioCodeForUrinePotassium24hr", "K24", "BioCodeForUrinaryPotassium", "1134")
                  '+++ Junaid
                  '420           Call CheckUrine24hr(BRres, "BioCodeForUrineCalcium24hr", "qq", "BioCodeForUrinaryCalcium", "1096")
                  '--- Junaid
46480             Call CheckUrine24hr(BRres, "BioCodeForUrineUrate24hr", "Ur24", "BioCodeForUrinaryUrate", "1041")

46490             Call CheckUrine24hr(BRres, "BioCodeForUrinePhosphate24hr", "", "BioCodeForUrinaryPhosphate", "")
46500             Call CheckUrine24hr(BRres, "BioCodeForUrineProtein24hr", "", "BioCodeForUrinaryProtein", "1044")
46510             Call CheckUrine24hr(BRres, "BioCodeForUrineChloride24hr", "cl24", "BioCodeForUrinaryChloride", "")
46520             Call CheckUrine24hr(BRres, "BioCodeForUrineCreatinine24hr", "", "BIOCODEFORUCREAT", "1068")
46530         End If

46540         If DoCheckUrineCreatinineRatio Then
46550             Call CheckUrineCreatinineRatio(BRres)
46560         End If

46570         If DoCheckCreatinineClearance Then
46580             Call CheckCreatinineClearance(BRres)
46590         End If

46600     End If

46610     PSA = False

46620     gBio.Visible = False
46630     gBio.Rows = 2
46640     gBio.AddItem ""
46650     gBio.RemoveItem 1

46660     FoundK = False

46670     lblSplit(0).BackColor = vbButtonFace

46680     For n = 1 To 6
46690         If n <> lblSplitView.Tag Then
46700             lblSplit(n).BackColor = vbButtonFace
46710             lblSplit(n).ForeColor = vbBlack
46720         End If
46730     Next

46740     If Not BRres Is Nothing Then
46750         SSTab1.TabCaption(2) = ">>Biochemistry<<"
46760         For Each BR In BRres
46770             If (UCase(BR.Analyser) <> "BIOMNIS") And (UCase(BR.Analyser) <> "MATER") And (UCase(BR.Analyser) <> "NVRL") Then
46780                 If BR.ShortName <> "L" And BR.ShortName <> "I" And BR.ShortName <> "H" Then
46790                     If BR.Code = CodeForUCreat Then
46800                         AdjustUCreat BR
46810                     End If
                          '+++ Junaid 04-04-2024
46820                     If BR.Analyser = "Calc" And BR.Code = "2000" Then
46830                         sql = "Delete from BioResults Where Analyser = 'Calc' And Code = '2000' And SampleID = '" & BR.SampleID & "'"
46840                         Cnxn(0).Execute sql
46850                     End If
                          '--- Junaid
46860                     If UCase$(BR.LongName) = "TROPONIN" Then
46870                         If Val(BR.Result) < 0.01 Then
46880                             BR.Result = "<0.01"
46890                             sql = "UPDATE BioResults SET Result = '<0.01' WHERE SampleID = '" & txtSampleID & "' AND Code = '" & BR.Code & "'"
46900                             Cnxn(0).Execute sql
46910                         End If
46920                     End If

46930                     If UCase$(BR.ShortName) = "PSA" And IsNumeric(BR.Result) Then
46940                         If Val(BR.Result) < 0.1 Then
46950                             BR.Result = "<0.1"
46960                             sql = "UPDATE BioResults SET Result = '<0.1' WHERE SampleID = '" & txtSampleID & "' AND Code = '" & BR.Code & "'"
46970                             Cnxn(0).Execute sql
46980                         End If
46990                     End If

47000                     Highlight = ProcessThisSplit(BR.ShortName)
47010                     lblSplit(Highlight).BackColor = vbGreen
47020                     lblSplit(0).BackColor = vbYellow
47030                     If Highlight = lblSplitView.Tag Then
47040                         lblSplit(Highlight).BackColor = vbRed
47050                         lblSplit(0).BackColor = vbYellow
47060                     End If
47070                     If lblSplitView.Tag = 0 Or Highlight = lblSplitView.Tag Then
47080                         If MaskResult(BR.Code) Then
47090                             BR.Result = "XXXXX"
47100                         End If

47110                         CheckIfMustPhone "Biochemistry", BR.ShortName, BR.Result

47120                         If BR.ShortName = "PSA" Then
47130                             PSA = True
47140                         End If

47150                         Rundate = Format$(BR.Rundate, "dd/mmm/yyyy")
47160                         Flag = ""
47170                         SampleType = BR.SampleType
47180                         If Len(SampleType) = 0 Then SampleType = "S"
47190                         S = BR.ShortName & vbTab

47200                         If IsNumeric(BR.Result) Then
47210                             Value = Val(BR.Result)
47220                             Select Case BR.Printformat
                                      Case 0: valu = Format$(Value, "0")
47230                                 Case 1: valu = Format$(Value, "0.0")
47240                                 Case 2: valu = Format$(Value, "0.00")
47250                                 Case 3: valu = Format$(Value, "0.000")
47260                                 Case Else: valu = Format$(Value, "0.000")
47270                             End Select
47280                         Else
47290                             valu = BR.Result
47300                             LessOrGreater = 0
47310                             LessOrGreater = Val(Replace(valu, "<", ""))
47320                             If Val(LessOrGreater) = 0 Then
47330                                 LessOrGreater = Val(Replace(valu, ">", ""))
47340                             End If
47350                             If LessOrGreater <> 0 Then
47360                                 BR.Result = LessOrGreater
47370                             End If
47380                         End If
47390                         S = S & valu & vbTab

47400                         If BR.PrintRefRange Then
47410                             If Fasting And (BR.Code = CodeForGlucose Or BR.Code = CodeForChol Or BR.Code = CodeForTrig) Then    '------------
47420                                 Set Fx = Nothing
47430                                 If BR.Code = CodeForGlucose Then
47440                                     Set Fx = colFastings("GLU")
47450                                 ElseIf BR.Code = CodeForChol Then
47460                                     Set Fx = colFastings("CHO")
47470                                 ElseIf BR.Code = CodeForTrig Then
47480                                     Set Fx = colFastings("TRI")
47490                                 End If
47500                                 If Not Fx Is Nothing Then
47510                                     If Fx.FastingLow = 0 And (Fx.FastingHigh = 999 Or Fx.FastingHigh = 0 Or Fx.FastingHigh = 9999) Then
47520                                         S = S & ""
47530                                     ElseIf Fx.FastingLow = 0 Then
47540                                         S = S & "< " & Format$(Fx.FastingHigh)
47550                                     ElseIf Fx.FastingHigh = 999 Or Fx.FastingHigh = 9999 Then
47560                                         S = S & "> " & Format$(Fx.FastingLow)
47570                                     Else
47580                                         S = S & Format$(Fx.FastingLow) & "-" & Format$(Fx.FastingHigh)
47590                                     End If
47600                                 End If

47610                             Else    '-----------------

47620                                 l = Val(BR.Low)
47630                                 H = Val(BR.High)
47640                                 If l = 0 And (H = 999 Or H = 0 Or H = 9999) Then
47650                                     S = S & ""
47660                                 ElseIf l = 0 Then
47670                                     S = S & "< " & Format$(H)
47680                                 ElseIf H = 999 Or H = 9999 Then
47690                                     S = S & "> " & Format$(l)
47700                                 Else
47710                                     S = S & Format$(l) & "-" & Format$(H)
47720                                 End If
47730                             End If
47740                         End If
47750                         S = S & vbTab & BR.Units & vbTab

47760                         If IsNumeric(BR.Result) Then

47770                             MaskFlag = MaskInhibit(BR, BRres)
47780                             If MaskFlag <> "" Then
47790                                 Flag = "X"
47800                                 S = BR.ShortName & vbTab & "XXXXX" & vbTab & vbTab & vbTab & MaskFlag
47810                             Else

47820                                 Value = BR.Result
47830                                 If Value > Val(BR.PlausibleHigh) Then
47840                                     Flag = "X"
47850                                     S = S & "X"
47860                                     BR.Result = "XXXXX"
47870                                     BR.Flags = "X"
47880                                     BR.Update "Bio"
47890                                     S = BR.ShortName & vbTab & "XXXXX" & vbTab & vbTab & vbTab & "X"
47900                                 ElseIf Value < Val(BR.PlausibleLow) Then
47910                                     Flag = "X"
47920                                     S = S & "X"
47930                                     BR.Result = "XXXXX"
47940                                     BR.Flags = "X"
47950                                     BR.Update "Bio"
47960                                     S = BR.ShortName & vbTab & "XXXXX" & vbTab & vbTab & vbTab & "X"
47970                                 ElseIf BR.Code = CodeForGlucose Or _
                                          BR.Code = CodeForChol Or _
                                          BR.Code = CodeForTrig Then
47980                                     If Fasting Then
47990                                         Set Fx = Nothing
48000                                         If BR.Code = CodeForGlucose Then
48010                                             Set Fx = colFastings("GLU")
48020                                         ElseIf BR.Code = CodeForChol Then
48030                                             Set Fx = colFastings("CHO")
48040                                         ElseIf BR.Code = CodeForTrig Then
48050                                             Set Fx = colFastings("TRI")
48060                                         End If
48070                                         If Not Fx Is Nothing Then
48080                                             If Value > Fx.FastingHigh Then
48090                                                 Flag = "H"
48100                                                 S = S & "H"
48110                                             ElseIf Value < Fx.FastingLow Then
48120                                                 Flag = "L"
48130                                                 S = S & "L"
48140                                             End If
48150                                         End If
48160                                     Else
48170                                         If Value < Val(BR.FlagLow) Then
48180                                             Flag = "L"
48190                                             S = S & "L"
48200                                         ElseIf Value > Val(BR.FlagHigh) Then
48210                                             Flag = "H"
48220                                             S = S & "H"
48230                                         End If
48240                                     End If
48250                                 Else
48260                                     If Value <= Val(BR.FlagLow) Then
48270                                         Flag = "L"
48280                                         S = S & "L"
48290                                     Else
48300                                         If Value >= Val(BR.FlagHigh) Then
48310                                             Flag = "H"
48320                                             S = S & "H"
48330                                         End If
48340                                     End If
48350                                 End If
48360                             End If
                                  '1900          ElseIf BR.Result = "XXXXX" Then
                                  '                Flag = "X"
48370                         End If
                              'Au Flags
48380                         If Trim$(BR.Flags) <> "" And Flag = "" Then
48390                             If Right(S, 1) <> vbTab Then
48400                                 S = Left(S, Len(S) - 1)
48410                             End If
48420                             S = S & BR.Flags
48430                         ElseIf Trim$(BR.Flags) <> "" And Flag <> "" Then
48440                             If Right(S, 1) <> vbTab Then
48450                                 S = Left(S, Len(S) - 1)
48460                             End If
48470                             S = S & Flag
48480                         End If
48490                         S = S & vbTab & BR.Analyser & vbTab & _
                                  IIf(BR.Valid, "V", " ") & _
                                  IIf(BR.Printed, "P", " ")
48500                         S = S & vbTab & BR.Comment & vbTab & vbTab & BR.Code
48510                         gBio.AddItem S

48520                         If BR.Valid Then
48530                             BioValBy = BR.Operator
48540                         End If

48550                         If Flag <> "" Then
48560                             gBio.row = gBio.Rows - 1
48570                             gBio.Col = 1
48580                             Select Case Flag
                                      Case "H":
48590                                     gBio.CellBackColor = vbRed
48600                                     gBio.CellForeColor = vbYellow
48610                                 Case "L":
48620                                     gBio.CellBackColor = vbBlue
48630                                     gBio.CellForeColor = vbYellow
48640                                 Case "X":
48650                                     gBio.CellBackColor = vbBlack
48660                                     gBio.CellForeColor = vbWhite
48670                             End Select
48680                         End If

48690                         If BR.DoDelta Then

                                  Dim DateQry As String
48700                             If IsDate(tSampleTime) Then
48710                                 DateQry = " AND (D.SampleDate < '" & Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm") & "')"
48720                             Else
48730                                 DateQry = " AND (D.SampleDate < '" & Format$((dtSampleDate), "dd/MMM/yyyy") & "')"
48740                             End If
48750                             sql = " SELECT     TOP (1)  R.result, R.Code,D.SampleDate,R.sampleid " & _
                                      " FROM         demographics AS D INNER JOIN BioResults AS R ON D.SampleID = R.sampleid  " & _
                                      " WHERE     (D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "') AND (D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "') " & _
                                      " AND (left(D.SampleDate,10) >= '" & Format$(dtSampleDate - BR.DeltaDaysBackLimit, "dd/mm/yyyy") & "')  " & _
                                      DateQry & _
                                      " AND (D.sampleid <> '" & txtSampleID & "')" & _
                                      " and R.code = '" & BR.Code & "'" & _
                                      "  ORDER BY D.SampleDate DESC,D.sampleid ASC "


48760                             Set DeltaSn = New Recordset
48770                             RecOpenClient 0, DeltaSn, sql
48780                             If Not DeltaSn.EOF Then
48790                                 OldValue = Val(DeltaSn!Result)
48800                                 If OldValue <> 0 Then
48810                                     DeltaLimit = BR.DeltaLimit
48820                                     If Abs(OldValue - Value) > DeltaLimit Then
48830                                         Res = Format$(DeltaSn!SampleDate, "dd/mm/yyyy") & " (" & DeltaSn!SampleID & ") " & _
                                                  BR.ShortName & " " & _
                                                  FormatNumber(OldValue, BR.Printformat) & vbCr

48840                                         lblDelta(2) = lblDelta(2) & Res
48850                                     End If
48860                                 End If
48870                                 PreviousBio = True
48880                             End If
48890                         End If
48900                         If UCase$(Left$(BR.LongName, 6)) = "POTASS" Then
48910                             FoundK = True
48920                         End If
48930                     Else
48940                         lblSplit(Highlight).BackColor = vbGreen
48950                         lblSplit(0).BackColor = vbYellow
48960                     End If

48970                 Else

48980                     Select Case BR.ShortName
                              Case "L": If lblLipaemic = "" Then lblLipaemic = TranslateLIH(BR)
48990                         Case "I": If lblIcteric = "" Then lblIcteric = TranslateLIH(BR)
49000                         Case "H": If lblHaemolysed = "" Then lblHaemolysed = TranslateLIH(BR)
49010                     End Select

49020                 End If
49030             End If
49040         Next
49050     End If

49060     SaveMasks

49070     With gBio
49080         For n = 1 To .Rows - 1
49090             If InStr(.TextMatrix(n, 4), "X") <> 0 Then
49100                 .row = n
49110                 .Col = 1
49120                 .CellForeColor = vbWhite
49130                 .CellBackColor = 1
49140             End If
49150         Next
49160         If .Rows > 2 Then
49170             .RemoveItem 1
49180         End If
49190         .Visible = True
49200         If .Rows > 22 Then
49210             lblGreaterThan.Caption = "Ensure all " & .Rows - 1 & " Results are reviewed"
49220             lblGreaterThan.Visible = True
49230         End If
49240     End With

49250     LoadOutstanding

49260     sql = "Select * from BioRepeats where " & _
              "SampleID = '" & Val(txtSampleID) & "'"
49270     Set tb = New Recordset
49280     RecOpenClient 0, tb, sql
49290     bViewBioRepeat.Visible = False
49300     If Not tb.EOF Then
49310         bViewBioRepeat.Visible = True
49320     End If

49330     If FoundK Then
49340         CheckRunSampleDates
49350     End If

49360     SetPrintInhibit "Bio"

49370     Exit Sub

LoadBiochemistry_Error:

          Dim strES As String
          Dim intEL As Integer

49380     intEL = Erl
49390     strES = Err.Description
49400     LogError "frmEditAll", "LoadBiochemistry", intEL, strES, sql

End Sub


'Public Sub LoadBiochemistry()
'
'      Dim DeltaSn As Recordset
'      Dim Deltatb As Recordset
'      Dim tb As Recordset
'      Dim sql As String
'      Dim S As String
'      Dim Value As Single
'      Dim OldValue As Single
'      Dim valu As String
'      Dim PreviousDate As String
'      Dim Res As String
'      Dim n As Integer
'      Dim DeltaLimit As Single
'      Dim SampleType As String
'      Dim BRs As New BIEResults
'      Dim BRres As BIEResults
'      Dim BR As BIEResult
'      Dim Fasting As Boolean
'      Dim Fx As Fasting
'      Dim Flag As String
'      Dim Rundate As String
'      Dim l As Single
'      Dim H As Single
'      Dim PSA As Boolean
'      Dim LessOrGreater As Single
'      Dim FoundK As Boolean
'      Dim DeltaDate As String
'      Dim DoCheckEGFR As Boolean
'      Dim Highlight As Integer
'      Dim CodeForUCreat As String
'      Dim MaskFlag As String
'
'      Dim DoCheckACR As Boolean
'      Dim DoCheckUrine24hr As Boolean
'      Dim DoCheckUrineCreatinineRatio As Boolean
'      Dim DoCheckCreatinineClearance As Boolean
'
'10    On Error GoTo LoadBiochemistry_Error
'
'20    SetViewReports "Biochemistry", txtSampleID
'
'30    CodeForUCreat = GetOptionSetting("BioCodeForUCreat", "")
'
'40    lblGreaterThan.Visible = False
'50    txtAutoComment(2) = ""
'60    txtAutoComment(2) = CheckAutoComments(txtSampleID, 2)
'
'70    DoCheckEGFR = GetOptionSetting("CheckEGFR", 0) = 1
'80    DoCheckACR = GetOptionSetting("CheckACR", 0) = 1    ' Masood 11-02-2016
'
'90    DoCheckUrine24hr = GetOptionSetting("CheckUrine24hr", 0) = 1   ' Masood 11-02-2016
'100   DoCheckUrineCreatinineRatio = GetOptionSetting("CheckUrineCreatinineRatio", 0) = 1  ' Masood 03-03-2016
'
'110   DoCheckCreatinineClearance = GetOptionSetting("CheckCreatinineClearance", 0) = 1    ' Masood 03-03-2016
'
'120   txtSampleID = Format$(Val(txtSampleID))
'
'130   lblDateConflict.Visible = False
'
'140   lblLipaemic = ""
'150   lblIcteric = ""
'160   lblHaemolysed = ""
'170   chkOld.Value = 0
'
'180   Set LiIcHas = New LIHs
'
'190   Fasting = lRandom = "Fasting Sample"
'
'200   lblDelta(2) = ""
'210   bViewBioRepeat.Visible = False
'
'220   SSTab1.TabCaption(2) = "Biochemistry"
'230   PreviousBio = False
'240   Set BRres = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE, , , Trim$(txtSex), Trim$(txtDoB))
'
'250   lblAss.Visible = False
'260   If Not BRres Is Nothing Then
'270       If sysOptDoAssGlucose(0) Then
'280           CheckAssGlucose BRres
'290       End If
'300       If Not lblAss.Visible Then    'lblAss is only visible if AssGlucose succeeds
'310           CheckAssTDM BRres
'320       End If
'
'330       CheckCalcPSA BRres
'340       CheckCorrCalcium BRres
'350       CheckuAlb24H BRres
'
'360       If DoCheckEGFR Then CheckEGFR BRres
'370       If sysOptCheckCholHDLRatio(0) Then CheckCholHDL BRres
'
'380       If DoCheckACR Then CheckACR BRres
'
'390       If DoCheckUrine24hr Then
'400           Call CheckUrine24hr(BRres, "BioCodeForUrineSodium24hr", "Na24", "BioCodeForUrinarySodium", "1133")
'410           Call CheckUrine24hr(BRres, "BioCodeForUrinePotassium24hr", "K24", "BioCodeForUrinaryPotassium", "1134")
'420           Call CheckUrine24hr(BRres, "BioCodeForUrineCalcium24hr", "qq", "BioCodeForUrinaryCalcium", "1096")
'430           Call CheckUrine24hr(BRres, "BioCodeForUrineUrate24hr", "Ur24", "BioCodeForUrinaryUrate", "1041")
'
'440           Call CheckUrine24hr(BRres, "BioCodeForUrinePhosphate24hr", "", "BioCodeForUrinaryPhosphate", "")
'450           Call CheckUrine24hr(BRres, "BioCodeForUrineProtein24hr", "", "BioCodeForUrinaryProtein", "1044")
'460           Call CheckUrine24hr(BRres, "BioCodeForUrineChloride24hr", "cl24", "BioCodeForUrinaryChloride", "")
'470           Call CheckUrine24hr(BRres, "BioCodeForUrineCreatinine24hr", "", "BIOCODEFORUCREAT", "1068")
'480       End If
'
'490       If DoCheckUrineCreatinineRatio Then
'500           Call CheckUrineCreatinineRatio(BRres)
'510       End If
'
'520       If DoCheckCreatinineClearance Then
'530           Call CheckCreatinineClearance(BRres)
'540       End If
'
'550   End If
'
'560   PSA = False
'
'570   gBio.Visible = False
'580   gBio.Rows = 2
'590   gBio.AddItem ""
'600   gBio.RemoveItem 1
'
'610   FoundK = False
'
'620   lblSplit(0).BackColor = vbButtonFace
'
'630   For n = 1 To 6
'640       If n <> lblSplitView.Tag Then
'650           lblSplit(n).BackColor = vbButtonFace
'660           lblSplit(n).ForeColor = vbBlack
'670       End If
'680   Next
'
'690   If Not BRres Is Nothing Then
'700       SSTab1.TabCaption(2) = ">>Biochemistry<<"
'710       For Each BR In BRres
'720           If (UCase(BR.Analyser) <> "BIOMNIS") And (UCase(BR.Analyser) <> "MATLAB") And (UCase(BR.Analyser) <> "NVRL") Then
'730               If BR.ShortName <> "L" And BR.ShortName <> "I" And BR.ShortName <> "H" Then
'740                   If BR.Code = CodeForUCreat Then
'750                       AdjustUCreat BR
'760                   End If
'
'770                   If UCase$(BR.LongName) = "TROPONIN" Then
'780                       If Val(BR.Result) < 0.01 Then
'790                           BR.Result = "<0.01"
'800                           sql = "UPDATE BioResults SET Result = '<0.01' WHERE SampleID = '" & txtSampleID & "' AND Code = '" & BR.Code & "'"
'810                           Cnxn(0).Execute Sql
'820                       End If
'830                   End If
'
'840                   If UCase$(BR.ShortName) = "PSA" And IsNumeric(BR.Result) Then
'850                       If Val(BR.Result) < 0.1 Then
'860                           BR.Result = "<0.1"
'870                           sql = "UPDATE BioResults SET Result = '<0.1' WHERE SampleID = '" & txtSampleID & "' AND Code = '" & BR.Code & "'"
'880                           Cnxn(0).Execute Sql
'890                       End If
'900                   End If
'
'910                   Highlight = ProcessThisSplit(BR.ShortName)
'920                   lblSplit(Highlight).BackColor = vbGreen
'930                   lblSplit(0).BackColor = vbYellow
'940                   If Highlight = lblSplitView.Tag Then
'950                       lblSplit(Highlight).BackColor = vbRed
'960                       lblSplit(0).BackColor = vbYellow
'970                   End If
'980                   If lblSplitView.Tag = 0 Or Highlight = lblSplitView.Tag Then
'990                       If MaskResult(BR.Code) Then
'1000                          BR.Result = "XXXXX"
'1010                      End If
'
'1020                      CheckIfMustPhone "Biochemistry", BR.ShortName, BR.Result
'
'1030                      If BR.ShortName = "PSA" Then
'1040                          PSA = True
'1050                      End If
'
'1060                      Rundate = Format$(BR.Rundate, "dd/mmm/yyyy")
'1070                      Flag = ""
'1080                      SampleType = BR.SampleType
'1090                      If Len(SampleType) = 0 Then SampleType = "S"
'1100                      S = BR.ShortName & vbTab
'
'1110                      If IsNumeric(BR.Result) Then
'1120                          Value = Val(BR.Result)
'1130                          Select Case BR.Printformat
'                              Case 0: valu = Format$(Value, "0")
'1140                          Case 1: valu = Format$(Value, "0.0")
'1150                          Case 2: valu = Format$(Value, "0.00")
'1160                          Case 3: valu = Format$(Value, "0.000")
'1170                          Case Else: valu = Format$(Value, "0.000")
'1180                          End Select
'1190                      Else
'1200                          valu = BR.Result
'1210                          LessOrGreater = 0
'1220                          LessOrGreater = Val(Replace(valu, "<", ""))
'1230                          If Val(LessOrGreater) = 0 Then
'1240                              LessOrGreater = Val(Replace(valu, ">", ""))
'1250                          End If
'1260                          If LessOrGreater <> 0 Then
'1270                              BR.Result = LessOrGreater
'1280                          End If
'1290                      End If
'1300                      S = S & valu & vbTab
'
'1310                      If BR.PrintRefRange Then
'1320                          If Fasting And (BR.Code = CodeForGlucose Or BR.Code = CodeForChol Or BR.Code = CodeForTrig) Then    '------------
'1330                              Set Fx = Nothing
'1340                              If BR.Code = CodeForGlucose Then
'1350                                  Set Fx = colFastings("GLU")
'1360                              ElseIf BR.Code = CodeForChol Then
'1370                                  Set Fx = colFastings("CHO")
'1380                              ElseIf BR.Code = CodeForTrig Then
'1390                                  Set Fx = colFastings("TRI")
'1400                              End If
'1410                              If Not Fx Is Nothing Then
'1420                                  If Fx.FastingLow = 0 And (Fx.FastingHigh = 999 Or Fx.FastingHigh = 0 Or Fx.FastingHigh = 9999) Then
'1430                                      S = S & ""
'1440                                  ElseIf Fx.FastingLow = 0 Then
'1450                                      S = S & "< " & Format$(Fx.FastingHigh)
'1460                                  ElseIf Fx.FastingHigh = 999 Or Fx.FastingHigh = 9999 Then
'1470                                      S = S & "> " & Format$(Fx.FastingLow)
'1480                                  Else
'1490                                      S = S & Format$(Fx.FastingLow) & "-" & Format$(Fx.FastingHigh)
'1500                                  End If
'1510                              End If
'
'1520                          Else    '-----------------
'
'1530                              l = Val(BR.Low)
'1540                              H = Val(BR.High)
'1550                              If l = 0 And (H = 999 Or H = 0 Or H = 9999) Then
'1560                                  S = S & ""
'1570                              ElseIf l = 0 Then
'1580                                  S = S & "< " & Format$(H)
'1590                              ElseIf H = 999 Or H = 9999 Then
'1600                                  S = S & "> " & Format$(l)
'1610                              Else
'1620                                  S = S & Format$(l) & "-" & Format$(H)
'1630                              End If
'1640                          End If
'1650                      End If
'1660                      S = S & vbTab & BR.Units & vbTab
'
'1670                      If IsNumeric(BR.Result) Then
'
'1680                          MaskFlag = MaskInhibit(BR, BRres)
'1690                          If MaskFlag <> "" Then
'1700                              Flag = "X"
'1710                              S = BR.ShortName & vbTab & "XXXXX" & vbTab & vbTab & vbTab & MaskFlag
'1720                          Else
'
'1730                              Value = BR.Result
'1740                              If Value > Val(BR.PlausibleHigh) Then
'1750                                  Flag = "X"
'1760                                  S = S & "X"
'1770                                  BR.Result = "XXXXX"
'1780                                  BR.Flags = "X"
'1790                                  BR.Update "Bio"
'1800                                  S = BR.ShortName & vbTab & "XXXXX" & vbTab & vbTab & vbTab & "X"
'1810                              ElseIf Value < Val(BR.PlausibleLow) Then
'1820                                  Flag = "X"
'1830                                  S = S & "X"
'1840                                  BR.Result = "XXXXX"
'1850                                  BR.Flags = "X"
'1860                                  BR.Update "Bio"
'1870                                  S = BR.ShortName & vbTab & "XXXXX" & vbTab & vbTab & vbTab & "X"
'1880                              ElseIf BR.Code = CodeForGlucose Or _
'                                         BR.Code = CodeForChol Or _
'                                         BR.Code = CodeForTrig Then
'1890                                  If Fasting Then
'1900                                      Set Fx = Nothing
'1910                                      If BR.Code = CodeForGlucose Then
'1920                                          Set Fx = colFastings("GLU")
'1930                                      ElseIf BR.Code = CodeForChol Then
'1940                                          Set Fx = colFastings("CHO")
'1950                                      ElseIf BR.Code = CodeForTrig Then
'1960                                          Set Fx = colFastings("TRI")
'1970                                      End If
'1980                                      If Not Fx Is Nothing Then
'1990                                          If Value > Fx.FastingHigh Then
'2000                                              Flag = "H"
'2010                                              S = S & "H"
'2020                                          ElseIf Value < Fx.FastingLow Then
'2030                                              Flag = "L"
'2040                                              S = S & "L"
'2050                                          End If
'2060                                      End If
'2070                                  Else
'2080                                      If Value < Val(BR.FlagLow) Then
'2090                                          Flag = "L"
'2100                                          S = S & "L"
'2110                                      ElseIf Value > Val(BR.FlagHigh) Then
'2120                                          Flag = "H"
'2130                                          S = S & "H"
'2140                                      End If
'2150                                  End If
'2160                              Else
'2170                                  If Value <= Val(BR.FlagLow) Then
'2180                                      Flag = "L"
'2190                                      S = S & "L"
'2200                                  Else
'2210                                      If Value >= Val(BR.FlagHigh) Then
'2220                                          Flag = "H"
'2230                                          S = S & "H"
'2240                                      End If
'2250                                  End If
'2260                              End If
'2270                          End If
'                              '1900          ElseIf BR.Result = "XXXXX" Then
'                              '                Flag = "X"
'2280                      End If
'                          'Au Flags
'2290                      If Trim$(BR.Flags) <> "" And Flag = "" Then
'2300                          If Right(S, 1) <> vbTab Then
'2310                              S = Left(S, Len(S) - 1)
'2320                          End If
'2330                          S = S & BR.Flags
'2340                      ElseIf Trim$(BR.Flags) <> "" And Flag <> "" Then
'2350                          If Right(S, 1) <> vbTab Then
'2360                              S = Left(S, Len(S) - 1)
'2370                          End If
'2380                          S = S & Flag
'2390                      End If
'2400                      S = S & vbTab & BR.Analyser & vbTab & _
'                              IIf(BR.Valid, "V", " ") & _
'                              IIf(BR.Printed, "P", " ")
'2410                      S = S & vbTab & BR.Comment & vbTab & vbTab & BR.Code
'2420                      gBio.AddItem S
'
'2430                      If BR.Valid Then
'2440                          BioValBy = BR.Operator
'2450                      End If
'
'2460                      If Flag <> "" Then
'2470                          gBio.row = gBio.Rows - 1
'2480                          gBio.Col = 1
'2490                          Select Case Flag
'                              Case "H":
'2500                              gBio.CellBackColor = vbRed
'2510                              gBio.CellForeColor = vbYellow
'2520                          Case "L":
'2530                              gBio.CellBackColor = vbBlue
'2540                              gBio.CellForeColor = vbYellow
'2550                          Case "X":
'2560                              gBio.CellBackColor = vbBlack
'2570                              gBio.CellForeColor = vbWhite
'2580                          End Select
'2590                      End If
'
'2600                      If BR.DoDelta Then
'
'                              Dim DateQry As String
'2610                          If IsDate(tSampleTime) Then
'2620                              DateQry = " AND (D.SampleDate < '" & Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm") & "')"
'2630                          Else
'2640                              DateQry = " AND (D.SampleDate < '" & Format$((dtSampleDate), "dd/MMM/yyyy") & "')"
'2650                          End If
'2660                          sql = " SELECT     TOP (1)  R.result, R.Code,D.SampleDate,R.sampleid " & _
'                                    " FROM         demographics AS D INNER JOIN BioResults AS R ON D.SampleID = R.sampleid  " & _
'                                    " WHERE     (D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "') AND (D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "') " & _
'                                    " AND (D.SampleDate >= '" & Format$(dtSampleDate - BR.DeltaDaysBackLimit, "dd/MMM/yyyy") & "')  " & _
'                                    DateQry & _
'                                    " AND (D.sampleid <> '" & txtSampleID & "')" & _
'                                    " and R.code = '" & BR.Code & "'" & _
'                                    "  ORDER BY D.SampleDate DESC,D.sampleid ASC "
'
'
'2670                          Set DeltaSn = New Recordset
'2680                          RecOpenClient 0, DeltaSn, sql
'2690                          If Not DeltaSn.EOF Then
'2700                              OldValue = Val(DeltaSn!Result)
'2710                              If OldValue <> 0 Then
'2720                                  DeltaLimit = BR.DeltaLimit
'2730                                  If Abs(OldValue - Value) > DeltaLimit Then
'2740                                      Res = Format$(DeltaSn!SampleDate, "dd/mm/yyyy") & " (" & DeltaSn!sampleid & ") " & _
'                                                BR.ShortName & " " & _
'                                                FormatNumber(OldValue, BR.Printformat) & vbCr
'
'2750                                      lblDelta(2) = lblDelta(2) & Res
'2760                                  End If
'2770                              End If
'2780                              PreviousBio = True
'2790                          End If
'2800                      End If
'2810                      If UCase$(Left$(BR.LongName, 6)) = "POTASS" Then
'2820                          FoundK = True
'2830                      End If
'2840                  Else
'2850                      lblSplit(Highlight).BackColor = vbGreen
'2860                      lblSplit(0).BackColor = vbYellow
'2870                  End If
'
'2880              Else
'
'2890                  Select Case BR.ShortName
'                      Case "L": If lblLipaemic = "" Then lblLipaemic = TranslateLIH(BR)
'2900                  Case "I": If lblIcteric = "" Then lblIcteric = TranslateLIH(BR)
'2910                  Case "H": If lblHaemolysed = "" Then lblHaemolysed = TranslateLIH(BR)
'2920                  End Select
'
'2930              End If
'2940          End If
'2950      Next
'2960  End If
'
'2970  SaveMasks
'
'2980  With gBio
'2990      For n = 1 To .Rows - 1
'3000          If InStr(.TextMatrix(n, 4), "X") <> 0 Then
'3010              .row = n
'3020              .Col = 1
'3030              .CellForeColor = vbWhite
'3040              .CellBackColor = 1
'3050          End If
'3060      Next
'3070      If .Rows > 2 Then
'3080          .RemoveItem 1
'3090      End If
'3100      .Visible = True
'3110      If .Rows > 22 Then
'3120          lblGreaterThan.Caption = "Ensure all " & .Rows - 1 & " Results are reviewed"
'3130          lblGreaterThan.Visible = True
'3140      End If
'3150  End With
'
'3160  LoadOutstanding
'
'3170  sql = "Select * from BioRepeats where " & _
'            "SampleID = '" & Val(txtSampleID) & "'"
'3180  Set tb = New Recordset
'3190  RecOpenClient 0, tb, sql
'3200  bViewBioRepeat.Visible = False
'3210  If Not tb.EOF Then
'3220      bViewBioRepeat.Visible = True
'3230  End If
'
'3240  If FoundK Then
'3250      CheckRunSampleDates
'3260  End If
'
'3270  SetPrintInhibit "Bio"
'
'3280  Exit Sub
'
'LoadBiochemistry_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'3290  intEL = Erl
'3300  strES = Err.Description
'3310  LogError "frmEditAll", "LoadBiochemistry", intEL, strES, sql
'
'End Sub
Private Sub MoveToNextRelevant(ByVal Index As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim strDept As String
          Dim strDirection As String
          Dim strSplitSelect As String
          Dim strArrow As String

49410     On Error GoTo MoveToNextRelevant_Error

49420     Select Case SSTab1.Tab
              Case 0:
49430             If Index = 0 Then
49440                 txtSampleID = Format$(Val(txtSampleID) - 1)
49450             Else
49460                 txtSampleID = Format$(Val(txtSampleID) + 1)
49470             End If
49480             Debug.Print "Movetonextrelevant 80"

49490             LoadAllDetails

49500             cmdSaveHoldDemographics.Enabled = False
49510             cmdSaveDemographics.Enabled = False
49520             cmdSaveHaem.Enabled = False
49530             cmdSaveBio.Enabled = False
49540             cmdSaveCoag.Enabled = False
49550             Exit Sub

49560         Case 1: strDept = "Haem"
49570         Case 2: strDept = "Bio"
49580         Case 3: strDept = "Coag"
49590         Case 6: lblResultOrRequest = "Results": strDept = "Ext"
49600     End Select

49610     strDirection = IIf(Index = 0, "Desc", "Asc")
49620     strArrow = IIf(Index = 0, "<", ">")

49630     Select Case lblResultOrRequest
              Case "Results", "SampleID":
49640             If SSTab1.Tab = 6 Then    'ext
49650                 sql = "Select top 1 SampleID from ExtResults where " & _
                          "SampleID " & strArrow & " " & txtSampleID & " " & _
                          "and (Result is Null or Result like '') " & _
                          "Order by SampleID " & strDirection
49660             Else
49670                 sql = "Select top 1 SampleID from " & strDept & "Results where " & _
                          "SampleID " & strArrow & " " & txtSampleID & " " & _
                          "Order by SampleID " & strDirection
49680             End If
49690         Case "Request":
49700             sql = "Select top 1 SampleID from " & strDept & "Requests where " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "Order by SampleID " & strDirection
49710         Case "Not Val":
49720             sql = "Select top 1 SampleID from " & strDept & "Results where " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and Valid = 0 " & _
                      "Order by SampleID " & strDirection
49730     End Select
          'MsgBox Sql
49740     If strDept = "Bio" And lblSplitView.Tag <> "0" Then
49750         strSplitSelect = LoadSplitList(Val(lblSplitView.Tag))
49760         If lblResultOrRequest = "Results" Then
49770             If strSplitSelect <> "" Then
49780                 sql = "Select top 1 SampleID from BioResults where " & _
                          "SampleID " & strArrow & " " & txtSampleID & " " & _
                          "and (" & strSplitSelect & ") " & _
                          "Order by SampleID " & strDirection
49790             End If
49800         Else
49810             If strSplitSelect <> "" Then
49820                 sql = "Select top 1 SampleID from BioRequests where " & _
                          "SampleID " & strArrow & " " & txtSampleID & " " & _
                          "and (" & strSplitSelect & ") " & _
                          "Order by SampleID " & strDirection
49830             End If
49840         End If
49850     End If

49860     Set tb = New Recordset
49870     RecOpenClient 0, tb, sql
49880     If Not tb.EOF Then
49890         txtSampleID = tb!SampleID & ""
49900     End If
49910     Debug.Print "Movetonextrelevant 500"

49920     LoadAllDetails

49930     cmdSaveHoldDemographics.Enabled = False
49940     cmdSaveDemographics.Enabled = False
49950     cmdSaveHaem.Enabled = False
49960     cmdSaveBio.Enabled = False
49970     cmdSaveCoag.Enabled = False

49980     Exit Sub

MoveToNextRelevant_Error:

          Dim strES As String
          Dim intEL As Integer

49990     intEL = Erl
50000     strES = Err.Description
50010     LogError "frmEditAll", "MoveToNextRelevant", intEL, strES, sql

End Sub

Private Function ProcessThisSplit(ByVal ShortName As String) As Integer
          'Returns Split number 1 to 6
          ''''or 0 if viewing all

          Dim RetVal As Integer
          Dim sql As String
          Dim tb As Recordset

50020     On Error GoTo ProcessThisSplit_Error

50030     RetVal = 0

50040     sql = "SELECT COALESCE(SplitList, 0) Split FROM BioTestDefinitions " & _
              "WHERE ShortName = '" & ShortName & "'"
50050     Set tb = New Recordset
50060     RecOpenServer 0, tb, sql
50070     If Not tb.EOF Then
50080         RetVal = tb!Split
50090     End If
50100     ProcessThisSplit = RetVal

          '
          '  sql = "SELECT COUNT(*) Tot FROM BioTestDefinitions " & _
          '        "WHERE SplitList = '" & lblSplitView.Tag & "' " & _
          '        "AND ShortName = '" & ShortName & "'"
          '  Set tb = New Recordset
          '  RecOpenServer 0, tb, sql
          '  RetVal = tb!Tot > 0
          'End If
          'ProcessThisSplit = RetVal

50110     Exit Function

ProcessThisSplit_Error:

          Dim strES As String
          Dim intEL As Integer

50120     intEL = Erl
50130     strES = Err.Description
50140     LogError "frmEditAll", "ProcessThisSplit", intEL, strES, sql

End Function


Private Sub SaveMasks()

          Dim Dept As String
          Dim intLIH As Integer
          Dim intH As Integer
          Dim intS As Integer
          Dim intL As Integer
          Dim intO As Integer
          Dim intG As Integer
          Dim intJ As Integer
          Dim sql As String

50150     On Error GoTo SaveMasks_Error

50160     intH = IIf(Val(lblHaemolysed) = 3, 1, 0)
50170     intS = IIf((Val(lblHaemolysed) = 1) Or (Val(lblHaemolysed) = 2), 1, 0)
50180     intL = IIf(Val(lblLipaemic) > 0, 1, 0)
50190     intO = IIf(chkOld.Value = 1, 1, 0)
50200     intG = IIf((Val(lblHaemolysed) > 3), 1, 0)
50210     intJ = IIf(Val(lblIcteric) > 0, 1, 0)
50220     intLIH = (Val(lblLipaemic) * 100) + (Val(lblIcteric) * 10) + (Val(lblHaemolysed))

50230     If (intLIH Or intO) = 0 Then
50240         sql = "DELETE FROM " & Dept & "Masks " & _
                  "WHERE SampleID = '" & txtSampleID & "'"
50250         Cnxn(0).Execute sql
50260     Else
50270         sql = "IF EXISTS (SELECT * FROM Masks " & _
                  "           WHERE SampleID = '" & txtSampleID & "') " & _
                  "  UPDATE Masks " & _
                  "  SET Rundate = getdate(), " & _
                  "  LIH = '" & intLIH & "', " & _
                  "  H = '" & intH & "', " & _
                  "  S = '" & intS & "', " & _
                  "  L = '" & intL & "', " & _
                  "  O = '" & intO & "', " & _
                  "  G = '" & intG & "', " & _
                  "  J = '" & intJ & "' " & _
                  "  WHERE SampleID = '" & txtSampleID & "' " & _
                  "ELSE " & _
                  "  INSERT INTO Masks " & _
                  "  (SampleID, RunDate, LIH, H, S, L, O, G, J) VALUES " & _
                  "  ('" & txtSampleID & "', " & _
                  "   getdate(), " & _
                  "  '" & intLIH & "', " & _
                  "  '" & intH & "', " & _
                  "  '" & intS & "', " & _
                  "  '" & intL & "', " & _
                  "  '" & intO & "', " & _
                  "  '" & intG & "', " & _
                  "  '" & intJ & "')"
50280         Cnxn(0).Execute sql
50290     End If

50300     Exit Sub

SaveMasks_Error:

          Dim strES As String
          Dim intEL As Integer

50310     intEL = Erl
50320     strES = Err.Description
50330     LogError "frmEditAll", "SaveMasks", intEL, strES, sql

End Sub

Private Sub SetPrintInhibit(ByVal Dept As String)

          Dim y As Integer

50340     On Error GoTo SetPrintInhibit_Error

50350     Select Case Dept

              Case "Bio"
50360             gBio.Col = 8
50370             If gBio.TextMatrix(1, 0) <> "" Then
50380                 For y = 1 To gBio.Rows - 1
50390                     gBio.row = y
50400                     If InStr(gBio.TextMatrix(y, 6), "P") Then
50410                         Set gBio.CellPicture = imgRedCross.Picture
50420                     Else
50430                         If gBio.TextMatrix(y, 5) <> "Manual" Then
50440                             Set gBio.CellPicture = imgGreenTick.Picture
50450                         Else
50460                             If InStr(gBio.TextMatrix(y, 6), "V") Then
50470                                 Set gBio.CellPicture = imgGreenTick.Picture
50480                             Else
50490                                 Set gBio.CellPicture = imgRedCross.Picture
50500                             End If
50510                         End If
50520                     End If
50530                 Next
50540             End If

50550         Case "Coa"
50560             gCoag.Col = 5
50570             If gCoag.TextMatrix(1, 0) <> "" Then
50580                 For y = 1 To gCoag.Rows - 1
50590                     If InStr(gCoag.TextMatrix(y, 4), "P") Then
50600                         gCoag.row = y
50610                         Set gCoag.CellPicture = imgRedCross.Picture
50620                     Else
50630                         gCoag.row = y
50640                         Set gCoag.CellPicture = imgGreenTick.Picture
50650                     End If
50660                 Next
50670             End If
50680     End Select

50690     Exit Sub

SetPrintInhibit_Error:

          Dim strES As String
          Dim intEL As Integer

50700     intEL = Erl
50710     strES = Err.Description
50720     LogError "frmEditAll", "SetPrintInhibit", intEL, strES

End Sub
Private Sub FillLists()

50730     On Error GoTo FillLists_Error

50740     FillWards cmbWard, HospName(0)
50750     FillClinicians cmbClinician, HospName(0)
50760     FillGPs cmbGP, HospName(0)

50770     FillGenericList cmbUnits, "UN"
50780     cmbUnits.ListIndex = -1

50790     FillGenericList cClDetails, "CD"

50800     FillGenericList cmbHospital, "HO"

50810     FillGenericList cmbDemogComment, "DE"

50820     FillGenericList cmbBioComment(0), "BI"

50830     FillGenericList cmbBioComment(2), "CO"

50840     FillGenericList cmbHaemComment, "HA"

50850     FillGenericList cmbFilmComment, "FI"

50860     FillGenericList cmbNewResult, "NewResult"

50870     Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

50880     intEL = Erl
50890     strES = Err.Description
50900     LogError "frmEditAll", "FillLists", intEL, strES

End Sub

Private Sub LoadComments()

          Dim OB As Observation
          Dim OBs As Observations

50910     On Error GoTo LoadComments_Error

50920     txtBioComment = ""
50930     txtHaemComment = ""
50940     cmbBioComment(0) = ""
50950     cmbBioComment(1) = ""
50960     txtDemographicComment = ""
50970     txtCoagComment = ""
50980     txtFilmComment = ""
50990     If Val(txtSampleID) = 0 Then Exit Sub

51000     Set OBs = New Observations
51010     Set OBs = OBs.Load(txtSampleID, "Biochemistry", "Haematology", "Demographic", "Coagulation", "Film")
51020     If Not OBs Is Nothing Then
51030         For Each OB In OBs
                  '+++ Junaid 24-07-2023
51040             Select Case UCase$(OB.Discipline)
                      Case "BIOCHEMISTRY": txtBioComment = Split_Comm(OB.Comment)
51050                 Case "HAEMATOLOGY": txtHaemComment = Split_Comm(OB.Comment)
51060                 Case "DEMOGRAPHIC": txtDemographicComment = Split_Comm(OB.Comment)
51070                 Case "COAGULATION": txtCoagComment = Split_Comm(OB.Comment)
51080                 Case "FILM": txtFilmComment = Split_Comm(OB.Comment)
51090             End Select

                  '140           Select Case UCase$(OB.Discipline)
                  '              Case "HAEMATOLOGY": txtHaemComment = Split_Comm(OB.Comment)
                  '160           Case "DEMOGRAPHIC": txtDemographicComment = Split_Comm(OB.Comment)
                  '170           Case "COAGULATION": txtCoagComment = Split_Comm(OB.Comment)
                  '180           Case "FILM": txtFilmComment = Split_Comm(OB.Comment)
                  '190           End Select
                  '--- Junaid
51100         Next
51110     End If

51120     Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

51130     intEL = Erl
51140     strES = Err.Description
51150     LogError "frmEditAll", "LoadComments", intEL, strES

End Sub

Private Function LoadLIH() As Boolean

          Dim sql As String
          Dim tb As Recordset
          Dim RetVal As Boolean
          Dim Lipaemic As Integer
          Dim Icteric As Integer
          Dim Haemolysed As Integer
          Dim LIHVal As Integer

51160     On Error GoTo LoadLIH_Error

51170     RetVal = False

51180     sql = "SELECT * FROM Masks WHERE " & _
              "SampleID = '" & Val(txtSampleID) & "' " & _
              "AND (COALESCE(LIH, 0) <> 0 OR COALESCE([O], 0) <> 0 )"
51190     Set tb = New Recordset
51200     RecOpenClient 0, tb, sql
51210     If Not tb.EOF Then
51220         RetVal = True
51230         LIHVal = tb!LIH
51240         If LIHVal > 99 Then
51250             Lipaemic = LIHVal \ 100
51260             If Lipaemic > 6 Then
51270                 Lipaemic = 0
51280             End If
51290             If Lipaemic > 0 And Lipaemic < 7 Then
51300                 lblLipaemic = Format$(Lipaemic) & "+"
51310             End If
51320             LIHVal = LIHVal Mod 100
51330         End If
51340         If LIHVal > 9 Then
51350             Icteric = LIHVal \ 10
51360             If Icteric > 6 Then
51370                 Icteric = 0
51380             End If
51390             If Icteric > 0 And Icteric < 7 Then
51400                 lblIcteric = Format$(Icteric) & "+"
51410             End If
51420             LIHVal = LIHVal Mod 10
51430         End If
51440         Haemolysed = LIHVal
51450         If Haemolysed > 6 Then
51460             Haemolysed = 0
51470         End If
51480         If Haemolysed > 0 And Haemolysed < 7 Then
51490             lblHaemolysed = Format$(Haemolysed) & "+"
51500         End If
51510         chkOld.Value = IIf(tb!o, 1, 0)
51520     End If

51530     LoadLIH = RetVal

51540     Exit Function

LoadLIH_Error:

          Dim strES As String
          Dim intEL As Integer

51550     intEL = Erl
51560     strES = Err.Description
51570     LogError "frmEditAll", "LoadLIH", intEL, strES, sql

End Function

Private Sub LoadPreviousCoag()

          Dim tb As Recordset
          Dim sql As String
          Dim CRs As CoagResults
          Dim CR As CoagResult
          Dim PrevDate As String
          Dim PrevID As String
          Dim S As String

51580     On Error GoTo LoadPreviousCoag_Error

51590     PreviousCoag = False
51600     If Trim$(txtChart) <> "" Then
              'If AddTicks(txtSurName & " " & txtForeName) & "' " <> "" Then

51610         sql = "select rundate, sampleid from demographics where " & _
                  "chart = '" & txtChart & "' and " & _
                  "sampleid < '" & txtSampleID & "' " & _
                  "order by rundate desc"

51620         Set tb = New Recordset
51630         RecOpenServer 0, tb, sql
51640         If Not tb.EOF Then
51650             PrevDate = Format$(tb!Rundate, "dd/mm/yy")
51660             PrevID = tb!SampleID

51670             Set CRs = New CoagResults
51680             Set CRs = CRs.Load(PrevID, gDONTCARE, gDONTCARE, "Results")

51690             If Not CRs Is Nothing Then
51700                 PreviousCoag = True
51710                 For Each CR In CRs
51720                     S = CR.TestName & vbTab & _
                              CR.Result
51730                     grdPrev.AddItem S
51740                 Next
51750                 lblPrevCoag = PrevDate & " Result for " & txtChart
51760             Else
51770                 lblPrevCoag = "No Previous Coag Details"
51780             End If
51790         Else
51800             lblPrevCoag = "No Previous Coag Details"
51810         End If
51820     Else
51830         lblPrevCoag = "No Chart # for Previous Details"
51840     End If

51850     If grdPrev.Rows > 2 Then
51860         grdPrev.RemoveItem 1
51870     End If

51880     Exit Sub

LoadPreviousCoag_Error:

          Dim strES As String
          Dim intEL As Integer

51890     intEL = Erl
51900     strES = Err.Description
51910     LogError "frmEditAll", "LoadPreviousCoag", intEL, strES, sql


End Sub

Private Function LoadSplitList(ByVal Index As Integer) As String

          Dim tb As Recordset
          Dim sql As String
          Dim strIndex As String
          Dim strReturn As String

51920     On Error GoTo LoadSplitList_Error

51930     strIndex = Index

51940     sql = "Select distinct Code, PrintPriority, SplitList " & _
              "from BioTestDefinitions " & _
              "where SplitList = " & strIndex & " " & _
              "order by PrintPriority"

51950     Set tb = New Recordset
51960     RecOpenClient 0, tb, sql

51970     strReturn = ""
51980     Do While Not tb.EOF
51990         strReturn = strReturn & "Code = '" & tb!Code & "' or "
52000         tb.MoveNext
52010     Loop
52020     If strReturn <> "" Then
52030         strReturn = Left$(strReturn, Len(strReturn) - 3)
52040     End If

52050     LoadSplitList = strReturn

52060     Exit Function

LoadSplitList_Error:

          Dim strES As String
          Dim intEL As Integer

52070     intEL = Erl
52080     strES = Err.Description
52090     LogError "frmEditAll", "LoadSplitList", intEL, strES, sql


End Function

Private Sub RemoveFromPhoneAlert(ByVal SampleID As String, _
          ByVal Discipline As String, _
          ByVal Analyte As String)

          Dim sql As String

52100     On Error GoTo RemoveFromPhoneAlert_Error

52110     sql = "DELETE FROM PhoneAlert WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND Discipline = '" & Discipline & "' "
52120     If Analyte <> "All" Then
52130         sql = sql & "AND Parameter = '" & Analyte & "'"
52140     End If

52150     Cnxn(0).Execute sql

52160     Exit Sub

RemoveFromPhoneAlert_Error:

          Dim strES As String
          Dim intEL As Integer

52170     intEL = Erl
52180     strES = Err.Description
52190     LogError "frmEditAll", "RemoveFromPhoneAlert", intEL, strES, sql

End Sub

Private Sub SaveComments()

          Dim OBs As New Observations

52200     On Error GoTo SaveComments_Error

52210     txtSampleID = Format(Val(txtSampleID))
52220     If Val(txtSampleID) = 0 Then Exit Sub



          '+++Junaid 21-07-2023
52230     OBs.Save txtSampleID.Text, True, _
              "Biochemistry", Trim$(txtBioComment), _
              "Demographic", Trim$(txtDemographicComment), _
              "Haematology", Trim$(txtHaemComment), _
              "Coagulation", Trim$(txtCoagComment), _
              "Film", Trim$(txtFilmComment)

          '40    OBs.Save txtSampleID, True, _
          '               "Demographic", Trim$(txtDemographicComment), _
          '               "Haematology", Trim$(txtHaemComment), _
          '               "Coagulation", Trim$(txtCoagComment), _
          '               "Film", Trim$(txtFilmComment)
          '---Junaid

52240     Exit Sub

SaveComments_Error:

          Dim strES As String
          Dim intEL As Integer

52250     intEL = Erl
52260     strES = Err.Description
52270     LogError "frmEditAll", "SaveComments", intEL, strES

End Sub


Private Sub LoadOutstanding()

          Dim tb As Recordset
          Dim sql As String
52280     On Error GoTo ErrorHandler

52290     grdOutstanding.Rows = 2
52300     grdOutstanding.AddItem ""
52310     grdOutstanding.RemoveItem 1

52320     sql = "DELETE FROM BioRequests WHERE Code IN " & _
              "  (SELECT Code FROM BioResults " & _
              "   WHERE SampleID = '" & Val(txtSampleID) & "') " & _
              "AND SampleID = '" & Val(txtSampleID) & "'"
52330     Cnxn(0).Execute sql




52340     sql = "SELECT DISTINCT BT.ShortName,BT.Code, BT.PrintPriority " & _
              "FROM BioRequests BR JOIN BioTestDefinitions BT " & _
              "ON BR.Code = BT.Code " & _
              "WHERE BR.SampleID = '" & Val(txtSampleID) & "' " & _
              "AND BT.InUse = 1 " & _
              "AND BT.SampleType = BR.SampleType " & _
              "ORDER BY BT.PrintPriority"

            
            
            
            
52350     Set tb = New Recordset
52360     RecOpenClient 0, tb, sql
52370     Do While Not tb.EOF
52380         grdOutstanding.AddItem tb!ShortName & vbTab & tb!Code & ""
52390         tb.MoveNext
52400     Loop

          

52410     If grdOutstanding.Rows > 2 Then
52420         grdOutstanding.RemoveItem 1
52430     End If

52440     Exit Sub

ErrorHandler:

          Dim strES As String
          Dim intEL As Integer

52450     intEL = Erl
52460     strES = Err.Description
52470     LogError "frmEditAll", "LoadOutstandingBio", intEL, strES, sql

End Sub
Private Sub LoadOutstandingHaem()

          Dim tb As Recordset
          Dim sql As String

52480     On Error GoTo LoadOutstandingHaem_Error

52490     gOutstandingHaem.Rows = 2
52500     gOutstandingHaem.AddItem ""
52510     gOutstandingHaem.RemoveItem 1
52520     gOutstandingHaem.ColWidth(1) = 0

          '40        sql = "select * from HaemResults where " & _
          '                "SampleID = '" & Val(txtSampleID) & "'"
          '50        Set tb = New Recordset
          '60        RecOpenClient 0, tb, sql
          '70        If Not tb.EOF Then
          '80            gOutstandingHaem.AddItem "FBC"
          '90            If tb!cESR <> 0 Then gOutstandingHaem.AddItem "ESR"
          '100           If tb!cRetics <> 0 Then gOutstandingHaem.AddItem "Retics"
          '110           If tb!cMonospot <> 0 Then gOutstandingHaem.AddItem "MonoSpot"
          '120           If tb!cMalaria <> 0 Then gOutstandingHaem.AddItem "Malaria"
          '130           If tb!cSickledex <> 0 Then gOutstandingHaem.AddItem "Sickledex"
          '140       End If

52530     sql = "SELECT Code as PanelName   FROM HaeRequests WHERE SampleID = '" & Val(txtSampleID) & "' "

52540     Set tb = New Recordset
52550     RecOpenClient 0, tb, sql
52560     Do While Not tb.EOF
52570         gOutstandingHaem.AddItem tb!PanelName
52580         tb.MoveNext
52590     Loop

52600     If gOutstandingHaem.Rows > 2 Then
52610         gOutstandingHaem.RemoveItem 1
52620     End If

52630     Exit Sub

LoadOutstandingHaem_Error:
          Dim strES As String
          Dim intEL As Integer

52640     intEL = Erl
52650     strES = Err.Description
52660     LogError "frmEditAll", "LoadOutstandingHaem", intEL, strES

End Sub

Private Sub LoadOutstandingCoag()

          Dim tb As Recordset
          Dim sql As String

52670     On Error GoTo LoadOutstandingCoag_Error

52680     With gOutstandingCoag
52690         .Rows = 2
52700         .AddItem ""
52710         .RemoveItem 1
52720     End With

52730     sql = "DELETE FROM CoagRequests WHERE Code IN " & _
              "  (SELECT Code FROM CoagResults " & _
              "   WHERE SampleID = '" & Val(txtSampleID) & "') " & _
              "AND SampleID = '" & Val(txtSampleID) & "'"
52740     Cnxn(0).Execute sql

52750     txtSampleID = Format$(Val(txtSampleID))

52760     sql = "SELECT DISTINCT D.TestName " & _
              "FROM CoagRequests C JOIN CoagTestDefinitions D ON D.Code = c.Code " & _
              "WHERE C.SampleID = '" & txtSampleID & "'"
52770     Set tb = New Recordset
52780     RecOpenClient 0, tb, sql
52790     Do While Not tb.EOF
52800         gOutstandingCoag.AddItem tb!TestName & ""
52810         tb.MoveNext
52820     Loop

52830     If gOutstandingCoag.Rows > 2 Then
52840         gOutstandingCoag.RemoveItem 1
52850     End If

52860     Exit Sub

LoadOutstandingCoag_Error:

          Dim strES As String
          Dim intEL As Integer

52870     intEL = Erl
52880     strES = Err.Description
52890     LogError "frmEditAll", "LoadOutstandingCoag", intEL, strES, sql

End Sub

Private Sub SaveBiochemistry(ByVal Validate As Boolean)

          Dim sql As String
          Dim tb As Recordset

52900     On Error GoTo SaveBiochemistry_Error

52910     txtSampleID = Format(Val(txtSampleID))
52920     If Val(txtSampleID) = 0 Then Exit Sub


          '40    sql = "DELETE FROM BioResults " & _
          '      "WHERE SampleID = '" & txtSampleID & "' " & _
          '      "AND (Code = '1071' OR Code = '1072' OR Code ='1073')"
          '50    Cnxn(0).Execute Sql

52930     If Validate Then
52940         sql = "UPDATE BioResults SET " & _
                  "Valid = 1, HealthLink = '0', " & _
                  " ValidateTime = '" & Format$(Now, "dd/MMM/yyyy HH:mm:ss") & "' ," & _
                  "Operator = '" & UserCode & "' " & _
                  "WHERE SampleID = '" & txtSampleID & "' "
              '                & _
              '                "AND (COALESCE(Valid, 0) = 0)"
52950         Cnxn(0).Execute sql

52960         BioValBy = UserCode
52970     End If

52980     SaveMasks

52990     sql = "Select * from Demographics where " & _
              "SampleID = '" & txtSampleID & "'"
53000     Set tb = New Recordset
53010     RecOpenClient 0, tb, sql
53020     If tb.EOF Then
53030         tb.AddNew
53040     End If
53050     If lRandom = "Fasting Sample" Then
53060         tb!Fasting = 1
53070     Else
53080         tb!Fasting = 0
53090     End If
53100     tb!FAXed = 0
53110     tb!RooH = cRooH(0)
53120     tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
53130     If IsDate(tSampleTime) Then
53140         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
53150     Else
53160         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
53170     End If
53180     tb!SampleID = txtSampleID
53190     tb.Update

53200     Exit Sub

SaveBiochemistry_Error:

          Dim strES As String
          Dim intEL As Integer

53210     intEL = Erl
53220     strES = Err.Description
53230     LogError "frmEditAll", "SaveBiochemistry", intEL, strES, sql

End Sub

Private Sub SaveLIH()

          Dim tb As Recordset
          Dim sql As String
          Dim LIH As Integer

53240     On Error GoTo SaveLIH_Error

53250     If Val(txtSampleID) = 0 Then Exit Sub

53260     sql = "IF EXISTS(SELECT LIH FROM masks WHERE SampleID = '" & txtSampleID & "') " & _
              "AND EXISTS(SELECT RESULT FROM BioResults WHERE " & _
              "           SampleID = '" & txtSampleID & "' AND Code = '96') " & _
              "  BEGIN " & _
              "    SELECT M.LIH, R.Result FROM Masks AS M, BioResults AS R WHERE " & _
              "    M.SampleID = '" & txtSampleID & "' " & _
              "    AND R.SampleID = '" & txtSampleID & "' AND R.Code = '96' " & _
              "  END " & _
              "ELSE " & _
              "  IF EXISTS(SELECT LIH FROM masks WHERE SampleID = '" & txtSampleID & "') " & _
              "    BEGIN " & _
              "      SELECT  LIH , -1 as Result FROM Masks WHERE " & _
              "      sampleid = '" & txtSampleID & "' " & _
              "    END " & _
              "  ELSE " & _
              "    IF EXISTS (SELECT Result FROM BioResults WHERE " & _
              "               SampleID = '" & txtSampleID & "' " & _
              "               AND Code = '96' ) " & _
              "      BEGIN " & _
              "        SELECT -1 AS LIH, Result FROM BioResults WHERE " & _
              "        SampleID = '" & txtSampleID & "' " & _
              "        AND Code = '96' " & _
              "      END " & _
              "    ELSE " & _
              "      SELECT -1 AS LIH, -1 AS Result"
53270     Set tb = New Recordset
53280     RecOpenServer 0, tb, sql
53290     If tb!Result = -1 Then
53300         Exit Sub
53310     Else
53320         LIH = Val(Left$(tb!Result & "", 1) * 100) + _
                  Val(Mid$(tb!Result & "", 3, 1) * 10) + _
                  Val(Right$(tb!Result & "", 1))

53330         sql = "SELECT * FROM Masks WHERE " & _
                  "SampleID = '" & txtSampleID & "'"
53340         Set tb = New Recordset
53350         RecOpenServer 0, tb, sql
53360         If tb.EOF Then tb.AddNew
53370         tb!SampleID = txtSampleID
53380         tb!H = 0
53390         tb!S = 0
53400         tb!l = 0
53410         tb!o = 0
53420         tb!g = 0
53430         tb!J = 0
53440         tb!Rundate = Format(dtRunDate, "dd/MMM/yyyy")
53450         tb!LIH = LIH
53460         tb.Update
53470     End If
53480     sql = "Delete from BioResults where " & _
              "SampleID = '" & txtSampleID & "' " & _
              "and Code = '96'"
53490     Cnxn(0).Execute sql

53500     Exit Sub

SaveLIH_Error:

          Dim strES As String
          Dim intEL As Integer

53510     intEL = Erl
53520     strES = Err.Description
53530     LogError "frmEditAll", "SaveLIH", intEL, strES, sql


End Sub

Private Sub SetFormOptions()
53540     On Error GoTo ErrorHandler

53550     Opts(0).Description = "DeptHaem"
53560     Opts(1).Description = "DeptBio"
53570     Opts(2).Description = "DeptCoag"
53580     Opts(3).Description = "DeptExt"
53590     Opts(4).Description = "Urgent"
53600     Opts(5).Description = "BloodBank"
53610     Opts(6).Description = "ActiveDate"

53620     Opts(7).Description = "BioCodeForGlucose"
53630     Opts(8).Description = "BioCodeForChol"
53640     Opts(9).Description = "BioCodeForTrig"

53650     Opts(10).Description = "AllowClinicianFreeText"

53660     LoadFormOptions Opts

53670     frmOptDeptHaem = Opts(0).Value = "1"
53680     frmOptDeptBio = Opts(1).Value = "1"
53690     frmOptDeptCoag = Opts(2).Value = "1"
53700     frmOptDeptExt = Opts(3).Value = "1"
53710     frmOptUrgent = Opts(4).Value = "1"
53720     frmOptBloodBank = Opts(5).Value = "1"
53730     CodeForGlucose = Opts(7).Value
53740     CodeForChol = Opts(8).Value
53750     CodeForTrig = Opts(9).Value

53760     frmOptAllowClinicianFreeText = Opts(10).Value = "1"
53770     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub SetViewHistory()

    '10    Select Case SSTab1.Tab
    '      Case 0: cmdHistory.Visible = False
    '20    Case 1: cmdHistory.Visible = PreviousHaem
    '30    Case 2: cmdHistory.Visible = PreviousBio
    '40    Case 3: cmdHistory.Visible = PreviousCoag
    '50    End Select

End Sub

Private Sub SetViewReports(ByVal Dept As String, ByVal SampleID As String)

          Dim sql As String
          Dim tb As New Recordset

53780     On Error GoTo SetViewReports_Error

53790     cmdViewReports.Visible = False

53800     sql = "SELECT COUNT(*) Tot FROM Reports " & _
              "WHERE SampleID = '" & SampleID & "' " & _
              "AND Dept = '" & Dept & "'"
53810     Set tb = Cnxn(0).Execute(sql)
53820     cmdViewReports.Visible = tb!Tot > 0

53830     Exit Sub

SetViewReports_Error:

          Dim strES As String
          Dim intEL As Integer

53840     intEL = Erl
53850     strES = Err.Description
53860     LogError "frmEditAll", "SetViewReports", intEL, strES, sql

End Sub

Private Sub ShowMenuLists()
53870     On Error GoTo ErrorHandler
53880     mnuListsBio.Visible = False
53890     mnuListsHaem.Visible = False
53900     mnuListsCoag.Visible = False


53910     Select Case SSTab1.Tab
              Case 1:
53920             mnuListsHaem.Visible = UserHasAuthority(UserMemberOf, "HaemLists")
53930         Case 2:
53940             mnuListsBio.Visible = UserHasAuthority(UserMemberOf, "BioLists")
53950             mnuSplits.Visible = UserHasAuthority(UserMemberOf, "BioLists")
53960         Case 3:
53970             mnuListsCoag.Visible = UserHasAuthority(UserMemberOf, "CoagLists")
53980     End Select
53990     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description

End Sub

Private Function TranslateLIH(ByVal BR As BIEResult) As String

          ''Returns "", "1+","2+" etc

          Dim RetVal As String
          Dim x As Integer
          Dim v(1 To 6) As Single

54000     On Error GoTo TranslateLIH_Error

54010     RetVal = ""

54020     If IsNumeric(BR.Result) Then
54030         For x = 6 To 1 Step -1
54040             v(x) = Val(GetOptionSetting("LIH_" & BR.ShortName & Format$(x), "0"))
54050             If v(x) > 0 Then
54060                 If BR.Result >= v(x) Then
54070                     RetVal = Format$(x) & "+"
54080                     Exit For
54090                 End If
54100             End If
54110         Next
54120     End If

54130     TranslateLIH = RetVal

54140     Exit Function

TranslateLIH_Error:

          Dim strES As String
          Dim intEL As Integer

54150     intEL = Erl
54160     strES = Err.Description
54170     LogError "frmEditAll", "TranslateLIH", intEL, strES

End Function



Private Sub btnPrintDoc_Click()
54180     m_ShowDoc = True
54190     m_Notes = ""
54200     m_Notes = Trim(Trim(etc(0).Text) & " " & Trim(etc(1).Text) & " " & Trim(etc(2).Text) & " " & Trim(etc(3).Text) & " " & Trim(etc(4).Text) & " " & Trim(etc(5).Text) & " " & Trim(etc(6).Text) & " " & Trim(etc(7).Text) & " " & Trim(etc(8).Text))
54210     DoEvents
54220     DoEvents
54230     Call cmdOrderExt_Click(1)
          
End Sub

Private Sub cClDetails_KeyPress(KeyAscii As Integer)
54240     KeyAscii = AutoComplete(cClDetails, KeyAscii, False)

End Sub

Private Sub chkBioReject_Click()
54250     cmdSaveHoldDemographics.Enabled = True
54260     cmdSaveDemographics.Enabled = True
End Sub

Private Sub chkCoagReject_Click()
54270     cmdSaveHoldDemographics.Enabled = True
54280     cmdSaveDemographics.Enabled = True
End Sub

Private Sub ChkExtReject_Click()
54290     cmdSaveHoldDemographics.Enabled = True
54300     cmdSaveDemographics.Enabled = True
End Sub

Private Sub chkHaemReject_Click()
54310     cmdSaveHoldDemographics.Enabled = True
54320     cmdSaveDemographics.Enabled = True
End Sub



Private Sub cmbAdd_LostFocus()
54330     On Error GoTo cmbAdd_LostFocus_Error

          Dim i As Integer

54340     For i = 0 To cmbAdd.ListCount - 1
54350         If UCase(cmbAdd) = UCase(cmbAdd.List(i)) Then
54360             cmbAdd.ListIndex = i
54370             Exit For
54380         End If
54390     Next i

54400     Exit Sub

cmbAdd_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

54410     intEL = Erl
54420     strES = Err.Description
54430     LogError "frmEditAll", "cmbAdd_LostFocus", intEL, strES

End Sub





Private Sub cmbOtherSamples_Click()
54440     On Error GoTo cmbOtherSamples_Click_Error

54450     If Left(cmbOtherSamples.Text, 6) = "Sample" Then
54460         Call txtsampleid_LostFocus
54470     Else
54480         txtSampleID.Text = Left(cmbOtherSamples.Text, 7)
54490         Call txtsampleid_LostFocus
54500     End If

54510     Exit Sub

cmbOtherSamples_Click_Error:
          
          Dim strES As String
          Dim intEL As Integer

54520     intEL = Erl
54530     strES = Err.Description
54540     LogError "frmEditAll", "cmbOtherSamples_Click", intEL, strES
End Sub

Private Sub cmdAddBio_Click()

          Dim sql As String
          Dim Code As String

54550     On Error GoTo cmdAdd_Click_Error

54560     pBar = 0

          'If cmbAdd.Text = "" Then
          '    MsgBox "Please fill required credentials.", vbInformation
          '    Exit Sub
          'End If
          'If cmbNewResult.Text = "" Then
          '    MsgBox "Please fill required credentials.", vbInformation
          '    Exit Sub
          'End If
          'If cmbUnits.Text = "" Then
          '    MsgBox "Please fill required credentials.", vbInformation
          '    Exit Sub
          'End If

54570     txtSampleID = Format(Val(txtSampleID))
54580     If Trim$(txtSampleID) = "" Then Exit Sub
          '+++ Junaid 26-05-2024

          '--- Junaid
54590     If cmbAdd.Text = "" Then Exit Sub
54600     Code = lstAdd.List(cmbAdd.ListIndex)



54610     sql = "IF EXISTS(SELECT * FROM BioResults " & _
              "          WHERE SampleID = @sampleid0 " & _
              "          AND Code = '@Code1' ) " & _
              "  INSERT INTO BioRepeats " & _
              "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
              "   Units, SampleType, Analyser, Faxed, " & _
              "   Healthlink) VALUES " & _
              "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
              "  '@Units9', '@SampleType10', '@Analyser11', " & _
              "  @Faxed12, @Healthlink18) " & _
              "ELSE " & _
              "  INSERT INTO BioResults " & _
              "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
              "   Units, SampleType, Analyser, Faxed, " & _
              "   Healthlink) VALUES " & _
              "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
              "  '@Units9', '@SampleType10', '@Analyser11', " & _
              "  @Faxed12, @Healthlink18) "

54620     sql = Replace(sql, "@sampleid0", txtSampleID)
54630     sql = Replace(sql, "@Code1", Code)
54640     sql = Replace(sql, "@result2", cmbNewResult)
54650     sql = Replace(sql, "@valid3", 0)
54660     sql = Replace(sql, "@printed4", 0)
54670     sql = Replace(sql, "@RunTime5", Format$(Now, "'dd/mmm/yyyy hh:mm:ss'"))
54680     sql = Replace(sql, "@RunDate6", Format$(Now, "'dd/mmm/yyyy'"))
54690     sql = Replace(sql, "@Units9", cmbUnits)
54700     sql = Replace(sql, "@SampleType10", ListCodeFor("ST", cmbSampleType))
54710     sql = Replace(sql, "@Analyser11", "Manual")
54720     sql = Replace(sql, "@Faxed12", 0)
54730     sql = Replace(sql, "@Healthlink18", 0)
          'MsgBox Sql
54740     Cnxn(0).Execute sql

54750     CheckIfMustPhone "Biochemistry", cmbAdd.Text, cmbNewResult
54760     LoadBiochemistry

54770     cmbAdd = ""
54780     cmbUnits = ""
54790     cmbNewResult = ""

54800     Exit Sub

cmdAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

54810     intEL = Erl
54820     strES = Err.Description
54830     LogError "frmEditAll", "cmdAdd_Click", intEL, strES, sql

End Sub


Private Sub bAddCoag_Click()

          Dim Units As String
          Dim S As String
          Dim tb As Recordset
          Dim sql As String

54840     On Error GoTo bAddCoag_Click_Error

54850     pBar = 0

54860     If cParameter = "" Then Exit Sub
54870     If Trim$(tResult) = "" Then Exit Sub

54880     sql = "Select Units from CoagTestDefinitions where " & _
              "TestName = '" & cParameter & "'"
54890     Set tb = New Recordset
54900     RecOpenServer 0, tb, sql
54910     If Not tb.EOF Then
54920         Units = tb!Units & ""
54930     End If

54940     S = cParameter & vbTab & _
              tResult & vbTab & _
              Units & vbTab & _
              vbTab & _
              "V"
54950     gCoag.AddItem S

54960     CheckIfMustPhone "Coagulation", cParameter, tResult

54970     If gCoag.TextMatrix(1, 0) = "" Then
54980         gCoag.RemoveItem 1
54990     End If

55000     cParameter = ""
55010     tResult = ""
55020     cmdSaveCoag.Enabled = True
55030     cmdValidateCoag.Enabled = True

55040     SaveCoag True

55050     SetPrintInhibit "Coa"
55060     LoadCoagulation

55070     Exit Sub

bAddCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

55080     intEL = Erl
55090     strES = Err.Description
55100     LogError "frmEditAll", "bAddCoag_Click", intEL, strES, sql


End Sub

Private Sub chkRA_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
55110     On Error GoTo ErrorHandler
55120     If chkRA = 0 Then
55130         If lblRA = "?" Then
55140             lblRA = ""
55150         ElseIf lblRA <> "" Then
55160             chkRA = 1
55170         End If
55180     Else
55190         If lblRA = "" Then
55200             lblRA = "?"
55210         End If
55220     End If

55230     cmdSaveHaem.Enabled = True
55240     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub


Private Sub chkUrgent_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

55250     cmdSaveHoldDemographics.Enabled = True
55260     cmdSaveDemographics.Enabled = True

End Sub




Private Sub cmbBioComment_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

55270     On Error GoTo cmbBioComment_Click_Error

55280     If Index = 0 Then
55290         txtBioComment = Trim$(txtBioComment & " " & cmbBioComment(0))
55300         cmbBioComment(0) = ""

55310         cmdSaveBio.Enabled = True
55320     Else
55330         If cmbBioComment(Index).Text = "" Then Exit Sub
55340         sql = "Select * From CommentsTemplate Where " & _
                  "CommentID = " & cmbBioComment(Index).ItemData(cmbBioComment(Index).ListIndex)
55350         Set tb = New Recordset
55360         RecOpenClient 0, tb, sql
55370         If Not tb.EOF Then
55380             Select Case Index
                      Case 1: txtBioComment = txtBioComment & tb!CommentTemplate
55390                 Case 2: txtCoagComment = txtCoagComment & tb!CommentTemplate
55400             End Select
55410         End If
55420         cmbBioComment(Index).Text = "*** Insert Comment Template ***"

55430     End If

55440     Exit Sub

cmbBioComment_Click_Error:

          Dim strES As String
          Dim intEL As Integer

55450     intEL = Erl
55460     strES = Err.Description
55470     LogError "frmEditAll", "cmbBioComment_Click", intEL, strES, sql

End Sub




Private Sub cmbBioComment_KeyPress(Index As Integer, KeyAscii As Integer)

55480     If Index = 0 Then
55490         cmdSaveBio.Enabled = True
55500     Else
55510         KeyAscii = 0
55520     End If

End Sub










Private Sub cmbBioComment_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

55530     On Error GoTo cmbBioComment_LostFocus_Error

55540     If Index = 0 Then
55550         sql = "SELECT Text FROM Lists WHERE " & _
                  "ListType = 'BI' " & _
                  "AND Code = '" & AddTicks(cmbBioComment(0)) & "' " & _
                  "AND InUse = 1"
55560         Set tb = New Recordset
55570         RecOpenServer 0, tb, sql
55580         If Not tb.EOF Then
55590             txtBioComment = Trim$(txtBioComment & " " & tb!Text & "")
55600         Else
55610             txtBioComment = Trim$(txtBioComment & " " & cmbBioComment(0))
55620         End If
55630         cmbBioComment(0) = ""
55640     End If

55650     Exit Sub

cmbBioComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

55660     intEL = Erl
55670     strES = Err.Description
55680     LogError "frmEditAll", "cmbBioComment_LostFocus", intEL, strES, sql

End Sub

Private Sub cmbClinician_Change()

55690     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP & ":" & cmbClinician



End Sub



Private Sub cmbDemogComment_Click()

55700     txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
55710     cmbDemogComment = ""

55720     cmdSaveHoldDemographics.Enabled = True
55730     cmdSaveDemographics.Enabled = True

End Sub

Private Sub cmbDemogComment_KeyPress(KeyAscii As Integer)

55740     KeyAscii = 0
55750     KeyAscii = AutoComplete(cmbDemogComment, KeyAscii, False)


End Sub


Private Sub cmbDemogComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

55760     On Error GoTo cmbDemogComment_LostFocus_Error

55770     sql = "Select * from Lists where " & _
              "ListType = 'DE' " & _
              "and Code = '" & cmbDemogComment & "' and InUse = 1"
55780     Set tb = New Recordset
55790     RecOpenServer 0, tb, sql
55800     If Not tb.EOF Then
55810         txtDemographicComment = Trim$(txtDemographicComment & " " & tb!Text & "")
55820     Else
55830         txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
55840     End If
55850     cmbDemogComment = ""

55860     Exit Sub

cmbDemogComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

55870     intEL = Erl
55880     strES = Err.Description
55890     LogError "frmEditAll", "cmbDemogComment_LostFocus", intEL, strES, sql


End Sub

Private Sub cmbFilmComment_Click()

55900     txtFilmComment = Trim$(txtFilmComment & " " & cmbFilmComment)
55910     cmbFilmComment = ""

55920     cmdSaveHaem.Enabled = True

End Sub

Private Sub cmbFilmComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

55930     On Error GoTo cmbFilmComment_LostFocus_Error

55940     sql = "Select * from Lists where " & _
              "ListType = 'FI' " & _
              "and Code = '" & cmbFilmComment & "' and InUse = 1"
55950     Set tb = New Recordset
55960     RecOpenServer 0, tb, sql
55970     If Not tb.EOF Then
55980         txtFilmComment = Trim$(txtFilmComment & " " & tb!Text & "")
55990     Else
56000         txtFilmComment = Trim$(txtFilmComment & " " & cmbFilmComment)
56010     End If
56020     cmbFilmComment = ""

56030     Exit Sub

cmbFilmComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

56040     intEL = Erl
56050     strES = Err.Description
56060     LogError "frmEditAll", "cmbFilmComment_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbHaemComment_Click()

56070     txtHaemComment = Trim$(txtHaemComment & " " & cmbHaemComment)
56080     cmbHaemComment = ""

56090     cmdSaveHaem.Enabled = True

End Sub

Private Sub cmbHaemComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

56100     On Error GoTo cmbHaemComment_LostFocus_Error

56110     sql = "Select * from Lists where " & _
              "ListType = 'HA' " & _
              "and Code = '" & AddTicks(cmbHaemComment) & "' and InUse = 1"
56120     Set tb = New Recordset
56130     RecOpenServer 0, tb, sql
56140     If Not tb.EOF Then
56150         txtHaemComment = Trim$(txtHaemComment & " " & tb!Text & "")
56160     Else
56170         txtHaemComment = Trim$(txtHaemComment & " " & cmbHaemComment)
56180     End If
56190     cmbHaemComment = ""

56200     Exit Sub

cmbHaemComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

56210     intEL = Erl
56220     strES = Err.Description
56230     LogError "frmEditAll", "cmbHaemComment_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbHospital_Change()

56240     FillWards cmbWard, cmbHospital
56250     FillClinicians cmbClinician, cmbHospital
56260     FillGPs cmbGP, cmbHospital
          'txtExtSampleID.Visible = (UCase(cmbHospital) <> UCase(HospName(0)))
          'lblExtSampleID.Visible = txtExtSampleID.Visible
End Sub

Private Sub cmbHospital_Click()

56270     FillWards cmbWard, cmbHospital
56280     FillClinicians cmbClinician, cmbHospital
56290     FillGPs cmbGP, cmbHospital

56300     cmdSaveHoldDemographics.Enabled = True
56310     cmdSaveDemographics.Enabled = True

          'txtExtSampleID.Visible = (UCase(cmbHospital) <> UCase(HospName(0)))
          'lblExtSampleID.Visible = txtExtSampleID.Visible
End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

56320     KeyAscii = 0

End Sub


Private Sub cmdCopyFromPrevious_Click()
    
'          Dim tb As Recordset
'          Dim sql As String
'          Dim PrevSID As Long
'          Dim OB As Observation
'          Dim OBs As Observations
'
'10        On Error GoTo cmdCopyFromPrevious_Click_Error
'
'20        PrevSID = Val(txtSampleID) - 1
'
'        '+++Junaid 15-10-2023
'        If MsgBox("Are you sure you want to copy all details from " & PrevSID & "?", vbInformation + vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'        '---Junaid
'
'30        sql = "Select * from Demographics where " & _
'                "SampleID = " & PrevSID
'40        Set tb = New Recordset
'50        RecOpenServer 0, tb, sql
'
'60        If Trim$(tb!Hospital & "") <> "" Then
'70            cmbHospital = Trim$(tb!Hospital)
'80            lblChartNumber = Trim$(tb!Hospital) & " Chart #" & "                     SurName                                         ForeName"
'90            If UCase$(tb!Hospital) = UCase$(HospName(0)) Then
'100               lblChartNumber.BackColor = &H8000000F
'110               lblChartNumber.ForeColor = vbBlack
'120           Else
'130               lblChartNumber.BackColor = vbRed
'140               lblChartNumber.ForeColor = vbYellow
'150           End If
'160       Else
'170           cmbHospital = HospName(0)
'180           lblChartNumber.Caption = HospName(0) & " Chart #" & "                     SurName                                         ForeName"
'190           lblChartNumber.BackColor = &H8000000F
'200           lblChartNumber.ForeColor = vbBlack
'210       End If
'220       If IsDate(tb!SampleDate) Then
'230           dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
'240       Else
'250           dtSampleDate = Format$(Now, "dd/mm/yyyy")
'260       End If
'270       If IsDate(tb!Rundate) Then
'280           dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
'290       Else
'300           dtRunDate = Format$(Now, "dd/mm/yyyy")
'310       End If
'320       StatusBar1.Panels(4).Text = dtRunDate
'330       mNewRecord = False
'340       If Not IsNull(tb!RooH) Then
'350           cRooH(0) = IIf(tb!RooH = 1, True, False)
'360           cRooH(1) = Not tb!RooH
'370       Else
'380           cRooH(0) = True
'390       End If
'400       txtChart = tb!Chart & ""
'410       txtLabNo = tb!LabNo & ""
'420       txtSurName = SurName(tb!PatName & "")
'430       txtForeName = ForeName(tb!PatName & "")
'440       txtAddress(0) = tb!Addr0 & ""
'450       txtAddress(1) = tb!Addr1 & ""
'460       Select Case Left$(Trim$(UCase$(tb!Sex & "")), 1)
'          Case "M": txtSex = "Male"
'470       Case "F": txtSex = "Female"
'480       Case Else: txtSex = ""
'490       End Select
'500       txtDoB = Format$(tb!DoB, "dd/mm/yyyy")
'510       txtAge = tb!Age & ""
'520       cmbWard = tb!Ward & ""
'530       cmbClinician = tb!Clinician & ""
'540       cmbGP = tb!GP & ""
'550       cClDetails = tb!ClDetails & ""
'560       If IsDate(tb!SampleDate) Then
'570           dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
'580           If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
'590               tSampleTime = Format$(tb!SampleDate, "hh:mm")
'600           Else
'610               tSampleTime.Mask = ""
'620               tSampleTime.Text = ""
'630               tSampleTime.Mask = "##:##"
'640           End If
'650       Else
'660           dtSampleDate = Format$(Now, "dd/mm/yyyy")
'670           tSampleTime.Mask = ""
'680           tSampleTime.Text = ""
'690           tSampleTime.Mask = "##:##"
'700       End If
'710       If Not IsNull(tb!Fasting) Then
'720           If tb!Fasting Then
'730               lRandom = "Fasting Sample"
'740           End If
'750       End If
'760       lblDemogValid = "Demographics Not Valid"
'770       lblDemogValid.BackColor = vbRed
'780       lblDemogValid.ForeColor = vbYellow
'790       cmdValidateDemographics.Visible = False
'
'800       If IsDate(tb!RecDate & "") Then
'810           dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
'820           If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
'830               tRecTime = Format$(tb!RecDate, "hh:mm")
'840           Else
'850               tRecTime.Mask = ""
'860               tRecTime.Text = ""
'870               tRecTime.Mask = "##:##"
'880           End If
'890       Else
'900           dtRecDate = Format$(Now, "dd/mm/yyyy")
'910           tRecTime.Mask = ""
'920           tRecTime.Text = ""
'930           tRecTime.Mask = "##:##"
'940       End If
'950       If frmOptUrgent Then
'960           If tb!Urgent Then
'970               lblUrgent.Visible = True
'980               chkUrgent.Value = 1
'990               UrgentTest = True
'1000          Else
'1010              chkUrgent.Value = 0
'1020              UrgentTest = False
'1030          End If
'1040      End If
'
'1050      cmdSaveHoldDemographics.Enabled = True
'1060      cmdSaveDemographics.Enabled = True
'
'1070      If frmOptBloodBank Then
'1080          If Trim$(txtChart) <> "" Then
'1090              sql = "Select  * from PatientDetails where " & _
'                        "PatNum = '" & txtChart & "'"
'1100              Set tb = New Recordset
'1110              RecOpenClientBB 0, tb, sql
'1120              bViewBB.Enabled = Not tb.EOF
'1130          End If
'1140      End If
'
'1150      Set OBs = New Observations
'1160      Set OBs = OBs.Load(PrevSID, "Demographic")
'1170      If Not OBs Is Nothing Then
'1180          Set OB = OBs.Item(1)
'1190          txtDemographicComment = OB.Comment
'1200      End If
'
'1210      CopyCC (PrevSID)
'1220      CheckCC
'
'1230      cmdCopyFromPrevious.Visible = False
'
'1240      Exit Sub
'
'cmdCopyFromPrevious_Click_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'1250      intEL = Erl
'1260      strES = Err.Description
'1270      LogError "frmEditAll", "cmdCopyFromPrevious_Click", intEL, strES, sql


End Sub

Private Sub cmdCopyTo_Click()

          Dim S As String

56330     S = cmbWard & " " & cmbClinician
56340     S = Trim$(S) & " " & cmbGP
56350     S = Trim$(S)

56360     frmCopyTo.lblOriginal = S
56370     frmCopyTo.lblSampleID = txtSampleID
56380     frmCopyTo.Show 1

56390     CheckCC

End Sub

Private Sub cmdMedibridge_Click()

          Dim medibridgepathtoviewer As String

56400     On Error GoTo cmdMedibridge_Click_Error

56410     medibridgepathtoviewer = GetOptionSetting("MedibridgePathToViewer", "")
56420     If medibridgepathtoviewer <> "" Then
56430         Shell medibridgepathtoviewer & " /SampleID=" & txtSampleID & _
                  " /UserName=""" & UserName & """" & _
                  " /Password=""" & TechnicianPassFor(UserName) & """" & _
                  " /Department=Medibridge", vbNormalFocus
56440     End If

56450     Exit Sub

cmdMedibridge_Click_Error:

          Dim strES As String
          Dim intEL As Integer

56460     intEL = Erl
56470     strES = Err.Description
56480     LogError "frmEditAll", "cmdMedibridge_Click", intEL, strES

End Sub

Private Sub cmdPatientNotePad_Click()
56490     On Error GoTo cmdPatientNotePad_Click_Error

56500     frmPatientNotePad.SampleID = txtSampleID
56510     frmPatientNotePad.Caller = "General"
56520     frmPatientNotePad.Show 1

56530     Exit Sub

cmdPatientNotePad_Click_Error:

          Dim strES As String
          Dim intEL As Integer

56540     intEL = Erl
56550     strES = Err.Description
56560     LogError "frmEditAll", "cmdPatientNotePad_Click", intEL, strES


End Sub

Private Sub cmdPrintMalaria_Click()

          Dim sql As String
          Dim tb As Recordset

56570     On Error GoTo cmdPrintMalaria_Click_Error

56580     pBar = 0

56590     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
56600         Exit Sub
56610     End If

56620     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

56630     SaveDemographics 1

56640     SaveHaematology 1

56650     LogTimeOfPrinting txtSampleID, "H"

56660     sql = "Select * from PrintPending where " & _
              "Department = 'L' " & _
              "and SampleID = '" & txtSampleID & "'"
56670     Set tb = New Recordset
56680     RecOpenClient 0, tb, sql
56690     If tb.EOF Then
56700         tb.AddNew
56710     End If
56720     tb!SampleID = txtSampleID
56730     tb!Ward = cmbWard
56740     tb!Clinician = cmbClinician
56750     tb!GP = cmbGP
56760     tb!Department = "L"
56770     tb!Initiator = HaemValBy
56780     tb!UsePrinter = pPrintToPrinter
56790     tb.Update

56800     sql = "Update HaemResults " & _
              "Set Printed = 1 where " & _
              "SampleID = '" & txtSampleID & "'"
56810     Cnxn(0).Execute sql


56820     Exit Sub

cmdPrintMalaria_Click_Error:

          Dim strES As String
          Dim intEL As Integer

56830     intEL = Erl
56840     strES = Err.Description
56850     LogError "frmEditAll", "cmdPrintMalaria_Click", intEL, strES, sql

End Sub

Private Sub cmdPrintSickledex_Click()

          Dim tb As Recordset
          Dim sql As String

56860     On Error GoTo cmdPrintSickledex_Click_Error

56870     pBar = 0

56880     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
56890         Exit Sub
56900     End If

56910     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

56920     SaveDemographics 1

56930     SaveHaematology 1

56940     LogTimeOfPrinting txtSampleID, "H"

56950     sql = "Select * from PrintPending where " & _
              "Department = 'J' " & _
              "and SampleID = '" & txtSampleID & "'"
56960     Set tb = New Recordset
56970     RecOpenClient 0, tb, sql
56980     If tb.EOF Then
56990         tb.AddNew
57000     End If
57010     tb!SampleID = txtSampleID
57020     tb!Ward = cmbWard
57030     tb!Clinician = cmbClinician
57040     tb!GP = cmbGP
57050     tb!Department = "J"
57060     tb!Initiator = HaemValBy
57070     tb!UsePrinter = pPrintToPrinter
57080     tb.Update

57090     sql = "Update HaemResults " & _
              "Set Printed = 1 where " & _
              "SampleID = '" & txtSampleID & "'"
57100     Cnxn(0).Execute sql

57110     Exit Sub

cmdPrintSickledex_Click_Error:

          Dim strES As String
          Dim intEL As Integer

57120     intEL = Erl
57130     strES = Err.Description
57140     LogError "frmEditAll", "cmdPrintSickledex_Click", intEL, strES, sql

End Sub



Private Sub cmdSavebio_Click()
57150     On Error GoTo ErrorHandler

57160     pBar = 0
          'Zyam 15-06-24
57170     cmdUpDown(0).Enabled = False
57180     cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

57190     If UserHasAuthority(UserMemberOf, "BioSave") = False Then
57200         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
57210         Exit Sub
57220     End If

57230     txtSampleID = Format(Val(txtSampleID))
57240     If Val(txtSampleID) = 0 Then Exit Sub

57250     SaveBiochemistry False
          'DoEvents
          ''+++Abubaker 24-11-23 (Temporary stoping it)
          '        If m_Code = "SYS" Then
          '            m_Code = ""
          '        Else
57260     SaveComments
          '        End If
          ''---Abubaker 24-11-23

57270     UpdateMRU Me
          'txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1

57280     Debug.Print "cmdsavebio_click"
57290     LoadAllDetails
          'Zyam 15-06-24
57300     cmdUpDown(0).Enabled = True
57310     cmdUpDown(1).Enabled = True
          'Zyam 15-06-24
57320     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub cmdSaveCoag_Click()
57330     On Error GoTo ErrorHandler

57340     pBar = 0
          'Zyam 15-06-24
57350     cmdUpDown(0).Enabled = False
57360     cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

57370     If UserHasAuthority(UserMemberOf, "CoagSave") = False Then
57380         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
57390         Exit Sub
57400     End If

57410     txtSampleID = Format(Val(txtSampleID))
57420     If Val(txtSampleID) = 0 Then Exit Sub

57430     SaveCoag False
          'DoEvents
57440     SaveComments
57450     UpdateMRU Me
          'txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1
57460     Debug.Print "cmdsavecoag_click"
57470     LoadAllDetails
          'Zyam 15-06-24
57480     cmdUpDown(0).Enabled = True
57490     cmdUpDown(1).Enabled = True
          'Zyam 15-06-24
57500     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub cmdSaveDemographics_Click()
57510     On Error GoTo ErrorHandler

57520     pBar = 0

57530     If UserHasAuthority(UserMemberOf, "DemSave") = False Then
57540         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
57550         Exit Sub
57560     End If

57570     txtSampleID = Format(Val(txtSampleID))
57580     If Val(txtSampleID) = 0 Then Exit Sub

57590     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
57600         Exit Sub
57610     End If
57620     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

57630     If lblChartNumber.BackColor = vbRed Then
57640         If iMsg("Confirm this Patient has" & vbCrLf & _
                  lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
57650             Exit Sub
57660         End If
57670     End If

57680     SaveDemographics 0
57690     SaveComments
57700     DoEvents
57710     SaveRejectedSample
57720     UpdateMRU Me
          '210   txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1

          '210   txtLabNo = Val(FndMaxID("demographics", "LabNo", ""))

57730     LoadAllDetails

57740     Exit Sub

ErrorHandler:

          Dim strES As String
          Dim intEL As Integer

57750     intEL = Erl
57760     strES = Err.Description
57770     LogError "frmEditAll", "cmdSave_Click", intEL, strES
End Sub


Private Sub cmdSaveExt_Click()
57780     On Error GoTo ErrorHandler
57790     pBar = 0

57800     If UserHasAuthority(UserMemberOf, "ExtSave") = False Then
57810         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
57820         Exit Sub
57830     End If

57840     txtSampleID = Format(Val(txtSampleID))
57850     If Val(txtSampleID) = 0 Then Exit Sub

57860     SaveExtern
57870     DoEvents
57880     SaveComments
57890     UpdateMRU Me
57900     Debug.Print "cmdsaveext_click"
57910     LoadAllDetails
57920     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub cmdSaveHaem_Click()
57930     On Error GoTo ErrorHandler
          'Zyam 15-06-24
57940     cmdUpDown(0).Enabled = False
57950     cmdUpDown(1).Enabled = False
          'Zyam 15-06-24
57960     pBar = 0

57970     If UserHasAuthority(UserMemberOf, "HaemSave") = False Then
57980         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
57990         Exit Sub
58000     End If

58010     txtSampleID = Format(Val(txtSampleID))
58020     If Val(txtSampleID) = 0 Then Exit Sub

58030     SaveHaematology 0
          'DoEvents
58040     SaveComments
58050     UpdateMRU Me
58060     Call DeleteOutstandings
          'txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1
58070     Debug.Print "cmdsavehaem_click"
58080     LoadAllDetails
          'Zyam 15-06-24
58090     cmdUpDown(0).Enabled = True
58100     cmdUpDown(1).Enabled = True
          'Zyam 15-06-24
58110     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub cmdTag_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim f As Form
          Dim Comment As String

58120     On Error GoTo cmdTag_Click_Error

58130     Comment = ""
58140     sql = "SELECT * FROM MicroTag WHERE Chart = '" & txtChart & "'"
58150     Set tb = New Recordset
58160     RecOpenServer 0, tb, sql
58170     If Not tb.EOF Then
58180         Comment = tb!Comment & ""
58190     End If
58200     Set f = New frmComment
58210     With f
58220         .Comment = Comment
58230         .Show 1
58240         Comment = .Comment
58250         Unload f
58260         Set f = Nothing
58270     End With

58280     If Trim$(Comment) <> "" Then
58290         sql = "IF EXISTS (SELECT * FROM MicroTag WHERE Chart = '" & txtChart & "') " & _
                  "  UPDATE MicroTag SET Comment = '" & AddTicks(Comment) & "', " & _
                  "  UserName = '" & AddTicks(UserName) & "' " & _
                  "  WHERE Chart = '" & txtChart & "' " & _
                  "ELSE " & _
                  "  INSERT INTO MicroTag (Chart, Comment, UserName) " & _
                  "  VALUES " & _
                  "  ('" & txtChart & "', " & _
                  "  '" & AddTicks(Comment) & "', " & _
                  "  '" & AddTicks(UserName) & "')"
58300         Cnxn(0).Execute sql
58310     Else
58320         sql = "DELETE FROM MicroTag WHERE Chart = '" & txtChart & "'"
58330         Cnxn(0).Execute sql
58340     End If

58350     CheckTag

58360     Exit Sub

cmdTag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

58370     intEL = Erl
58380     strES = Err.Description
58390     LogError "frmEditAll", "cmdTag_Click", intEL, strES, sql

End Sub

Private Sub cmdUnLock_Click()


58400     On Error GoTo cmdUnLock_Click_Error

58410     If UserHasAuthority(UserMemberOf, "DemUnlock") = False Then
58420         iMsg "You do not have authority to unlock demographics" & vbCrLf & "Please contact system administrator"
58430         Exit Sub
58440     End If
58450     If UCase(iBOX("Enter password to unValidate ?", , , True)) = UCase$(TechnicianPassFor(UserName)) Then

58460         LockDemographics Me, False
58470         txtSurName.SetFocus


58480         SaveDemographics False
58490         EnableDemographicEntry True
58500         lblDemogValid = "Demographics Not Valid"
58510         lblDemogValid.BackColor = vbRed
58520         lblDemogValid.ForeColor = vbYellow
58530         cmdValidateDemographics.Visible = False

58540     End If




58550     Exit Sub

cmdUnLock_Click_Error:
          Dim strES As String
          Dim intEL As Integer

58560     intEL = Erl
58570     strES = Err.Description
58580     LogError "frmEditAll", "cmdUnLock_Click", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdUpDown_MouseDown
' Author    : XPMUser
' Date      : 03/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdUpDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

58590     On Error GoTo cmdUpDown_MouseDown_Error


58600     If Index = 0 Then
58610         UpDownDirection = 1
58620     Else
58630         UpDownDirection = -1
58640     End If

          'txtLabNo = Val(FndMaxID("demographics", "LabNo", "")) + 1
58650     frmEditAll.txtMultiSeltdDemoForLabNoUpd = ""
58660     txtSampleID = Val(txtSampleID) + UpDownDirection

58670     If Val(txtSampleID) < 1 Then
58680         txtSampleID = "1"
              '110   ElseIf Val(txtSampleID) > 9999999 Then
              '120       txtSampleID = "9999999"
58690     End If

          'tmrUpDown.Enabled = True
          'LoadAllDetails


58700     Exit Sub


cmdUpDown_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

58710     intEL = Erl
58720     strES = Err.Description
58730     LogError "frmEditAll", "cmdUpDown_MouseDown", intEL, strES

End Sub


Private Sub cmdUpDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

58740     UpDownDirection = 0

58750     tmrUpDown.Enabled = False

58760     pBar = 0

          'iMsg "", vbOKOnly

58770     LoadAllDetails

58780     cmdSaveHoldDemographics.Enabled = False
58790     cmdSaveDemographics.Enabled = False
58800     cmdSaveHaem.Enabled = False
58810     cmdSaveBio.Enabled = False
58820     cmdSaveCoag.Enabled = False

End Sub


Private Sub cmdValidateBio_Click()

58830     pBar = 0
          'Zyam 15-06-24
58840     cmdUpDown(0).Enabled = False
58850     cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

58860     If UserHasAuthority(UserMemberOf, "BioSave") = False Then
58870         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
58880         Exit Sub
58890     End If

58900     SaveBiochemistry True
58910     SaveComments
          'DoEvents
          'txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1
58920     Debug.Print "cmdvalidatebio_click"
58930     LoadAllDetails

58940     UpdateMRU Me
          'Zyam 15-06-24
58950     cmdUpDown(0).Enabled = True
58960     cmdUpDown(1).Enabled = True
          'Zyam 15-06-24

End Sub

Private Sub cmdValidateCoag_Click()

58970     pBar = 0
          'Zyam 15-06-24
58980     cmdUpDown(0).Enabled = False
58990     cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

59000     If UserHasAuthority(UserMemberOf, "CoagSave") = False Then
59010         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
59020         Exit Sub
59030     End If

59040     SaveCoag True
59050     SaveComments
          'DoEvents
          'txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1
59060     Debug.Print "cmdvalidatecoag_click"
59070     LoadAllDetails

59080     UpdateMRU Me
          'Zyam 15-06-24
59090     cmdUpDown(0).Enabled = True
59100     cmdUpDown(1).Enabled = True
          'Zyam 15-06-24

End Sub


Private Sub cmdValidateDemographics_Click()

59110     pBar = 0

          '30    Select Case Index
          '        Case 0: 'Demographics
          '40        If cmdValidate(0).Caption = "&Validate" Then
          '50          If EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.text) Then
          '60            SaveDemographics True
          '70            EnableDemographicEntry False
          '80          End If
          '90        Else
59120     If UCase(iBOX("Enter password to unValidate ?", , , True)) = UCase$(TechnicianPassFor(UserName)) Then
59130         SaveDemographics False
59140         EnableDemographicEntry True
59150         lblDemogValid = "Demographics Not Valid"
59160         lblDemogValid.BackColor = vbRed
59170         lblDemogValid.ForeColor = vbYellow
59180         cmdValidateDemographics.Visible = False
59190     End If
          '  Exit Sub
          '160       End If

          '  Case 1: SaveHaematology True

          '  Case 2: SaveBiochemistry True

          '  Case 3: SaveCoag True

          'End Select

          'SaveComments
          'txtSampleID = Val(txtSampleID + 1) ' MoveToNextRelevant 1
          'LoadAllDetails

          'UpdateMRU Me

End Sub

Private Sub cmdDiff_Click()

59200     With frmDifferentials
59210         .lWBC = tWBC
59220         .Show 1
59230     End With

End Sub

Private Sub cmdFAX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

          Dim tb As Recordset
          Dim sql As String
          Dim FaxNumber As String
          Dim Department As String
59240     ReDim strFaxNumbers(0 To 0) As String
          Dim f As Form
          Dim n As Integer

59250     On Error GoTo cmdFAX_MouseUp_Error

59260     pBar = 0

59270     If UserHasAuthority(UserMemberOf, Choose(SSTab1.Tab, "Haem", "Bio", "Coag", "Imm", "End", "Ext") & "Fax") = False Then
59280         iMsg "You do not have authority to fax" & vbCrLf & "Please contact system administrator"
59290         Exit Sub
59300     End If


59310     If Button = vbRightButton Then
59320         If cmdFAX.BackColor = vbYellow Then
59330             frmPhoneLogHistory.PhoneOrFAX = "FAX"
59340             frmPhoneLogHistory.SampleID = txtSampleID
59350             frmPhoneLogHistory.Show 1
59360         End If
59370         Exit Sub
59380     End If

59390     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
59400         Exit Sub
59410     End If

59420     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

59430     SaveDemographics 1

          Dim Gx As New GP
59440     Gx.LoadName cmbGP
59450     If Gx.FAX <> "" Then
59460         FaxNumber = Gx.FAX
59470     End If
59480     If FaxNumber = "" Then
59490         FaxNumber = IsFaxable("Wards", cmbGP)
59500     End If

          Dim GXs As New GPs
59510     GXs.LoadListFaxNumbers
59520     n = -1
59530     For Each Gx In GXs
59540         n = n + 1
59550         ReDim Preserve strFaxNumbers(0 To n)
59560         strFaxNumbers(n) = Gx.FAX
59570     Next

59580     Set f = New fcdrDBox
59590     With f
59600         .Default = FaxNumber
59610         .ListOrCombo = "List"
59620         .Options = strFaxNumbers
59630         .Prompt = cmbGP & vbCrLf & "Select FAX Number"
59640         .Show 1
59650         FaxNumber = .ReturnValue
59660     End With
59670     Unload f
59680     Set f = Nothing

59690     If FaxNumber = "" Then
59700         iMsg "FAX Cancelled!", vbInformation
59710         Exit Sub
59720     End If

59730     If SSTab1.Tab <> 0 Then
59740         Department = Choose(SSTab1.Tab, "H", "B", "C", "I", "E", "X")
59750         sql = "Select * from PrintPending where " & _
                  "Department = '" & Department & "' " & _
                  "and SampleID = '" & txtSampleID & "' " & _
                  "and UsePrinter = 'FAX'"
59760         Set tb = New Recordset
59770         RecOpenClient 0, tb, sql
59780         If tb.EOF Then
59790             tb.AddNew
59800         End If
59810         tb!SampleID = txtSampleID
59820         tb!Ward = cmbWard
59830         tb!Clinician = cmbClinician
59840         tb!GP = cmbGP
59850         tb!UsePrinter = "FAX"
59860         tb!FaxNumber = FaxNumber

59870         Select Case SSTab1.Tab
                  Case 1: SaveHaematology 1
59880                 sql = "Update HaemResults " & _
                          "Set FAXed = 1 where " & _
                          "SampleID = '" & txtSampleID & "'"
59890                 Cnxn(0).Execute sql
59900                 tb!Department = "H"
59910                 tb!Initiator = HaemValBy
59920             Case 2: SaveBiochemistry 1
59930                 sql = "Update BioResults " & _
                          "Set Faxed = 1 where " & _
                          "SampleID = '" & txtSampleID & "'"
59940                 tb!Department = "B"
59950                 tb!Initiator = BioValBy
59960             Case 3: SaveCoag 1
59970                 sql = "Update CoagResults " & _
                          "Set Faxed = 1 where " & _
                          "SampleID = '" & txtSampleID & "'"
59980                 Cnxn(0).Execute sql
59990                 tb!Department = "C"
60000                 tb!Initiator = CoagValBy
60010         End Select
60020         tb.Update

60030         UpdateFaxLog txtSampleID, Department, FaxNumber

60040         cmdFAX.BackColor = vbYellow
60050         cmdFAX.Caption = "Results Faxed"
60060         cmdFAX.ToolTipText = "Right Click to view Fax Log"

60070     End If

60080     Exit Sub

cmdFAX_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60090     intEL = Erl
60100     strES = Err.Description
60110     LogError "frmEditAll", "cmdFAX_MouseUp", intEL, strES, sql

End Sub

Private Sub cmdGreenTick_Click(Index As Integer)

          Dim y As Integer

60120     Select Case Index
              Case 0    'Bio
60130             gBio.Col = 8
60140             For y = 1 To gBio.Rows - 1
60150                 gBio.row = y
60160                 Set gBio.CellPicture = imgGreenTick.Picture
60170             Next

60180         Case 1:    'Coag
60190             gCoag.Col = 5
60200             For y = 1 To gCoag.Rows - 1
60210                 gCoag.row = y
60220                 Set gCoag.CellPicture = imgGreenTick.Picture
60230             Next

60240     End Select

End Sub

Private Sub cmdOrderExt_Click(Index As Integer)

          Dim frm As New frmAddToTests

10        On Error GoTo cmdOrderExt_Click_Error

          '20     If UserHasAuthority(UserMemberOf, "OrderScreenExternal") = False Then
          '30        iMsg "You do not have authority to Order External test" & vbCrLf & "Please contact system administrator"
          '40        Exit Sub
          '50    End If

20        If Val(txtSampleID) = 0 Then
30            Exit Sub
40        End If

50        If txtSurName = "" And txtDoB = "" Then
60            iMsg "Please provide Surname and DoB first", vbInformation
70            Exit Sub
80        End If

90        SaveDemographics 0
100       frm.SampleID = Format$(Val(txtSampleID))
110       frm.Sex = txtSex
120       If IsDate(tSampleTime) Then
130           frm.SampleDateTime = Format$(dtSampleDate, "dd/MMM/yyyy") & " " & tSampleTime
140       Else
150           frm.SampleDateTime = Format$(dtSampleDate, "dd/MMM/yyyy") & " " & "00:01"
160       End If
170       frm.ClinicalDetails = cClDetails.Text
180       frm.Show 1
190       DoEvents
200       DoEvents
210       Unload frm
220       Set frm = Nothing

230       LoadExt


360       Exit Sub

cmdOrderExt_Click_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "frmEditAll", "cmdOrderExt_Click", intEL, strES

End Sub


Private Sub bDoB_Click()

60520     pBar = 0

60530     LogEvent "DoB Search Click", "frmEditAll", "bDoB_Click"
60540     With frmPatHistory
              '  If HospName(0) = "Monaghan" And SSTab1.Tab = 0 Then
              '    .optBoth = True
              '  Else
60550         If SSTab1.Tab = 0 Then
                  'only for demographic tab
60560             .chkShort = 1
60570         Else
60580             .chkShort = 0
60590         End If
60600         .optHistoric = True
              '  End If
60610         .oFor(2) = True
60620         .lblDept = ""
60630         .txtName = txtDoB
60640         .FromEdit = True
60650         .EditScreen = Me
60660         .bsearch = True

60670         If Not .NoPreviousDetails Then
60680             .Show 1
60690             If txtLabNo = "" Then
60700                 If FormLoaded Then txtLabNo = ""   'Val(FndMaxID("demographics", "LabNo", ""))
60710             End If
60720         Else
60730             FlashNoPrevious Me
60740         End If
60750     End With


60760     CheckCC

End Sub

Private Sub bHaemGraphs_Click()

60770     frmHaemGraphs.SampleID = txtSampleID
60780     frmHaemGraphs.Show 1

End Sub

Private Sub cmdHistory_Click()

60790     pBar = 0

60800     Select Case SSTab1.Tab
              Case 1:
60810             With frmFullHaem
60820                 .lblChart = txtChart
60830                 .lblDoB = txtDoB
60840                 .lblName = Trim$(txtSurName & " " & txtForeName)
60850                 .Show 1
60860             End With
60870         Case 2:
60880             With frmFullHistory
60890                 .Dept = "Bio"
60900                 .lblSex = txtSex
60910                 .lblChart = txtChart
60920                 .lblDoB = txtDoB
60930                 .lblName = Trim$(txtSurName & " " & txtForeName)
60940                 .Show 1
60950             End With
60960         Case 3:
60970             With frmFullHistory
60980                 .Dept = "Coag"
60990                 .lblSex = txtSex
61000                 .lblChart = txtChart
61010                 .lblDoB = txtDoB
61020                 .lblName = Trim$(txtSurName & " " & txtForeName)
61030                 .Show 1
61040             End With
61050     End Select


End Sub



Private Sub cmdOrderBio_Click()

61060     On Error GoTo ErrorHandler

61070     pBar = 0

61080     If UserHasAuthority(UserMemberOf, "OrderScreen") = False Then
61090         iMsg "You do not have authority to Order test" & vbCrLf & "Please contact system administrator"
61100         Exit Sub
61110     End If

61120     With frmNewOrder
61130         .FromEdit = True
61140         .SampleID = txtSampleID
61150         .Chart = txtChart
61160         .Show 1


61170         If IsSampleFasting(txtSampleID) Then
61180             lRandom = "Fasting Sample"
61190         Else
61200             lRandom = "Random Sample"
61210         End If
61220     End With

61230     LoadOutstanding
61240     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub
Function IsSampleFasting(SampleID As String) As Boolean
          Dim sql As String
          Dim tb As Recordset
          Dim Fasting As Boolean

61250     On Error GoTo IsSampleFasting_Error

61260     sql = "SELECT Fasting FROM demographics WHERE " & _
              "SampleID = '" & SampleID & "'"
61270     Set tb = New Recordset
61280     RecOpenClient 0, tb, sql
61290     If tb.EOF Then
61300         IsSampleFasting = False
61310     Else
61320         IsSampleFasting = tb!Fasting
61330     End If


61340     Exit Function

IsSampleFasting_Error:
          Dim strES As String
          Dim intEL As Integer

61350     intEL = Erl
61360     strES = Err.Description
61370     LogError "frmEditAll", "IsSampleFasting", intEL, strES, sql

End Function


Private Sub bOrderTests_Click()

          Dim f As Form

61380     On Error GoTo bOrderTests_Click_Error

61390     pBar = 0

61400     If UserHasAuthority(UserMemberOf, "OrderScreen") = False Then
61410         iMsg "You do not have authority to Order Test" & vbCrLf & "Please contact system administrator"
61420         Exit Sub
61430     End If

61440     SaveDemographics 0
61450     Set f = New frmNewOrder

61460     With f
61470         .FromEdit = True
61480         .SampleID = txtSampleID
61490         .Chart = txtChart
61500         .Show 1
61510         If Not .Cancelled And sysOptAutoScrollAfterOrder(0) Then
61520             cmdSaveDemographics_Click
61530         End If
61540     End With

61550     Set f = Nothing
61560     Debug.Print "bordertests_click"

61570     LoadAllDetails

61580     Exit Sub

bOrderTests_Click_Error:

          Dim strES As String
          Dim intEL As Integer

61590     intEL = Erl
61600     strES = Err.Description
61610     LogError "frmEditAll", "bOrderTests_Click", intEL, strES

End Sub


Private Sub bPrint_Click()

          Dim tb As Recordset
          Dim sql As String

61620     On Error GoTo bPrint_Click_Error
          'Zyam 15-06-24
61630     cmdUpDown(0).Enabled = False
61640     cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

61650     pBar = 0

61660     If UserHasAuthority(UserMemberOf, Choose(SSTab1.Tab, "Haem", "Bio", "Coag", "Imm", "End", "Ext") & "Print") = False Then
61670         iMsg "You do not have authority to print" & vbCrLf & "Please contact system administrator"
61680         Exit Sub
61690     End If

61700     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
61710         Exit Sub
61720     End If

61730     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

61740     SaveDemographics 1

61750     If SavePrintInhibit() Then

61760         If SSTab1.Tab <> 0 Then
61770             LogTimeOfPrinting txtSampleID, Choose(SSTab1.Tab, "H", "B", "C", "I", "E", "X")
61780             sql = "Select * from PrintPending where " & _
                      "Department = '" & Choose(SSTab1.Tab, "H", "B", "C", "I", "E", "X") & "' " & _
                      "and SampleID = '" & txtSampleID & "'"
61790             Set tb = New Recordset
61800             RecOpenClient 0, tb, sql
61810             If tb.EOF Then
61820                 tb.AddNew
61830             End If
61840             tb!SampleID = txtSampleID
61850             tb!PrintOnCondition = 3
61860             tb!Ward = cmbWard
61870             tb!Clinician = cmbClinician
61880             tb!GP = cmbGP
61890             Select Case SSTab1.Tab
                      Case 1: SaveHaematology 1
61900                     tb!Department = "H"
61910                     tb!Initiator = HaemValBy
61920                 Case 2: SaveBiochemistry 1
61930                     tb!Department = "B"
61940                     tb!Initiator = BioValBy
61950                 Case 3: SaveCoag 1
61960                     tb!Department = "D"  'Print All: Changed from "C"
61970                     tb!Initiator = CoagValBy
61980                 Case 4: tb!Department = "I"
61990                     tb!Initiator = UserName
62000                 Case 5: tb!Department = "E"
62010                     tb!Initiator = UserName
62020                 Case 6: SaveExtern
62030                     tb!Department = "X"
62040                     tb!Initiator = UserName
62050             End Select
62060             tb!UsePrinter = pPrintToPrinter
62070             tb.Update
62080         End If
62090     End If

          'txtSampleID = Format$(Val(txtSampleID) + 1)
62100     Debug.Print "bprint_click"
62110     LoadAllDetails
          'Zyam 15-06-24
62120     cmdUpDown(0).Enabled = True
62130     cmdUpDown(1).Enabled = True
          'Zyam 15-06-24

62140     Exit Sub

bPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

62150     intEL = Erl
62160     strES = Err.Description
62170     LogError "frmEditAll", "bPrint_Click", intEL, strES, sql

End Sub

Private Sub SaveDemographics(ByVal Validate As Integer)
          '-1 validate
          '0 un-validate
          '+1 dont change

          Dim sql As String
          Dim tb As Recordset
          Dim Hosp As String
          Dim NewLabNumber As String

62180     On Error GoTo SaveDemographics_Error

62190     txtSampleID = Format(Val(txtSampleID))
62200     If Val(txtSampleID) = 0 Then Exit Sub
          '+++ Junaid 21-05-2024
62210     Call SaveAuditDemo(Trim(txtSampleID.Text))
          '--- Junaid
62220     NewLabNumber = DemographicsUniLabNoSelect(txtSurName & " " & txtForeName, txtDoB, txtSex, txtChart, txtLabNo)   'txtLabNo.Text

62230     SaveComments

62240     If Trim$(tSampleTime) <> "__:__" Then
62250         If Not IsDate(tSampleTime) Then
62260             iMsg "Invalid Time", vbExclamation
62270             Exit Sub
62280         End If
62290     End If

62300     If InStr(UCase$(lblChartNumber), "CAVAN") Then
62310         Hosp = "Cavan"
62320     ElseIf InStr(UCase$(lblChartNumber), "MONAGHAN") Then
62330         Hosp = "Monaghan"
62340     Else
62350         Hosp = HospName(0)
62360     End If

62370     sql = "Select * from Demographics where " & _
              "SampleID = '" & txtSampleID & "'"

62380     Set tb = New Recordset
62390     RecOpenClient 0, tb, sql
62400     If tb.EOF Then
62410         tb.AddNew
62420         If lRandom = "Fasting Sample" Then
62430             tb!Fasting = 1
62440         Else
62450             tb!Fasting = 0
62460         End If
62470         If chkFasting.Value = vbChecked Then
62480             tb!Fasting = 1
62490         Else
62500             tb!Fasting = 0
62510         End If
62520         tb!FAXed = 0
62530     Else

62540         If Trim$(tb!PatName & "") <> "" And _
                  Trim$(UCase$(tb!PatName & "")) <> Trim$(UCase$(txtSurName & " " & txtForeName)) Then
62550             If FlagMessage("Name", tb!PatName, txtSurName & " " & txtForeName, txtSampleID) Then
62560                 txtSurName = SurName(tb!PatName & "")
62570                 txtForeName = ForeName(tb!PatName & "")
62580             End If
62590         End If
62600         If Not IsNull(tb!DoB) Then
62610             If Format(tb!DoB, "dd/mm/yyyy") <> Format(txtDoB, "dd/mm/yyyy") Then
62620                 If FlagMessage("DoB", tb!DoB, txtDoB, txtSampleID) Then
62630                     txtDoB = Format(tb!DoB, "dd/mm/yyyy")
62640                 End If
62650             End If
62660         End If
62670         If Trim$(tb!Chart & "") <> "" And Trim$(UCase$(tb!Chart & "")) <> Trim$(UCase$(txtChart)) Then
62680             If FlagMessage("Chart", tb!Chart, txtChart, txtSampleID) Then
62690                 txtChart = tb!Chart & ""
62700             End If
62710         End If
62720         If Trim$(tb!Ward & "") <> "" And Trim$(UCase$(tb!Ward & "")) <> Trim$(UCase$(cmbWard)) Then
62730             If FlagMessage("Ward", tb!Ward, cmbWard, txtSampleID) Then
62740                 cmbWard = tb!Ward & ""
62750             End If
62760         End If
62770         If Trim$(tb!Clinician & "") <> "" And Trim$(UCase$(tb!Clinician & "")) <> Trim$(UCase$(cmbClinician)) Then
62780             If FlagMessage("Clinician", tb!Clinician, cmbClinician, txtSampleID) Then
62790                 cmbClinician = tb!Clinician & ""
62800             End If
62810         End If
62820     End If


62830     tb!RooH = cRooH(0)
62840     tb!AandE = ""
62850     tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
62860     If IsDate(tSampleTime) Then
62870         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
62880     Else
62890         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
62900     End If
62910     If IsDate(tRecTime) Then
62920         tb!RecDate = Format$(dtRecDate & " " & tRecTime, "dd/mmm/yyyy hh:mm")
62930     Else
62940         tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy")
62950     End If
62960     tb!SampleID = txtSampleID
62970     tb!Chart = txtChart
62980     tb!PatName = Trim$(txtSurName & " " & txtForeName)
62990     tb!SurName = txtSurName & ""
63000     tb!ForeName = txtForeName & ""
63010     If IsDate(txtDoB) Then
63020         tb!DoB = Format$(txtDoB, "dd/mmm/yyyy")
63030     Else
63040         tb!DoB = Null
63050     End If
63060     tb!Age = txtAge
63070     tb!Sex = Left$(txtSex, 1)
63080     tb!Addr0 = txtAddress(0)
63090     tb!Addr1 = txtAddress(1)
63100     If txtSampleID <> "" And (UCase(cmbHospital) <> UCase(HospName(0))) Then
63110         tb!ExtSampleID = txtExtSampleID
63120     End If
63130     tb!Ward = Left$(cmbWard, 50)
63140     tb!Clinician = Left$(cmbClinician, 50)
63150     tb!GP = Left$(cmbGP, 50)
63160     tb!ClDetails = Left$(cClDetails, 300)
63170     tb!Hospital = Hosp
63180     tb!RecordDateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
63190     tb!Operator = Left$(UserName, 20)
63200     If frmOptUrgent Then
63210         If chkUrgent.Value = 1 Then tb!Urgent = 1 Else tb!Urgent = 0
63220     End If
63230     If chkUrgent.Value = 1 Then tb!Urgent = 1 Else tb!Urgent = 0
          '-1 validate
          '0 un-validate
          '+1 dont change
63240     If Validate < 1 Then
63250         tb!Valid = Validate
63260     End If

63270     tb!LabNo = NewLabNumber
63280     tb.Update
63290     Call SaveNotes

63300     LabNoUpdatePrvData txtChart, Trim$(UCase$(AddTicks(txtSurName & " " & txtForeName))), txtDoB, Left$(txtSex, 1), txtLabNo

63310     LogTimeOfPrinting txtSampleID, "D"

63320     Screen.MousePointer = 0

63330     Exit Sub

SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

63340     intEL = Erl
63350     strES = Err.Description
63360     LogError "frmEditAll", "SaveDemographics", intEL, strES, sql

End Sub

Private Sub SaveNotes()
63370     On Error GoTo SaveNotes_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          
63380     sql = "Select RequestID from ocmRequestDetails Where SampleID = '" & txtSampleID.Text & "'"
63390     Set tb = New Recordset
63400     RecOpenServer 0, tb, sql
63410     If Not tb Is Nothing Then
63420         If Not tb.EOF Then
63430             sql = "Update ocmRequest Set Notes = '" & txtNote.Text & "' Where RequestID = '" & tb!RequestID & "'"
63440             Cnxn(0).Execute sql
63450         End If
63460     End If

63470     Exit Sub

SaveNotes_Error:
          
          Dim strES As String
          Dim intEL As Integer

63480     intEL = Erl
63490     strES = Err.Description
63500     LogError "frmEditAll", "SaveDemographics", intEL, strES, sql
End Sub

Private Sub SaveRejectedSample()

          Dim sql As String
          Dim tb As Recordset



63510     On Error GoTo SaveRejectedSample_Error

63520     txtSampleID = Format(Val(txtSampleID))
63530     If Val(txtSampleID) = 0 Then Exit Sub
63540     If chkBioReject.Value = 1 And chkBioReject.Enabled Then
63550         sql = "  INSERT INTO BioResults " & _
                  "  (SampleID, Code, Result, Operator ,Valid, Printed, RunTime, RunDate, " & _
                  "  Analyser, Healthlink ) VALUES " & _
                  "  (" & txtSampleID & ", 'REJ', 'xxx','" & UserCode & "'  ,1, 0, " & Format$(Now, "'dd/mmm/yyyy hh:mm:ss'") & ", " & Format$(Now, "'dd/mmm/yyyy'") & ", " & _
                  "  'Manual', 0) "


63560         Cnxn(0).Execute sql
63570     End If
63580     If chkCoagReject.Value = 1 And chkCoagReject.Enabled Then '+++ Junaid "Remove Operator Field from the query" 04-12-2023
63590         sql = "  INSERT INTO CoagResults " & _
                  "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
                  "   Analyser, Healthlink) VALUES " & _
                  "  (" & txtSampleID & ", 'REJ', 'xxx', 1, 0, " & Format$(Now, "'dd/mmm/yyyy hh:mm:ss'") & ", " & Format$(Now, "'dd/mmm/yyyy'") & ", " & _
                  "   'Manual', 0) "
              '--- Junaid

63600         Cnxn(0).Execute sql
63610     End If

63620     If chkHaemReject.Value = 1 And chkHaemReject.Enabled Then
63630         sql = "  INSERT INTO Observations " & _
                  "  (SampleID, Discipline, Comment, UserName, DateTimeOfRecord) VALUES " & _
                  "  (" & txtSampleID & ", 'Haematology', 'Haematology sample rejected', '" & UserCode & "', " & Format$(Now, "'dd/mmm/yyyy hh:mm:ss'") & ") "


63640         Cnxn(0).Execute sql
63650     End If
63660     If ChkExtReject.Value = 1 And ChkExtReject.Enabled Then
63670         sql = "  INSERT INTO ExtResults " & _
                  "  (SampleID, Analyte, result, Date ) VALUES " & _
                  "  (" & txtSampleID & ", 'REJ', 'xxx'," & Format$(Now, "'dd/mmm/yyyy hh:mm:ss'") & ") "
63680         Cnxn(0).Execute sql

63690         sql = "  INSERT INTO BioResults " & _
                  "  (SampleID, Code, Result, Operator, Valid, Printed, RunTime, RunDate, " & _
                  "  SampleType, Analyser, Healthlink ) VALUES " & _
                  "  (" & txtSampleID & ", 'REJEX', 'xxx','" & UserCode & "', 1, 0, " & Format$(Now, "'dd/mmm/yyyy hh:mm:ss'") & ", " & Format$(Now, "'dd/mmm/yyyy'") & ", " & _
                  "  's', 'Biomnis', 0) "
63700         Cnxn(0).Execute sql


63710     End If
63720     LoadRejectedSample

63730     Exit Sub

SaveRejectedSample_Error:
          Dim strES As String
          Dim intEL As Integer

63740     intEL = Erl
63750     strES = Err.Description
63760     LogError "frmEditAll", "SaveRejectedSample", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ValidateDemo
' Author    : Masood
' Date      : 15/Mar/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ValidateDemo()
63770     On Error GoTo ValidateDemo_Error

          Dim DVs As New DemogValidations
          Dim DV As New DemogValidation

63780     Set DV = New DemogValidation
63790     DV.SampleID = txtSampleID
63800     DV.EnteredBy = UserName
63810     DV.EnteredDateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
63820     DV.ValidatedBy = UserName
63830     DVs.Add DV
63840     DVs.Save DV


63850     Exit Sub


ValidateDemo_Error:

          Dim strES As String
          Dim intEL As Integer

63860     intEL = Erl
63870     strES = Err.Description
63880     LogError "frmEditAll", "ValidateDemo", intEL, strES
End Sub

Private Sub SaveHaematology(ByVal Validate As Boolean)

          Dim tb As Recordset
          Dim sql As String

63890     On Error GoTo SaveHaematology_Error

63900     txtSampleID = Format(Val(txtSampleID))
63910     If Val(txtSampleID) = 0 Then Exit Sub



          '40    If Trim$(tRBC & tHgb & tMCV & tHct & tMCH & tMCHC & tPlt & tMPV & _
          '               tWBC & tLymA & tLymP & tNeutA & tNeutP & tMonoA & tMonoP & _
          '               tEosA & tEosP & tBasA & tBasP & tESR & tRetA & tRetP & _
          '               tMonospot & tWarfarin & txtHaemComment & _
          '               lblMalaria & lblSickledex & lblRA) <> "" Or _
          '         cESR Or cRetics Or cMonospot Or chkMalaria Or chkRA Or chkSickledex Or _
          '         cFilm = 1 Then
          '50    End If

          '40    If Trim$(tRBC & tHgb & tMCV & tHct & tMCH & tMCHC & tPlt & tMPV & _
          '               tWBC & tLymA & tLymP & tNeutA & tNeutP & tMonoA & tMonoP & _
          '               tEosA & tEosP & tBasA & tBasP & tESR & tRetA & tRetP & _
          '               tMonospot & tWarfarin & txtHaemComment & txtFilmComment & _
          '               lblMalaria & lblSickledex & lblRA) = "" And _
          '               cESR = 0 And cRetics = 0 And cMonospot = 0 And _
          '               chkMalaria = 0 And chkRA = 0 And _
          '               chkSickledex = 0 And cFilm = 0 Then
          '         Exit Sub
          '50    End If

63920     If Trim$(tWBC & tLymA & tLymP & tNeutA & tNeutP & tMonoA & tMonoP & _
              tEosA & tEosP & tBasA & tBasP) = "" Then
63930         If iMsg("Save with no WBC/Diff results?", vbYesNo + vbQuestion, , vbRed) = vbNo Then
63940             Exit Sub
63950         End If
63960     End If

63970     sql = "Select * from HaemResults where " & _
              "SampleID = '" & txtSampleID & "'"
63980     Set tb = New Recordset
63990     RecOpenClient 0, tb, sql
64000     If tb.EOF Then
64010         tb.AddNew
64020         tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
64030         tb!RunDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
64040         tb!SampleID = txtSampleID
64050         tb!FAXed = 0
64060         tb!Printed = 0
64070     End If

64080     tb!rbc = tRBC
64090     CheckIfMustPhone "Haematology", "RBC", tRBC
64100     tb!Hgb = tHgb
64110     CheckIfMustPhone "Haematology", "Hgb", tHgb
64120     tb!MCV = tMCV
64130     CheckIfMustPhone "Haematology", "MCV", tMCV
64140     tb!hct = tHct
64150     CheckIfMustPhone "Haematology", "Hct", tHct
64160     tb!RDWCV = tRDWCV
64170     CheckIfMustPhone "Haematology", "RDWCV", tRDWCV
64180     tb!rdwsd = tRDWSD
64190     CheckIfMustPhone "Haematology", "RDWSD", tRDWSD
64200     tb!mch = tMCH
64210     CheckIfMustPhone "Haematology", "MCH", tMCH
64220     tb!mchc = tMCHC
64230     CheckIfMustPhone "Haematology", "MCHC", tMCHC
64240     tb!plt = tPlt
64250     CheckIfMustPhone "Haematology", "Plt", tPlt
64260     tb!mpv = tMPV
64270     CheckIfMustPhone "Haematology", "MPV", tMPV
64280     tb!WBC = tWBC
64290     CheckIfMustPhone "Haematology", "WBC", tWBC
64300     tb!LymA = tLymA
64310     CheckIfMustPhone "Haematology", "LymA", tLymA
64320     tb!LymP = Left$(tLymP, 5)
64330     CheckIfMustPhone "Haematology", "LymP", tLymP
64340     tb!MonoA = tMonoA
64350     CheckIfMustPhone "Haematology", "MonoA", tMonoA
64360     tb!MonoP = Left$(tMonoP, 5)
64370     CheckIfMustPhone "Haematology", "MonoP", tMonoP
64380     tb!NeutA = tNeutA
64390     CheckIfMustPhone "Haematology", "NeutA", tNeutA
64400     tb!NeutP = Left$(tNeutP, 5)
64410     CheckIfMustPhone "Haematology", "NeutP", tNeutP
64420     tb!EosA = tEosA
64430     CheckIfMustPhone "Haematology", "EosA", tEosA
64440     tb!EosP = Left$(tEosP, 5)
64450     CheckIfMustPhone "Haematology", "EosP", tEosP
64460     tb!BasA = tBasA
64470     CheckIfMustPhone "Haematology", "BasA", tBasA
64480     tb!BasP = Left$(tBasP, 5)
64490     CheckIfMustPhone "Haematology", "BasP", tBasP

64500     tb!nrbcA = Left$(Trim$(tnrbcA), 5)
64510     CheckIfMustPhone "Haematology", "nrbcA", tnrbcA
64520     tb!nrbcP = Left$(Trim$(tnrbcP), 5)
64530     CheckIfMustPhone "Haematology", "nrbcP", tnrbcP

64540     tb!cESR = IIf(cESR = 1, 1, 0)
64550     If tESR = "Pending" Then
64560         tb!ESR = "?"
64570     Else
64580         tb!ESR = Left$(tESR, 3)
64590     End If

64600     tb!cRetics = IIf(cRetics = 1, 1, 0)
64610     tb!RetA = tRetA
64620     tb!RetP = tRetP

64630     tb!cMonospot = IIf(cMonospot = 1, 1, 0)
64640     tb!MonoSpot = Left$(tMonospot, 1)

64650     tb!cMalaria = chkMalaria
64660     tb!Malaria = lblMalaria

64670     tb!cSickledex = chkSickledex
64680     tb!Sickledex = lblSickledex

64690     tb!cRA = chkRA
64700     tb!RA = Left$(lblRA, 1)



64710     tb!ccoag = 0
64720     tb!cFilm = cFilm = 1

64730     tb!Warfarin = tWarfarin

64740     tb!CD3A = txtCD3A.Text
64750     tb!CD3P = txtCD3P.Text
64760     tb!CD4A = txtCD4A.Text
64770     tb!CD4P = txtCD4P.Text
64780     tb!CD8A = txtCD8A.Text
64790     tb!CD8P = txtCD8P.Text
64800     tb!CD48 = txtCD48.Text


          '------------farhan--------------
64810     tb!SignOff = Null
64820     tb!SignOffBy = Null
64830     tb!SignOffDateTime = Null
          '================================

64840     tb!Valid = IIf(Validate = True, 1, 0)

64850     If Validate Then
64860         tb!HealthLink = 0
              'If tb!Operator & "" = "" Then
64870         tb!Operator = UserCode
64880         tb!ValidateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
64890         HaemValBy = UserCode
              'End If
64900     End If

64910     tb.Update

64920     Call SaveHaemResults50(txtSampleID, "ESR", IIf((tESR = "Pending"), "?", Left$(tESR, 3)), "", "")

64930     Call SaveHaemResults50(txtSampleID, "RetA", tRetA, "", "")
64940     Call SaveHaemResults50(txtSampleID, "RetP", tRetP, "", "")

          'Call SaveHaemResults50(txtSampleID, "cMonospot", IIf(cMonospot = 1, 1, 0), "", "")
64950     Call SaveHaemResults50(txtSampleID, "MonoSpot", Left$(tMonospot, 1), "", "")

          'Call SaveHaemResults50(txtSampleID, "cMalaria", chkMalaria, "", "")
64960     Call SaveHaemResults50(txtSampleID, "Malaria", lblMalaria, "", "")

          'Call SaveHaemResults50(txtSampleID, "cSickledex", chkSickledex, "", "")
64970     Call SaveHaemResults50(txtSampleID, "Sickledex", lblSickledex, "", "")

          'Call SaveHaemResults50(txtSampleID, "cRA", chkRA, "", "")
64980     Call SaveHaemResults50(txtSampleID, "RA", lblRA, "", "")


64990     If Validate Then
65000         sql = "UPDATE HaemResults50 " & _
                  "Set ValidateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' WHERE " & _
                  "SampleID = '" & txtSampleID & "' AND ValidateTime IS NULL "
65010         Cnxn(0).Execute sql
65020     End If
65030     If cFilm.Value = 1 Then
65040         sql = "SELECT * FROM HaeRequests WHERE " & _
                  "SampleID = '" & txtSampleID & "'"
65050         Set tb = New Recordset
65060         RecOpenServer 0, tb, sql
65070         If tb.EOF Then
65080             tb.AddNew
65090             tb!SampleID = Val(txtSampleID)
65100             tb!Code = "SP"
65110             tb!Programmed = 0
65120             tb!SampleType = "Blood EDTA"
65130             tb!Analyser = "IPU"
65140             tb!UserName = UserName
65150             tb.Update
65160             tb.Close
65170         End If
65180     End If
65190     Screen.MousePointer = 0

65200     Exit Sub

SaveHaematology_Error:

          Dim strES As String
          Dim intEL As Integer

65210     intEL = Erl
65220     strES = Err.Description
65230     LogError "frmEditAll", "SaveHaematology", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveHaemResults50
' Author    : Masood
' Date      : 29/Jul/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveHaemResults50(SampleID As String, Code As String, Result As String, Flags As String, Units As String)

65240     On Error GoTo SaveHaemResults50_Error

          Dim sql As String

65250     sql = "IF EXISTS (SELECT * FROM HaemResults50 " & _
              "           WHERE SampleID = '" & SampleID & "' " & _
              "           AND Code = '" & Code & "' " & _
              "        ) " & _
              "    UPDATE HaemResults50 " & _
              "    SET Code = '" & Code & "' " & _
              "    , Result = '" & Result & "' " & _
              "    , Flags = '" & Flags & "' " & _
              "    , Units = '" & Units & "' " & _
              "    WHERE SampleID = '" & SampleID & "' " & _
              "    AND Code = '" & Code & "' " & _
              "ELSE " & _
              "    INSERT INTO HaemResults50 (SampleID, Code, Result, Flags, Units,DefIndex,DateTimeOfRecord) " & _
              "    VALUES ('" & SampleID & "', " & _
              "            '" & Code & "', " & _
              "            '" & Result & "', " & _
              "            '" & Flags & "'" & _
              "            ,'" & Units & "'" & _
              " ,0" & _
              " , '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'" & _
              " )"
65260     Cnxn(0).Execute sql

65270     Exit Sub


SaveHaemResults50_Error:

          Dim strES As String
          Dim intEL As Integer

65280     intEL = Erl
65290     strES = Err.Description
65300     LogError "frmEditAll", "SaveHaemResults50", intEL, strES, sql
End Sub

Private Sub SaveCoag(ByVal Validate As Boolean)

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim Code As String

65310     On Error GoTo SaveCoag_Error

65320     txtSampleID = Format(Val(txtSampleID))
65330     If Val(txtSampleID) = 0 Then Exit Sub
65340     If gCoag.Rows = 2 And gCoag.TextMatrix(1, 0) = "" Then Exit Sub

65350     For n = 1 To gCoag.Rows - 1

65360         CheckIfMustPhone "Coagulation", gCoag.TextMatrix(n, 0), gCoag.TextMatrix(n, 1)

65370         sql = "Select Code from CoagTestDefinitions where " & _
                  "TestName = '" & gCoag.TextMatrix(n, 0) & "'"
65380         Set tb = New Recordset
65390         RecOpenServer 0, tb, sql
65400         If tb.EOF Then Exit Sub

65410         Code = tb!Code & ""

65420         sql = "Select * from CoagResults where " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "and Code = '" & Code & "'"
65430         Set tb = New Recordset
65440         RecOpenClient 0, tb, sql
65450         If tb.EOF Then
65460             tb.AddNew
65470             tb!ValidateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
65480         End If
65490         tb!Code = Code
              '  tb!Flag = gCoag.TextMatrix(n, 3)
              '  tb!OperatorCode = UserCode
65500         tb!Result = gCoag.TextMatrix(n, 1)
65510         tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
65520         tb!RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
65530         tb!SampleID = txtSampleID
10            tb!Printed = IIf(InStr(gCoag.TextMatrix(n, 4), "P"), 1, 0)
              'check validation using the "V" in the grid'
20            If Validate Then
30                tb!Valid = 1
40                If IsNull(tb!ValidateTime) Then
50                    tb!ValidateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
60                End If

                  '         If (InStr(gCoag.TextMatrix(n, 4), "V")) = 0 Then
                  '              tb!ValidateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
                  '         End If
70            Else
80                tb!Valid = IIf(InStr(gCoag.TextMatrix(n, 4), "V"), 1, 0)
                  '        If (InStr(gCoag.TextMatrix(n, 4), "V")) = 0 Then
                  '            tb!ValidateTime = Null
                  '        End If
90            End If

100           tb.Update
110       Next
120       tb.Close

130       If Validate Then
140           sql = "UPDATE CoagResults " & _
                  "Set UserName = '" & UserCode & "', HealthLink = '0' WHERE " & _
                  "SampleID = '" & txtSampleID & "' "
              '                & _
              '                "AND COALESCE(UserName, '') = ''"
150           Cnxn(0).Execute sql
160           CoagValBy = UserCode
170       End If

180       sql = "SELECT * FROM HaemResults WHERE " & _
              "SampleID = '" & txtSampleID & "'"
190       Set tb = New Recordset
200       RecOpenClient 0, tb, sql
210       If Not tb.EOF Then
220           tb!Warfarin = Trim$(tWarfarin)
230           tb.Update
240       End If

250       Exit Sub

SaveCoag_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditAll", "SaveCoag", intEL, strES, sql

End Sub



Private Sub bPrintAll_Click()

          Dim tb As Recordset
          Dim sql As String

290       On Error GoTo bPrintAll_Click_Error
          'Zyam 15-06-24
300       cmdUpDown(0).Enabled = False
310       cmdUpDown(1).Enabled = False
          'Zyam 15-06-24


320       pBar = 0

330       If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
340           Exit Sub
350       End If

360       If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

370       SaveDemographics 1

380       If SSTab1.Tab <> 0 Then
390           sql = "Select * from PrintPending where " & _
                  "Department = 'D' " & _
                  "and SampleID = '" & txtSampleID & "'"
400           Set tb = New Recordset
410           RecOpenClient 0, tb, sql
420           If tb.EOF Then
430               tb.AddNew
440           End If
450           tb!SampleID = txtSampleID
460           tb!Ward = cmbWard
470           tb!Clinician = cmbClinician
480           tb!GP = cmbGP
490           tb!Department = "D"
500           tb!Initiator = UserName
510           tb!UsePrinter = pPrintToPrinter
520           tb.Update
530       End If

540       SaveCoag 1
550       sql = "Update CoagResults " & _
              "Set Valid = 1," & _
              " ValidateTime = '" & Format$(Now, "dd/MMM/yyyy HH:mm:ss") & "' ," & _
              " Printed = 1 where " & _
              "SampleID = '" & txtSampleID & "'"
560       Cnxn(0).Execute sql

          'txtSampleID = Format$(Val(txtSampleID) + 1)
570       Debug.Print "bprintall_click"
580       LoadAllDetails

          'Zyam 15-06-24
590       cmdUpDown(0).Enabled = True
600       cmdUpDown(1).Enabled = True
          'Zyam 15-06-24

610       Exit Sub

bPrintAll_Click_Error:

          Dim strES As String
          Dim intEL As Integer

620       intEL = Erl
630       strES = Err.Description
640       LogError "frmEditAll", "bPrintAll_Click", intEL, strES, sql


End Sub

Private Sub bprintesr_Click()

          Dim tb As Recordset
          Dim sql As String

650       On Error GoTo bprintesr_Click_Error

660       pBar = 0

670       If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
680           Exit Sub
690       End If

700       If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

710       SaveDemographics 1

720       SaveHaematology 1

730       LogTimeOfPrinting txtSampleID, "H"

740       sql = "Select * from PrintPending where " & _
              "Department = 'K' " & _
              "and SampleID = '" & txtSampleID & "'"
750       Set tb = New Recordset
760       RecOpenClient 0, tb, sql
770       If tb.EOF Then
780           tb.AddNew
790       End If
800       tb!SampleID = txtSampleID
810       tb!Ward = cmbWard
820       tb!Clinician = cmbClinician
830       tb!GP = cmbGP
840       tb!Department = "K"
850       tb!Initiator = HaemValBy
860       tb!UsePrinter = pPrintToPrinter
870       tb.Update

880       sql = "Update HaemResults " & _
              "Set Printed = 1 where " & _
              "SampleID = '" & txtSampleID & "'"
890       Cnxn(0).Execute sql

900       Exit Sub

bprintesr_Click_Error:

          Dim strES As String
          Dim intEL As Integer

910       intEL = Erl
920       strES = Err.Description
930       LogError "frmEditAll", "bprintesr_Click", intEL, strES, sql

End Sub

Private Sub bPrintHold_Click()

          Dim tb As Recordset
          Dim sql As String

940       On Error GoTo bPrintHold_Click_Error
          'Zyam 15-06-24
950       cmdUpDown(0).Enabled = False
960       cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

970       pBar = 0

980       If UserHasAuthority(UserMemberOf, Choose(SSTab1.Tab, "Haem", "Bio", "Coag", "Imm", "End", "Ext") & "Print") = False Then
990           iMsg "You do not have authority to print" & vbCrLf & "Please contact system administrator"
1000          Exit Sub
1010      End If

1020      If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
1030          Exit Sub
1040      End If

1050      If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

1060      SaveDemographics 1
1070      If SavePrintInhibit() Then

1080          If SSTab1.Tab <> 0 Then
1090              LogTimeOfPrinting txtSampleID, Choose(SSTab1.Tab, "H", "B", "C", "I", "E", "X")
1100              sql = "SELECT * FROM PrintPending WHERE " & _
                      "Department = '" & Choose(SSTab1.Tab, "H", "B", "C", "I", "E", "X") & "' " & _
                      "AND SampleID = '" & txtSampleID & "'"
1110              Set tb = New Recordset
1120              RecOpenClient 0, tb, sql
1130              If tb.EOF Then
1140                  tb.AddNew
1150              End If
1160              tb!SampleID = txtSampleID
1170              tb!Ward = cmbWard
1180              tb!Clinician = cmbClinician
1190              tb!GP = cmbGP
1200              Select Case SSTab1.Tab
                      Case 1: SaveHaematology 1
1210                      tb!Department = "H"
1220                      tb!Initiator = HaemValBy
1230                  Case 2: SaveBiochemistry 1
1240                      tb!Department = "B"
1250                      tb!Initiator = BioValBy
1260                  Case 3: SaveCoag 1
1270                      tb!Department = "D"    'Print All: Changed from "C"
1280                      tb!Initiator = CoagValBy
1290                  Case 4: tb!Department = "I"
1300                      tb!Initiator = UserName
1310                  Case 5: tb!Department = "E"
1320                      tb!Initiator = UserName
1330                  Case 6: SaveExtern
1340                      tb!Department = "X"
1350                      tb!Initiator = UserName
1360              End Select
1370              tb!UsePrinter = pPrintToPrinter
1380              tb.Update
1390          End If
1400      End If
          'Zyam 15-06-24
1410      cmdUpDown(0).Enabled = True
1420      cmdUpDown(1).Enabled = True
          'Zyam 15-06-24

1430      Exit Sub

bPrintHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1440      intEL = Erl
1450      strES = Err.Description
1460      LogError "frmEditAll", "bPrintHold_Click", intEL, strES, sql

End Sub

Private Sub bPrintINR_Click()

1470      If Trim$(txtChart) = "" Then
1480          iMsg "Enter Chart Number", vbExclamation
1490          Exit Sub
1500      End If

1510      pBar = 0

1520      If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

1530      SaveDemographics 1
1540      SaveCoag 1

1550      With frmINR
1560          .tChart = txtChart
1570          .Ward = cmbWard
1580          .LoadDetails
1590          .Show 1
1600      End With

1610      txtSampleID = Format$(Val(txtSampleID) + 1)
1620      Debug.Print "bprintinr_click"
1630      LoadAllDetails

End Sub

Private Sub cmdSaveHoldDemographics_Click()

1640      pBar = 0

1650      If UserHasAuthority(UserMemberOf, "DemSave") = False Then
1660          iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
1670          Exit Sub
1680      End If

1690      If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
1700          Exit Sub
1710      End If

1720      If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub    'If  tSampleTime, txtDemographicComment, tRecTime, cmbHospital) Then Exit Sub

1730      cmdSaveHoldDemographics.Caption = "Saving"

1740      SaveDemographics 1
1750      SaveRejectedSample
1760      UpdateMRU Me

1770      cmdSaveHoldDemographics.Caption = "Save && &Hold"
1780      cmdSaveHoldDemographics.Enabled = False
1790      cmdSaveDemographics.Enabled = False

End Sub

Private Sub cmdPhone_Click()

1800      With frmPhoneLog
1810          .SampleID = txtSampleID
1820          If cmbGP <> "" Then
1830              .GP = cmbGP
1840              .WardOrGP = "GP"
1850          Else
1860              .GP = cmbWard
1870              .WardOrGP = "Ward"
1880          End If
1890          .Show 1
1900      End With

1910      CheckIfPhoned
1920      LoadComments


End Sub

Private Sub cmdPrintMonospot_Click()

          Dim tb As Recordset
          Dim sql As String

1930      On Error GoTo cmdPrintMonospot_Click_Error

1940      pBar = 0

1950      If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
1960          Exit Sub
1970      End If

1980      If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

1990      SaveDemographics 1

2000      SaveHaematology 1

2010      LogTimeOfPrinting txtSampleID, "H"

2020      sql = "Select * from PrintPending where " & _
              "Department = 'P' " & _
              "and SampleID = '" & txtSampleID & "'"
2030      Set tb = New Recordset
2040      RecOpenClient 0, tb, sql
2050      If tb.EOF Then
2060          tb.AddNew
2070      End If
2080      tb!SampleID = txtSampleID
2090      tb!Ward = cmbWard
2100      tb!Clinician = cmbClinician
2110      tb!GP = cmbGP
2120      tb!Department = "P"
2130      tb!Initiator = HaemValBy
2140      tb!UsePrinter = pPrintToPrinter
2150      tb.Update

2160      sql = "Update HaemResults " & _
              "Set Printed = 1 where " & _
              "SampleID = '" & txtSampleID & "'"
2170      Cnxn(0).Execute sql

2180      Exit Sub

cmdPrintMonospot_Click_Error:

          Dim strES As String
          Dim intEL As Integer

2190      intEL = Erl
2200      strES = Err.Description
2210      LogError "frmEditAll", "cmdPrintMonospot_Click", intEL, strES, sql

End Sub

Private Sub cmdRedCross_Click(Index As Integer)

          Dim y As Integer

2220      Select Case Index
              Case 0    'Bio
2230              gBio.Col = 8
2240              For y = 1 To gBio.Rows - 1
2250                  gBio.row = y
2260                  Set gBio.CellPicture = imgRedCross.Picture
2270              Next

2280          Case 1:    'Coag
2290              gCoag.Col = 5
2300              For y = 1 To gCoag.Rows - 1
2310                  gCoag.row = y
2320                  Set gCoag.CellPicture = imgRedCross.Picture
2330              Next

2340      End Select

End Sub

Private Sub SaveExtern()

          Dim tb As Recordset
          Dim tbD As ADODB.Recordset
          Dim n As Integer
          Dim sql As String
          Dim AnalyteName As String
          Dim NetAcquireTestCode As String

2350      On Error GoTo SaveExtern_Error



2360      txtSampleID = Format(Val(txtSampleID))
2370      If Val(txtSampleID) = 0 Then Exit Sub

2380      For n = 1 To g.Rows - 1
          
2390          If Trim(g.TextMatrix(n, 0)) <> "" Then
              
2400              AnalyteName = g.TextMatrix(n, 0)
2410              sql = "Select * from ExtResults where " & _
                      "sampleid = '" & txtSampleID & "' " & _
                      "and Analyte = '" & AnalyteName & "'"
2420              Set tb = New Recordset
2430              RecOpenServer 0, tb, sql
2440              If tb.EOF Then
2450                  tb.AddNew
2460              End If
2470              tb!SampleID = txtSampleID
2480              tb!Analyte = AnalyteName
2490              tb!Result = g.TextMatrix(n, 1)
2500              tb!Units = g.TextMatrix(n, 3)
2510              tb!SendTo = g.TextMatrix(n, 4)
2520              If IsDate(g.TextMatrix(n, 5)) Then
2530                  tb!Date = Format(g.TextMatrix(n, 5), "dd/mmm/yyyy")
2540              Else
2550                  tb!Date = Null
2560              End If
                  '+++ Junaid 28-06-2022
2570              sql = "Select External_ from ocmRequestDetails Where SampleID = '" & txtSampleID.Text & "' and External_ = '1'"
2580              Set tbD = New Recordset
2590              RecOpenServer 0, tbD, sql
2600              If Not tbD Is Nothing Then
2610                  If Not tbD.EOF Then
2620                      If tbD!External_ = "1" Then
2630                          tb!SentDate = Format(Date, "dd/mmm/yyyy")
2640                      End If
2650                  End If
2660              End If
                  '--- Junaid 28-06-2022
2670              tb.Update
              
2680              NetAcquireTestCode = "SYS001"
2690              If g.TextMatrix(n, 1) <> "" Then
2700                  sql = "Select * from BioResults where " & _
                          "sampleid = '" & txtSampleID & "' " & _
                          "and Code = '" & NetAcquireTestCode & "'"
2710                  Set tb = New Recordset
2720                  RecOpenServer 0, tb, sql
2730                  If tb.EOF Then
2740                      tb.AddNew
2750                  End If
2760                  tb!SampleID = txtSampleID
2770                  tb!Code = NetAcquireTestCode
2780                  tb!Result = "View Scan"
2790                  tb!Valid = 1
2800                  tb!Printed = 1
2810                  If IsDate(g.TextMatrix(n, 5)) Then
2820                      tb!RunTime = Format(g.TextMatrix(n, 5), "dd/mmm/yyyy hh:mm:ss")
2830                      tb!Rundate = Format(g.TextMatrix(n, 5), "dd/mmm/yyyy")
2840                  Else
2850                      tb!RunDateTime = Null
2860                      tb!Rundate = Null
2870                  End If
2880                  tb!Operator = UserCode
2890                  tb!Units = g.TextMatrix(n, 3)
2900                  tb!SampleType = "S"
2910                  tb!Analyser = "Manual"   'should be send to
                  
2920                  tb.Update
2930              End If
              
2940          End If
2950      Next

2960      sql = "Select * from etc where " & _
              "sampleid = '" & txtSampleID & "'"
2970      Set tb = New Recordset
2980      RecOpenServer 0, tb, sql
2990      If tb.EOF Then
3000          tb.AddNew
3010          tb!SampleID = txtSampleID
3020      End If
3030      tb!etc0 = etc(0)
3040      tb!etc1 = etc(1)
3050      tb!etc2 = etc(2)
3060      tb!etc3 = etc(3)
3070      tb!etc4 = etc(4)
3080      tb!etc5 = etc(5)
3090      tb!etc6 = etc(6)
3100      tb!etc7 = etc(7)
3110      tb!etc8 = etc(8)
3120      tb.Update

3130      Exit Sub

SaveExtern_Error:

          Dim strES As String
          Dim intEL As Integer

3140      intEL = Erl
3150      strES = Err.Description
3160      LogError "frmEditAll", "SaveExtern", intEL, strES, sql


End Sub

'---------------------------------------------------------------------------------------
' Procedure : bsearch_Click
' Author    : XPMUser
' Date      : 23/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : bsearch_Click
' Author    : XPMUser
' Date      : 23/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bsearch_Click()

3170      On Error GoTo bsearch_Click_Error


3180      pBar = 0

3190      LogEvent "Name Search Click", "frmEditAll", "bsearch_Click"


3200      With frmPatHistory
3210          If SSTab1.Tab = 0 Then
                  'only for demographic tab
3220              .chkShort = 1
3230          Else
3240              .chkShort = 0
3250          End If
3260          If HospName(0) = "Monaghan" And SSTab1.Tab = 0 Then
3270              .optBoth = True
3280          Else
3290              .optHistoric = True
3300          End If
3310          .lblDept = ""
3320          .oFor(0) = True
3330          .txtName = Trim$(txtSurName & " " & txtForeName)
3340          .FromEdit = True

3350          .EditScreen = Me
3360          .bsearch = True
3370          If Not .NoPreviousDetails Then
3380              .Show 1
3390          Else
3400              FlashNoPrevious Me
3410          End If
3420      End With
3430      CheckCC
3440      LabNoUpdatePrvColor

3450      Exit Sub


bsearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer

3460      intEL = Erl
3470      strES = Err.Description
3480      LogError "frmEditAll", "bsearch_Click", intEL, strES

End Sub

Private Sub bViewBB_Click()

3490      pBar = 0

3500      If Trim$(txtChart) <> "" Then
3510          frmViewBB.lchart = txtChart
3520          frmViewBB.Show 1
3530      End If

End Sub

Private Sub ShowWait(ByVal Show As Boolean)
3540      On Error GoTo ShowWait_Error

3550      If Show Then
3560          fraWait.Left = 5000
3570      Else
3580          fraWait.Left = 15000
3590      End If
3600      fraWait.Top = 5000

3610      Exit Sub
ShowWait_Error:
         
3620      LogError "frmEditAll", "ShowWait", Erl, Err.Description


End Sub


Sub LoadAllDetails()

3630      On Error GoTo LoadAllDetails_Error
          '    txtLabNo = Val(FndMaxID("demographics", "LabNo", ""))

3640      ShowWait True

3650      ClearLabNoSelection
3660      cmdViewReports.Visible = False

3670      HaemLoaded = False
3680      BioLoaded = False
3690      CoagLoaded = False
3700      ExtLoaded = False

3710      cmbAdd = ""
3720      cmbUnits = ""

3730      ClearDemographics
3740      ClearHaematologyResults
3750      ClearCoagulation

3760      HaemValBy = ""
3770      BioValBy = ""
3780      CoagValBy = ""
3790      ExtValBy = ""
3800      chkBioReject.Value = 0
3810      chkCoagReject.Value = 0
3820      chkHaemReject.Value = 0
3830      ChkExtReject.Value = 0
3840      chkBioReject.BackColor = &H8000000F
3850      chkCoagReject.BackColor = &H8000000F
3860      chkHaemReject.BackColor = &H8000000F
3870      ChkExtReject.BackColor = &H8000000F
3880      chkBioReject.Enabled = True
3890      chkCoagReject.Enabled = True
3900      chkHaemReject.Enabled = True
3910      ChkExtReject.Enabled = True


3920      SSTab1.TabCaption(1) = "Haematology"
3930      SSTab1.TabCaption(2) = "Biochemistry"
3940      SSTab1.TabCaption(3) = "Coagulation"
3950      SSTab1.TabCaption(6) = "External"


3960      LoadDemographics
3970      LoadRejectedSample
3980      DoEvents
3990      LockDemographics Me, False
4000      If Trim$(txtSurName) <> "" Or Trim$(txtForeName) <> "" Or Trim$(txtChart) <> "" Then
4010          LockDemographics Me, True
4020      End If

4030      CheckDepartments
4040      CheckCC

4050      Debug.Print m_StartInDepartment

4060      Select Case m_StartInDepartment
              Case "H": SSTab1.Tab = 1
4070          Case "B": SSTab1.Tab = 2
4080          Case "C": SSTab1.Tab = 3
4090      End Select

4100      Select Case SSTab1.Tab
              Case 0:
4110          Case 1:
4120              If SSTab1.TabVisible(1) = True Then

4130                  LoadHaematology
4140                  HaemLoaded = True
4150              End If
4160          Case 2:
4170              If SSTab1.TabVisible(2) = True Then
4180                  LoadBiochemistry
4190                  BioLoaded = True
4200              End If
4210          Case 3:
4220              If SSTab1.TabVisible(3) = True Then

4230                  LoadCoagulation
4240                  CoagLoaded = True
4250              End If
4260          Case 6:
4270              If SSTab1.TabVisible(6) = True Then

4280                  LoadExt
4290                  ExtLoaded = True
4300              End If
4310      End Select
4320      DoEvents
4330      LoadComments
4340      Call ShowNotes
4350      Call ShowSamples
4360      Call ShowQuestions
          'SetViewHistory
4370      lblRequestID.Caption = GetRequestID(txtSampleID.Text)

4380      ShowHistory (SSTab1.Tab)

4390      EnableCopyFrom

4400      CheckIfPhoned
4410      CheckIfFaxed
        
        
          'Abubaker+++ 05/10/2023 (locked the combo box to stop manual entry)
          'cmbWard.Locked = True


          '740   CheckIfWardClinicianOK

4420      CheckTag

4430      If cmdViewScan.Visible = True And CheckScanViewLog(txtSampleID, SSTab1.TabCaption(SSTab1.Tab)) = False Then
4440          frmViewScan.CallerDepartment = SSTab1.TabCaption(SSTab1.Tab)
4450          frmViewScan.SampleID = txtSampleID
4460          frmViewScan.txtSampleID = txtSampleID
4470          frmViewScan.Show 1
4480      End If
4490      MatchingDemoLoaded = False

4500      If IsNotepadExists(txtSampleID, "") = True Then
4510          cmdPatientNotePad.BackColor = vbYellow
4520      Else
4530          cmdPatientNotePad.BackColor = &H8000000F
4540      End If
4550      cmdValidateDemographics.Visible = False
4560      ShowWait False

4570      Exit Sub

LoadAllDetails_Error:

4580      ShowWait False
          Dim strES As String
          Dim intEL As Integer

4590      intEL = Erl
4600      strES = Err.Description
4610      LogError "frmEditAll", "LoadAllDetails", intEL, strES

End Sub

Private Sub ShowNotes()
4620      On Error GoTo ShowNotes_Error

          Dim sql As String
          Dim l_RequestID As String
          Dim tb As ADODB.Recordset
          Dim tbR As ADODB.Recordset
          Dim tbAB As ADODB.Recordset
          Dim tbIAB As ADODB.Recordset
          
4630      sql = "Select IsNULL(RequestID,'') RequestID from ocmRequestDetails Where SampleID = '" & txtSampleID.Text & "'"
4640      Set tbR = New Recordset
4650      RecOpenServer 0, tbR, sql
4660      If Not tbR Is Nothing Then
4670          If Not tbR.EOF Then
4680              l_RequestID = ConvertNull(tbR!RequestID, "")
4690          End If
4700      End If
          
4710      txtNote.Text = ""
4720      sql = "Select Notes from ocmRequest  "
4730      sql = sql & "Where RequestID = '" & l_RequestID & "'"
4740      Set tb = New Recordset
4750      RecOpenServer 0, tb, sql
4760      If Not tb Is Nothing Then
4770          If Not tb.EOF Then
4780              txtNote.Text = ConvertNull(tb!Notes, "")
4790          End If
4800      End If
          
4810      txtAddress(2).Text = ""
4820      sql = "Select IsNULL(antibiotics,'') antibiotics from ocmRequest  "
4830      sql = sql & "Where RequestID = '" & l_RequestID & "'"
4840      Set tbAB = New Recordset
4850      RecOpenServer 0, tbAB, sql
4860      If Not tbAB Is Nothing Then
4870          If Not tbAB.EOF Then
4880              txtAddress(2).Text = ConvertNull(tbAB!Antibiotics, "")
4890          End If
4900      End If
          
4910      txtAddress(3).Text = ""
4920      sql = "Select IsNULL(intendedantibiotics,'') intendedantibiotics from ocmRequest  "
4930      sql = sql & "Where RequestID = '" & l_RequestID & "'"
4940      Set tbIAB = New Recordset
4950      RecOpenServer 0, tbIAB, sql
4960      If Not tbIAB Is Nothing Then
4970          If Not tbAB.EOF Then
4980              txtAddress(3).Text = ConvertNull(tbIAB!IntendedAntibiotics, "")
4990          End If
5000      End If
          
5010      Exit Sub

ShowNotes_Error:
          
          Dim strES As String
          Dim intEL As Integer

5020      intEL = Erl
5030      strES = Err.Description
5040      LogError "frmEditAll", "ShowNotes", intEL, strES, sql
End Sub

Private Sub bcancel_Click()

5050      pBar = 0

5060      Unload Me

End Sub

Private Sub bcleardiff_click()

5070      pBar = 0

5080      ClearDifferential

5090      cmdSaveHaem.Enabled = True
5100      cmdValidateHaem.Enabled = True

End Sub

Private Sub bhgb_Click()

5110      pBar = 0

5120      ClearHaemExceptHgb

5130      cmdSaveHaem.Enabled = True
5140      cmdValidateHaem.Enabled = True

End Sub

Private Sub cmdOrderCoag_Click()

5150      pBar = 0

5160      If UserHasAuthority(UserMemberOf, "OrderScreen") = False Then
5170          iMsg "You do not have authority to Order Test" & vbCrLf & "Please contact system administrator"
5180          Exit Sub
5190      End If

5200      With frmNewOrder
5210          .FromEdit = True
5220          .Chart = txtChart
5230          .SampleID = txtSampleID
5240          .Show 1
5250      End With

5260      LoadOutstandingCoag

End Sub

Private Sub bViewBioRepeat_Click()

5270      pBar = 0

5280      With frmViewBioRepeat
5290          .SampleID = txtSampleID
5300          .Discipline = "Bio"
5310          .Show 1
5320      End With

5330      LoadBiochemistry

End Sub

Private Sub bViewCoagRepeat_Click()

5340      pBar = 0

5350      With frmCoagRepeats
5360          .EditForm = Me
5370          .SampleID = txtSampleID
5380          .Show 1
5390      End With

End Sub

Private Sub bViewHaemRepeat_Click()

5400      pBar = 0

5410      With frmViewRep
5420          .EditForm = Me
5430          .lSampleID = txtSampleID
5440          .lname = Trim$(txtSurName & " " & txtForeName)
5450          .Show 1
5460      End With
5470      Debug.Print "bViewHaemRepeat_click"
5480      LoadHaematology

End Sub

Private Sub cmbAdd_Click()

          Dim SampleType As String
          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Code As String
          Dim EGFRCode As String
          Dim EGFROK As Boolean
          Dim sCreatCode As String

5490      On Error GoTo cmbAdd_Click_Error

5500      pBar = 0

5510      SampleType = ListCodeFor("ST", cmbSampleType)

5520      sCreatCode = GetOptionSetting("BioCodeForCreatinine", "234")

5530      sql = "SELECT Units FROM BioTestDefinitions WHERE " & _
              "Code = '" & lstAdd.List(cmbAdd.ListIndex) & "' " & _
              "AND SampleType = '" & SampleType & "'"

5540      Set tb = New Recordset
5550      RecOpenServer 0, tb, sql
5560      If Not tb.EOF Then
5570          cmbUnits = tb!Units & ""
5580      Else
5590          cmbUnits = ""
5600      End If

5610      EGFRCode = GetOptionSetting("BioCodeForEGFR", "5555")
5620      If EGFRCode = lstAdd.List(cmbAdd.ListIndex) Then
5630          EGFROK = False
5640          For n = 1 To gBio.Rows - 1
5650              Code = gBio.TextMatrix(n, 9)
5660              If Code = sCreatCode Then
                      '+++ Junaid
                      '190               cmbNewResult = CalculateEGFR(gBio.TextMatrix(n, 1))
                      '---Junaid
5670                  EGFROK = True
5680                  Exit For
5690              End If
5700          Next
5710          If Not EGFROK Then
5720              iMsg "No Creatinine Result." & vbCrLf & "Can't add eGFR.", vbInformation
5730              cmbAdd = ""
5740              cmbUnits = ""
5750              cmbNewResult = ""
5760          End If
5770      End If

5780      Exit Sub

cmbAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

5790      intEL = Erl
5800      strES = Err.Description
5810      LogError "frmEditAll", "cmbAdd_Click", intEL, strES, sql

End Sub


Private Sub cmbAdd_KeyPress(KeyAscii As Integer)

5820      KeyAscii = AutoComplete(cmbAdd, KeyAscii, False)

End Sub


Private Sub cClDetails_Click()

5830      cmdSaveHoldDemographics.Enabled = True
5840      cmdSaveDemographics.Enabled = True

End Sub


Private Sub cClDetails_LostFocus()

          Dim NewText As String

5850      pBar = 0

5860      If Trim$(cClDetails) = "" Then Exit Sub

5870      NewText = ListTextFor("CD", cClDetails)

5880      If NewText <> "" Then
5890          cClDetails = NewText
5900      End If

End Sub


Private Sub cmbClinician_Click()

5910      cmdSaveHoldDemographics.Enabled = True
5920      cmdSaveDemographics.Enabled = True

5930      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP & ":" & cmbClinician


End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

5940      cmdSaveHoldDemographics.Enabled = True
5950      cmdSaveDemographics.Enabled = True
5960      KeyAscii = AutoComplete(cmbClinician, KeyAscii, False)


End Sub


Private Sub cmbClinician_LostFocus()

          Dim strOrig As String

5970      pBar = 0

5980      strOrig = cmbClinician

5990      cmbClinician = ""

6000      cmbClinician = QueryKnown(strOrig, cmbHospital)

6010      If frmOptAllowClinicianFreeText And cmbClinician = "" Then
6020          cmbClinician = strOrig
6030      End If

End Sub

Private Sub cESR_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

6040      pBar = 0

6050      If cESR = 0 Then
6060          If Trim$(tESR) = "Pending" Then
6070              tESR = ""
6080          ElseIf Trim$(tESR) <> "" Then
6090              cESR = 1
6100          End If
6110      Else
6120          If Trim$(tESR) = "" Then
6130              tESR = "Pending"
6140          End If
6150      End If

6160      cmdSaveHaem.Enabled = True

End Sub


Private Sub cFilm_Click()

6170      cmdSaveHaem.Enabled = True

End Sub

Private Sub cmbGP_Change()

6180      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP & ":" & cmbClinician

End Sub

Private Sub cmbGP_Click()


6190      pBar = 0

6200      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP & ":" & cmbClinician

6210      cmdSaveHoldDemographics.Enabled = True
6220      cmdSaveDemographics.Enabled = True

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

6230      cmdSaveHoldDemographics.Enabled = True
6240      cmdSaveDemographics.Enabled = True
6250      KeyAscii = AutoComplete(cmbGP, KeyAscii, False)


End Sub


Private Sub cmbGP_LostFocus()

          Dim strOrig As String
          Dim Gx As New GP

          Dim S As String
          Dim GXs As New GPs
6260      On Error GoTo ErrorHandler

6270      pBar = 0

6280      strOrig = cmbGP

6290      cmbGP = ""

6300      Gx.LoadCodeOrText strOrig
6310      cmbGP = Gx.Text
6320      If sysOptAllowGPFreeText(0) And cmbGP = "" Then
6330          cmbGP = strOrig
6340      End If

6350      If cmdCopyTo.Caption = "cc" Then
6360          If GXs.GpCCed(ListCodeFor("HO", cmbHospital), cmbGP) Then
6370              S = cmbWard & " " & cmbClinician
6380              S = Trim$(S) & " " & cmbGP
6390              S = Trim$(S)

6400              frmCopyTo.lblOriginal = S
6410              frmCopyTo.lblSampleID = txtSampleID
6420              frmCopyTo.Show 1

6430              CheckCC
6440          End If
6450      End If
6460      Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub


Private Sub chkMalaria_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

6470      If chkMalaria = 0 Then
6480          If lblMalaria = "?" Then
6490              lblMalaria = ""
6500          ElseIf lblMalaria <> "" Then
6510              chkMalaria = 1
6520          End If
6530      Else
6540          If lblMalaria = "" Then
6550              lblMalaria = "?"
6560          End If
6570      End If

6580      cmdSaveHaem.Enabled = True

End Sub


Private Sub chkSickledex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

6590      If chkSickledex = 0 Then
6600          If lblSickledex = "?" Then
6610              lblSickledex = ""
6620          ElseIf lblSickledex <> "" Then
6630              chkSickledex = 1
6640          End If
6650      Else
6660          If lblSickledex = "" Then
6670              lblSickledex = "?"
6680          End If
6690      End If

6700      cmdSaveHaem.Enabled = True

End Sub


Private Sub cmdDeleteExt_Click()

          Dim sql As String
          Dim S As String

6710      On Error GoTo cmdDeleteExt_Click_Error

6720      If g.TextMatrix(g.row, 0) = "Test Name" Then Exit Sub

6730      If g.TextMatrix(g.row, 0) = "" Then Exit Sub

6740      S = "Test Name : " & g.TextMatrix(g.row, 0) & vbCrLf & _
              "Delete this test?"
6750      If iMsg(S, vbQuestion + vbYesNo, "Confirm Deletion") = vbYes Then
6760          sql = "DELETE FROM ExtResults WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Analyte = '" & g.TextMatrix(g.row, 0) & "'"
6770          Cnxn(0).Execute sql
6780          LoadExt
6790      End If

6800      Exit Sub

cmdDeleteExt_Click_Error:

          Dim strES As String
          Dim intEL As Integer

6810      intEL = Erl
6820      strES = Err.Description
6830      LogError "frmEditAll", "cmdDeleteExt_Click", intEL, strES, sql


End Sub

Private Sub LoadExt()
       
          Dim sql As String
          Dim tb As Recordset
          Dim tbR As Recordset
          Dim tbD As Recordset
          Dim S As String
          Dim AnalyteName As String
          Dim SelectNormalRange As String
          Dim n As Integer
          Dim l_Request As Boolean

6840      On Error GoTo LoadExt_Error
6850      l_Request = False
6860      g.Rows = 2
6870      g.AddItem ""
6880      g.RemoveItem 1

6890      For n = 0 To 8
6900          etc(n) = ""
6910      Next

6920      If Val(txtSampleID) = 0 Then
6930          Exit Sub
6940      End If

6950      sql = "Select Sex from Demographics where " & _
              "SampleID = '" & txtSampleID & "'"
6960      Set tb = Cnxn(0).Execute(sql)
6970      If Not tb.EOF Then
6980          Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
                  Case "M": SelectNormalRange = " MaleLow as Low, MaleHigh as High, "
6990              Case "F": SelectNormalRange = " FemaleLow as Low, FemaleHigh as High, "
7000              Case Else: SelectNormalRange = " FemaleLow as Low, MaleHigh as High, "
7010          End Select
7020      Else
7030          SelectNormalRange = " FemaleLow as Low, MaleHigh as High, "
7040      End If

7050      sql = "Select E.*, " & SelectNormalRange & " D.Units, D.SendTo  " & _
              "from ExtResults as E, ExternalDefinitions as D " & _
              "where SampleID = '" & txtSampleID & "' " & _
              "and E.Analyte = D.AnalyteName order by Date desc"
7060      l_Request = True
7070      Set tb = New Recordset
7080      RecOpenServer 0, tb, sql
7090      Do While Not tb.EOF
7100          l_Request = False
7110          AnalyteName = tb!Analyte & ""
7120          S = AnalyteName & vbTab
7130          S = S & tb!Result & vbTab
              '+++ Junaid 07-08-2023
7140          S = S & "" & vbTab 'Format(tb!Low & "") & "-" & Format(tb!High & "") & vbTab
7150          S = S & "" & vbTab 'tb!Units & vbTab
              '--- Junaid
7160          S = S & tb!SendTo & vbTab
7170          If Not IsNull(tb!Date) Then
7180              S = S & Format(tb!Date, "dd/mmm/yyyy")
7190          End If
7200          g.AddItem S
7210          tb.MoveNext
7220      Loop
7230      If g.Rows > 2 Then
7240          g.RemoveItem 1
7250      End If
          '+++ Junaid 27-06-2022
7260      If l_Request Then
7270          sql = "Select IsNULL(M.TargetValue,'') TargetValue from ocmmapping M Inner Join ocmRequestDetails D On D.TestCode = M.SourceValue Where D.External_ = '1' And D.SampleID = '" & txtSampleID.Text & "' And Display = 1"
7280          Set tbR = New Recordset
7290          RecOpenServer 0, tbR, sql
7300          If Not tbR Is Nothing Then
7310              If Not tbR.EOF Then
7320                  g.Rows = 1
7330                  g.row = 0
7340                  While Not tbR.EOF
7350                      sql = "Select * from ExternalDefinitions Where AnalyteName = '" & tbR!TargetValue & "'"
7360                      Set tbD = New Recordset
7370                      RecOpenServer 0, tbD, sql
7380                      If Not tbD Is Nothing Then
7390                          If Not tbD.EOF Then
7400                              While Not tbD.EOF
7410                                  S = tbD!AnalyteName & vbTab & "" & vbTab & "" & vbTab & tbD!Units & vbTab & tbD!SendTo & vbTab & Format(Date, "dd/mmm/yyyy")
7420                                  g.AddItem (S)
7430                                  tbD.MoveNext
7440                              Wend
7450                          End If
7460                      End If
7470                      tbR.MoveNext
                          '                    MsgBox Sql
7480                  Wend
7490              End If
7500          End If
7510      End If
          '-- Junaid 27-06-2022

7520      sql = "SELECT * FROM etc WHERE SampleID = '" & txtSampleID & "'"
7530      Set tb = New Recordset
7540      RecOpenServer 0, tb, sql
7550      If Not tb.EOF Then
7560          For n = 0 To 8
7570              etc(n) = tb("etc" & Format$(n)) & ""
7580          Next
7590      End If

7600      cmdSaveExt.Enabled = False

7610      sql = "SELECT COUNT(*) Tot FROM MedibridgeResults WHERE SampleID = '" & txtSampleID & "'"
7620      Set tb = New Recordset
7630      RecOpenServer 0, tb, sql
7640      cmdMedibridge.Visible = tb!Tot > 0


          Dim sql3 As String
          Dim sql4 As String
          Dim tb3 As Recordset
          Dim tb4 As Recordset
          
240       sql3 = "SELECT DateTimeDemographics FROM Demographics WHERE SampleID = '" & Trim(txtSampleID.Text) & "'"
250       Set tb3 = New Recordset
260       RecOpenClient 0, tb3, sql3
270       If Not tb3.EOF Then
280           etc(0).Text = "Date Time of Demographics: " & tb3!DateTimeDemographics
290       End If
          
300       sql4 = "SELECT DateTimeOfRecord FROM BiomnisRequests WHERE SampleID = '" & Trim(txtSampleID.Text) & "'"
310       Set tb4 = New Recordset
320       RecOpenClient 0, tb4, sql4
330       If Not tb4.EOF Then
340           etc(1).Text = "Date Time of Ext: " & tb4!DateTimeOfRecord
350       End If

7650      Exit Sub

LoadExt_Error:

          Dim strES As String
          Dim intEL As Integer

7660      intEL = Erl
7670      strES = Err.Description
7680      LogError "frmEditAll", "LoadExt", intEL, strES, sql

End Sub

Private Sub cmdSetPrinter_Click()

7690      frmForcePrinter.From = Me
7700      frmForcePrinter.Show 1

7710      If pPrintToPrinter = "Automatic Selection" Then
7720          pPrintToPrinter = ""
7730      End If

7740      If pPrintToPrinter <> "" Then
7750          cmdSetPrinter.BackColor = vbRed
7760          cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
7770      Else
7780          cmdSetPrinter.BackColor = vbButtonFace
7790          pPrintToPrinter = ""
7800          cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
7810      End If

End Sub

Private Sub cmdValidateHaem_Click()

7820      pBar = 0
          'Zyam 15-06-24
7830      cmdUpDown(0).Enabled = False
7840      cmdUpDown(1).Enabled = False
          'Zyam 15-06-24

7850      If UserHasAuthority(UserMemberOf, "HaemSave") = False Then
7860          iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
7870          Exit Sub
7880      End If

7890      SaveHaematology 1
7900      SaveComments
          'DoEvents
          'txtSampleID = Val(txtSampleID + 1)    ' MoveToNextRelevant 1
7910      Debug.Print "cmdvalidatehaem_click"
7920      LoadAllDetails

7930      UpdateMRU Me
          'Zyam 15-06-24
7940      cmdUpDown(0).Enabled = True
7950      cmdUpDown(1).Enabled = True
          'Zyam 15-06-24

End Sub

Private Sub cmdValidationList_Click()

7960      frmDemographicValidation.Show 1

End Sub

Private Sub cmdViewReports_Click()

          Dim f As Form

7970      Set f = New frmReportViewer

7980      f.Dept = "Biochemistry"
7990      f.SampleID = txtSampleID
8000      f.Show 1

8010      Set f = Nothing

End Sub



Private Sub cmdViewScan_Click()
8020      frmViewScan.CallerDepartment = SSTab1.TabCaption(SSTab1.Tab)
8030      frmViewScan.SampleID = txtSampleID
8040      frmViewScan.txtSampleID = txtSampleID
8050      frmViewScan.Show 1

End Sub

Private Sub cMonospot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

8060      If cMonospot = 0 Then
8070          If Trim$(tMonospot) = "?" Then
8080              tMonospot = ""
8090          ElseIf Trim$(tMonospot) <> "" Then
8100              cMonospot = 1
8110          End If
8120      Else
8130          If Trim$(tMonospot) = "" Then
8140              tMonospot = "?"
8150          End If
8160      End If

8170      cmdSaveHaem.Enabled = True

End Sub


Private Sub cMRU_Click()

8180      txtSampleID = cMRU
8190      Debug.Print "cmru_click"

8200      LoadAllDetails

8210      cmdSaveHoldDemographics.Enabled = False
8220      cmdSaveDemographics.Enabled = False
8230      cmdSaveHaem.Enabled = False
8240      cmdSaveBio.Enabled = False
8250      cmdSaveCoag.Enabled = False

End Sub


Private Sub cMRU_KeyPress(KeyAscii As Integer)

8260      KeyAscii = 0

End Sub






'---------------------------------------------------------------------------------------
' Procedure : bsearchDob_Click
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bsearchDob_Click()
8270      On Error GoTo bsearchDob_Click_Error


8280      LoadMatchingDemo


8290      Exit Sub


bsearchDob_Click_Error:

          Dim strES As String
          Dim intEL As Integer

8300      intEL = Erl
8310      strES = Err.Description
8320      LogError "frmEditAll", "bsearchDob_Click", intEL, strES
End Sub

Private Sub cmdOrderHeam_Click()
          Dim f As Form

8330      On Error GoTo cmdOrderHeam_Click_Error

8340      pBar = 0

8350      If UserHasAuthority(UserMemberOf, "OrderScreen") = False Then
8360          iMsg "You do not have authority to Order Test" & vbCrLf & "Please contact system administrator"
8370          Exit Sub
8380      End If

8390      SaveDemographics 0
8400      Set f = New frmNewOrder

8410      With f
8420          .FromEdit = True
8430          .SampleID = txtSampleID
8440          .Chart = txtChart
8450          .Show 1
8460          If Not .Cancelled And sysOptAutoScrollAfterOrder(0) Then
8470              cmdSaveDemographics_Click
8480          End If
8490      End With

8500      Set f = Nothing
8510      Debug.Print "bordertests_click"

8520      LoadAllDetails

8530      Exit Sub

cmdOrderHeam_Click_Error:
          Dim strES As String
          Dim intEL As Integer

8540      intEL = Erl
8550      strES = Err.Description
8560      LogError "frmEditAll", "cmdOrderHeam_Click", intEL, strES
End Sub

Private Sub cRetics_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

8570      If cRetics = 0 Then
8580          If Trim$(tRetA) = "?" Then
8590              tRetA = ""
8600              tRetP = ""
8610          ElseIf Trim$(tRetA) <> "" Then
8620              cRetics = 1
8630          End If
8640      Else
8650          If Trim$(tRetA) = "" Then
8660              tRetA = "?"
                  '110       tRetP = "?"
8670          End If
8680      End If

8690      cmdSaveHaem.Enabled = True

End Sub


Private Sub cRooH_Click(Index As Integer)

8700      cmdSaveHoldDemographics.Enabled = True
8710      cmdSaveDemographics.Enabled = True

End Sub

Private Sub cmbSampleType_Click()

8720      FillcmbAdd

End Sub


Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

8730      KeyAscii = 0

End Sub


Private Sub cmbUnits_KeyPress(KeyAscii As Integer)

8740      KeyAscii = 0

End Sub


Private Sub cmbWard_Change()

8750      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP & ":" & cmbClinician

End Sub

Private Sub cmbWard_Click()

8760      lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

8770      cmdSaveHoldDemographics.Enabled = True
8780      cmdSaveDemographics.Enabled = True

End Sub


Private Sub cmbWard_KeyPress(KeyAscii As Integer)

8790      cmdSaveHoldDemographics.Enabled = True
8800      cmdSaveDemographics.Enabled = True
8810      KeyAscii = AutoComplete(cmbWard, KeyAscii, False)

End Sub


Private Sub cmbWard_LostFocus()

          Dim Found As Boolean
          Dim tb As Recordset
          Dim sql As String
          Dim strWard As String

8820      On Error GoTo cmbWard_LostFocus_Error

8830      If Trim$(cmbWard) = "" Then
8840          cmbWard = "GP"
8850          Exit Sub
8860      End If

8870      strWard = cmbWard
8880      Found = False
8890      sql = "Select * from Wards where " & _
              "(Text = '" & AddTicks(cmbWard) & "' " & _
              "or Code = '" & AddTicks(cmbWard) & "') " & _
              "and HospitalCode = '" & ListCodeFor("HO", cmbHospital) & "' and InUse = 1"
8900      Set tb = New Recordset
8910      RecOpenServer 0, tb, sql
8920      If Not tb.EOF Then
8930          strWard = tb!Text & ""
8940          Found = True
8950      End If
          
          

8960      cmbWard = strWard
8970      If Not sysOptAllowWardFreeText(0) Then
8980          If Not Found Then
8990              cmbWard = "GP"
9000          End If
9010      End If

9020      Exit Sub

cmbWard_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

9030      intEL = Erl
9040      strES = Err.Description
9050      LogError "frmEditAll", "cmbWard_LostFocus", intEL, strES, sql


End Sub


Private Sub dtRecDate_CloseUp()

9060      pBar = 0

9070      cmdSaveHoldDemographics.Enabled = True
9080      cmdSaveDemographics.Enabled = True

End Sub


Private Sub dtRecDate_LostFocus()
9090      SetDatesColour
End Sub

Private Sub dtRunDate_CloseUp()

9100      pBar = 0

9110      If DateDiff("d", dtRunDate, dtSampleDate) > 0 Then
9120          dtRunDate = dtSampleDate
9130      End If

9140      cmdSaveHoldDemographics.Enabled = True
9150      cmdSaveDemographics.Enabled = True

End Sub


Private Sub dtRunDate_LostFocus()
9160      SetDatesColour
End Sub

Private Sub dtSampleDate_CloseUp()

9170      pBar = 0

9180      cmdSaveHoldDemographics.Enabled = True
9190      cmdSaveDemographics.Enabled = True

End Sub


Private Sub dtSampleDate_LostFocus()
9200      SetDatesColour
End Sub

Private Sub etc_KeyPress(Index As Integer, KeyAscii As Integer)

9210      cmdSaveExt.Enabled = True

End Sub


Private Sub Form_Activate()
9220      On Error GoTo ErrorHandler
9230      m_ShowDoc = False
9240      pBar = 0
9250      pBar.max = LogOffDelaySecs
9260      TimerBar.Enabled = True

9270      SSTab1.TabVisible(4) = False
9280      SSTab1.TabVisible(5) = False

9290      SaveOptionSetting "DEMOGRAPHICSNAMECAPS", "1"

9300      If sysOptDontPrintAllCoag(0) Then bPrintAll.Visible = False

9310      StatusBar1.Panels(1).Text = UserName

9320      g.RowHeight(0) = 300

9330      Call ChangeComboHeight(cmbClinician, GetOptionSetting("ClinicianListLength", 8))
9340      Call ChangeComboHeight(cmbGP, GetOptionSetting("GPListLength", 8))
9350      Call ChangeComboHeight(cmbHospital, GetOptionSetting("ClinicianListLength", 8))
9360      Call ChangeComboHeight(cmbWard, GetOptionSetting("WardListLength", 8))
9370      Call ChangeComboHeight(cmbDemogComment, GetOptionSetting("CommentListLength", 8))
9380      Call ChangeComboHeight(cmbAdd, GetOptionSetting("TestListLength", 8))
          'CheckDemographics
9390      cmbClinician.SelLength = 0
9400      cmbGP.SelLength = 0
9410      cmbHospital.SelLength = 0
9420      cmbWard.SelLength = 0

9430      Call FixComboWidth(cmbHaemComment)
9440      Call FixComboWidth(cmbFilmComment)
9450      Call FixComboWidth(cmbBioComment(0))
9460      Call FixComboWidth(cmbDemogComment)
9470      Call FixComboWidth(cClDetails)
9480      Call FixComboWidth(cmbAdd)

9490      ShowMenuLists

9500      gBio.ColWidth(9) = 0
9510      grdOutstanding.ColWidth(1) = 0
9520      Call FormatGrid
          '290   If InStr(UCase$(App.Path), "TEST") Then
          '300       cmdPatientNotePad.Visible = True
          '310   Else
          '320       cmdPatientNotePad.Visible = False
          '330   End If

9530      Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
          Dim strES As String
          Dim intEL As Integer
9540      intEL = Erl
9550      strES = Err.Description
9560      LogError "frmEditAll", "LabNoUpdatePrv", intEL, strES

End Sub

Private Sub Form_Click()
    'Dim medibridgepathtoviewer As String
    '20    medibridgepathtoviewer = GetOptionSetting("MedibridgePathToViewer", "")
    '30    If medibridgepathtoviewer <> "" Then
    '40      SaveOptionSetting "MedibridgeSampleID", txtSampleID
    '50      Shell medibridgepathtoviewer & " /SampleID=" & txtSampleID & _
    '                                       " /UserName=" & UserName & _
    '                                       " /Password=" & TechnicianPassFor(UserName) & _
    '                                       " /Department=Haematology", vbNormalFocus
    '60    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LabNoUpdatePrv
' Author    : XPMUser
' Date      : 23/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LabNoUpdatePrvColor()
9570      On Error GoTo LabNoUpdatePrv_Error


          'If UCase(LabNoUpdatePrviousData) = UCase(txtSurName & txtForeName & txtDoB) Then
9580      If LabNoUpdatePrviousData = "1" Then
9590          txtLabNo.BackColor = vbGreen
              '        lAddWardGP = FindLatestAddress(txtChart, Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, Left$(txtSex, 1), txtLabNo)
9600      Else
              'LabNoUpdatePrviousData = ""
9610          txtLabNo.BackColor = vbRed
9620      End If


9630      Exit Sub


LabNoUpdatePrv_Error:

          Dim strES As String
          Dim intEL As Integer

9640      intEL = Erl
9650      strES = Err.Description
9660      LogError "frmEditAll", "LabNoUpdatePrv", intEL, strES
End Sub

Private Sub gOutstandingHaem_Click()

          Dim sql As String



9670      On Error GoTo gOutstandingHaem_Click_Error

9680      With gOutstandingHaem
9690          If .MouseRow = 0 Then Exit Sub
9700          .row = .MouseRow
9710          If .Text = "" Then Exit Sub
9720          If iMsg("Remove " & .TextMatrix(.row, 0) & " from Requests?", vbQuestion + vbYesNo) = vbYes Then

9730              sql = "DELETE FROM HaeRequests " & _
                      "WHERE SampleID = '" & txtSampleID & "' " & _
                      "AND Code = '" & .TextMatrix(.row, 0) & "'"
9740              Cnxn(0).Execute sql

9750              If .Rows > 2 Then
9760                  .RemoveItem .row
9770              Else
9780                  .AddItem ""
9790                  .RemoveItem 1
9800              End If

9810          End If

9820      End With


9830      Exit Sub

gOutstandingHaem_Click_Error:
          Dim strES As String
          Dim intEL As Integer

9840      intEL = Erl
9850      strES = Err.Description
9860      LogError "frmEditAll", "gOutstandingHaem_Click", intEL, strES

End Sub

Private Sub lblIcteric_Click()

9870      With lblIcteric()
9880          Select Case .Caption
                  Case "", "0": .Caption = "1+"
9890              Case "1+": .Caption = "2+"
9900              Case "2+": .Caption = "3+"
9910              Case "3+": .Caption = "4+"
9920              Case "4+": .Caption = "5+"
9930              Case "5+": .Caption = "6+"
9940              Case "6+": .Caption = ""
9950          End Select
9960      End With

9970      cmdSaveBio.Enabled = True

End Sub

Private Sub lblLipaemic_Click()

9980      With lblLipaemic()
9990          Select Case .Caption
                  Case "", "0": .Caption = "1+"
10000             Case "1+": .Caption = "2+"
10010             Case "2+": .Caption = "3+"
10020             Case "3+": .Caption = "4+"
10030             Case "4+": .Caption = "5+"
10040             Case "5+": .Caption = "6+"
10050             Case "6+": .Caption = ""
10060         End Select
10070     End With

10080     cmdSaveBio.Enabled = True

End Sub


Private Sub lblMalaria_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
10090     On Error GoTo ErrorHandler

          Dim f As Form

10100     cmdSaveHaem.Enabled = True

10110     If lblMalaria = "" Or lblMalaria = "?" Then
10120         lblMalaria = "Negative"
10130     ElseIf lblMalaria = "Negative" Then
10140         lblMalaria = "Positive"
10150     Else
10160         lblMalaria = ""
10170     End If

10180     If lblMalaria <> "" Then
10190         If Not CheckReagentLotNumber("Malaria", txtSampleID) Then
10200             Set f = New frmCheckReagentLotNumber
10210             With f
10220                 .Analyte = "Malaria"
10230                 .SampleID = txtSampleID
10240                 .Show 1
10250             End With
10260             Unload f
10270             Set f = Nothing
10280         End If
10290     End If
10300     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub lblRA_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10310     cmdSaveHaem.Enabled = True

10320     If lblRA = "" Or lblRA = "?" Then
10330         lblRA = "Negative"
10340     ElseIf lblRA = "Negative" Then
10350         lblRA = "Positive"
10360     Else
10370         lblRA = ""
10380     End If

End Sub


Private Sub lblSickledex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
10390     On Error GoTo ErrorHandler
          Dim f As Form

10400     cmdSaveHaem.Enabled = True

10410     If lblSickledex = "" Or lblSickledex = "?" Then
10420         lblSickledex = "Negative"
10430     ElseIf lblSickledex = "Negative" Then
10440         lblSickledex = "Positive"
10450     Else
10460         lblSickledex = ""
10470     End If

10480     If lblSickledex <> "" Then
10490         If Not CheckReagentLotNumber("Sickledex", txtSampleID) Then
10500             Set f = New frmCheckReagentLotNumber
10510             With f
10520                 .Analyte = "Sickledex"
10530                 .SampleID = txtSampleID
10540                 .Show 1
10550             End With
10560             Unload f
10570             Set f = Nothing
10580         End If
10590     End If
10600     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub mnuAddCoagTest_Click()

10610     frmAddCoagTest.Show 1

End Sub

Private Sub mnuAmendAnalyte_Click()

10620     With frmBioAmendCode
10630         .Discipline = "Bio"
10640         .Show 1
10650     End With

End Sub

Private Sub MnuAssignPanelsForHealthLink_Click()

10660     frmAssignPanelsForHealthLink.Show 1

End Sub

Private Sub mnuAutoValHaem_Click()

10670     frmAutoValHaem.Show 1

End Sub

Private Sub mnuBarCodesH_Click()

10680     frmBarCodes.Show 1

End Sub

Private Sub mnuCoagDefinitions_Click()

10690     frmCoagDefinitions.Show 1

End Sub

Private Sub mnuCoagPanels_Click()

10700     frmCoagPanels.Show 1

End Sub


Private Sub mnuFasting_Click()

10710     frmTestFastings.Show 1

End Sub

Private Sub mnuHaemDefinitions_Click()

10720     frmHaemDefinitions.Show 1

End Sub

Private Sub mnuListsCoagulation_Click()

End Sub

Private Sub mnuLIH_Click()

10730     frmSetLIH.Show 1

10740     LiIcHas.Clear
10750     LiIcHas.Load

End Sub



Private Sub mnuNewAnalyte_Click()

10760     With frmBioAddAnalyte
10770         .Discipline = "Bio"
10780         .Show 1
10790     End With
10800     FillcmbAdd

End Sub

Private Sub mnuReRunTimes_Click()

10810     frmReRunTimes.Show 1

End Sub


Private Sub Form_Deactivate()

10820     pBar = 0
10830     TimerBar.Enabled = False

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

10840     pBar = 0

End Sub

Private Sub Form_Load()
10850     On Error GoTo ErrorHandler
          '+++Abubaker 16-11-2023
          'm_Counter = 0
          '---Abubaker 16-11-2023

          Dim n As Integer
          Dim HideComments As Boolean
          Dim LastUsedBioSplitName As String
          Dim LastUsedBioSplitIndex As Integer

10860     CheckScannedImagesInDb

10870     SSTab1.TabVisible(0) = UserHasAuthority(UserMemberOf, "DemTab")
10880     SSTab1.TabVisible(1) = UserHasAuthority(UserMemberOf, "HaemTab")
10890     SSTab1.TabVisible(2) = UserHasAuthority(UserMemberOf, "BioTab")
10900     SSTab1.TabVisible(3) = UserHasAuthority(UserMemberOf, "CoagTab")
10910     SSTab1.TabVisible(6) = UserHasAuthority(UserMemberOf, "ExtTab")

10920     mnuListsBio.Visible = False
10930     mnuListsHaem.Visible = False
10940     mnuListsCoag.Visible = False
10950     If UserMemberOf = "Managers" Then
10960         mnuNull.Visible = True
10970     End If

10980     SetFormOptions

10990     HideComments = GetOptionSetting("HideBioComment", "0") = "1"
11000     If HideComments Then
11010         gBio.ColWidth(2) = gBio.ColWidth(2) + gBio.ColWidth(7)
11020         gBio.ColWidth(7) = 0
11030     End If

11040     cmdCopyTo.Visible = True

11050     cmdFAX.Visible = True

11060     n = 1

11070     If Not frmOptDeptHaem Then SSTab1.TabVisible(1) = False Else n = n + 1
11080     If Not frmOptDeptBio Then SSTab1.TabVisible(2) = False Else n = n + 1
11090     If Not frmOptDeptCoag Then SSTab1.TabVisible(3) = False Else n = n + 1

11100     If Not frmOptDeptExt Then
11110         SSTab1.TabVisible(6) = False
11120         cmdOrderExt(0).Visible = False
11130         cmdOrderExt(1).Visible = False
11140     Else
11150         n = n + 1
11160         cmdOrderExt(0).Visible = True
              '    cmdOrderExt(1).Visible = True
11170     End If

11180     chkUrgent.Visible = frmOptUrgent
11190     chkUrgent.Visible = True

11200     LastUsedBioSplitName = GetOptionSetting("LastUsedBioSplitName", "All")
11210     LastUsedBioSplitIndex = GetOptionSetting("LastUsedBioSplitIndex", "0")
11220     lblSplitView.Caption = "Viewing " & LastUsedBioSplitName
11230     lblSplitView.Tag = LastUsedBioSplitIndex
11240     lblSplitView.BackColor = vbButtonFace
11250     lblSplitView.ForeColor = vbBlack
11260     lblSplitView.FontBold = False
11270     For n = 1 To 6
11280         lblSplit(n).Caption = GetOptionSetting("BioSplitName" & Format$(n), "Split " & Format$(n))
11290         If lblSplit(n).Caption = LastUsedBioSplitName Then
11300             lblSplit(n).BackColor = vbRed
11310             lblSplit(n).ForeColor = vbYellow
11320             lblSplitView.BackColor = vbRed
11330             lblSplitView.ForeColor = vbYellow
11340             lblSplitView.FontBold = True
11350         End If
11360     Next

11370     FillcmbSampleType
11380     FillcParameter
11390     FillcmbAdd

11400     FillLists

11410     FillMRU Me
11420     FillCommentTemplates cmbBioComment(1), "B"
11430     FillCommentTemplates cmbBioComment(2), "C"

11440     With lblChartNumber
11450         .BackColor = &H8000000F
11460         .ForeColor = vbBlack
11470         .Caption = HospName(0) & " Chart #"
              '    txtSurName.Left = 3960
              '    txtSurName.Width = 2505
              '    txtForeName.Left = 6480
              '    txtForeName.Width = 3495
              '    lblSurNameTitle.Left = 3960
              '    lblForeNameTitle.Left = 6480
11480     End With

11490     dtRunDate = Format$(Now, "dd/mm/yyyy")
11500     dtSampleDate = Format$(Now, "dd/mm/yyyy")

11510     txtSampleID = Format$(Val(GetSetting("NetAcquire", "StartUp", "LastUsed", "1")))
11520     LoadAllDetails
11530     lblResultOrRequest = "Results"
11540     If SSTab1.Tab = 0 Then
11550         lblResultOrRequest = "SampleID"
11560     End If

11570     cmdSaveHoldDemographics.Enabled = False
11580     cmdSaveDemographics.Enabled = False
11590     cmdSaveHaem.Enabled = False
11600     cmdSaveBio.Enabled = False
11610     cmdSaveCoag.Enabled = False
11620     cmdSaveExt.Enabled = False

11630     Activated = False
          'txtLabNo.Text = ""
11640     MatchingDemoLoaded = False
11650     FormLoaded = True

          'txtExtSampleID = 10344
          'txtExtSampleID_LostFocus
11660     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FndMaxID
' Author    : XPMUser
' Date      : 17/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'


Public Sub LoadCoagulation()

          Dim CRs As New CoagResults
          Dim CRr As New CoagResults
          Dim CR As CoagResult
          Dim S As String

          Dim l_Result As String

          Dim udtCoag As udtHaem

11670     On Error GoTo LoadCoagulation_Error
11680     SetViewReports "Coagulation", txtSampleID
11690     lblDateConflict.Visible = False
11700     lblDelta(0) = ""
11710     Set CRs = CRs.Load(txtSampleID, gDONTCARE, gDONTCARE, "Results")
11720     Set CRr = CRr.Load(txtSampleID, gDONTCARE, gDONTCARE, "Repeats")

11730     lblCoagAnalyser = ""
11740     txtAutoComment(3) = CheckAutoComments(txtSampleID, 3)

11750     ClearCoagulation

11760     SSTab1.TabCaption(3) = "Coagulation"

11770     For Each CR In CRs
11780         If CR.InUse Then
11790             If Val(CR.Result) < Val(CR.PLow) Then
11800                 l_Result = "XXXX"
11810             ElseIf Val(CR.Result) > Val(CR.PHigh) Then
11820                 l_Result = "XXXX"
11830             Else
11840                 l_Result = CR.Result
11850             End If
11860             lblCoagAnalyser = CR.Analyser
11870             S = CR.TestName & vbTab & _
                      l_Result & vbTab & _
                      CR.Units & vbTab & _
                      CR.Flag & vbTab & _
                      IIf(CR.Valid, "V", "") & _
                      IIf(CR.Printed, "P", "")
11880             gCoag.AddItem S

11890             udtCoag = GetCoagInfo(CR.TestName, txtSex, txtDoB)
11900             If udtCoag.DoDelta = True Then
11910                 lblDelta(0) = lblDelta(0) & FindCoagPrvTestValue(CR.Code, udtCoag.DeltaDaysBackLimit, CR.TestName)
11920             End If
11930             If CR.Low = 0 And (CR.High = 999 Or CR.High = 0 Or CR.High = 9999) Then
11940             Else
11950                 If Val(CR.Result) < CR.Low Then
11960                     gCoag.row = gCoag.Rows - 1
11970                     gCoag.Col = 1
11980                     gCoag.CellBackColor = vbBlue
11990                     gCoag.CellForeColor = vbYellow
12000                     gCoag.CellFontBold = True
12010                     gCoag.Col = 3
12020                     gCoag.Text = "L"
12030                     gCoag.CellAlignment = flexAlignCenterCenter
12040                 ElseIf Val(CR.Result) > CR.High Then
12050                     gCoag.row = gCoag.Rows - 1
12060                     gCoag.Col = 1
12070                     gCoag.CellBackColor = vbRed
12080                     gCoag.CellForeColor = vbYellow
12090                     gCoag.CellFontBold = True
12100                     gCoag.Col = 3
12110                     gCoag.Text = "H"
12120                     gCoag.CellAlignment = flexAlignCenterCenter
12130                 End If
12140             End If
12150             CheckIfMustPhone "Coagulation", CR.TestName, CR.Result
12160             If CR.Valid Then CoagValBy = CR.OperatorCode
12170         End If
12180     Next

12190     bViewCoagRepeat.Visible = CRr.Count <> 0

12200     If gCoag.Rows > 2 Then
12210         SSTab1.TabCaption(3) = ">>Coagulation<<"
12220         CheckRunSampleDates
12230         gCoag.RemoveItem 1
12240     End If

12250     LoadOutstandingCoag
12260     LoadPreviousCoag

12270     SetPrintInhibit "Coa"

12280     Exit Sub

LoadCoagulation_Error:

          Dim strES As String
          Dim intEL As Integer

12290     intEL = Erl
12300     strES = Err.Description
12310     LogError "frmEditAll", "LoadCoagulation", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : DeltaCheckCoag
' Author    : XPMUser
' Date      : 13/Aug/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : FindCoagPrvTestValue
' Author    : XPMUser
' Date      : 13/Aug/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FindCoagPrvTestValue(AnalyteCode As String, DeltaDaysBackLimit As Single, AnalyteName As String) As String

          Dim sql As String
          Dim sn As Recordset
          Dim DateQry As String

12320     On Error GoTo FindCoagPrvTestValue_Error


12330     If IsDate(tSampleTime) Then
12340         DateQry = " AND (D.SampleDate < '" & Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm") & "')"
12350     Else
12360         DateQry = " AND (D.SampleDate < '" & Format$((dtSampleDate), "dd/MMM/yyyy") & "')"
12370     End If


12380     If AddTicks(txtSurName & " " & txtForeName) & "' " <> "" Then
12390         sql = " SELECT     TOP (1) R.Rundate,D.SampleDate,R.sampleid " & _
                  ",R.Result " & _
                  " FROM         demographics AS D INNER JOIN CoagResults AS R ON D.SampleID = R.sampleid  " & _
                  " WHERE     (D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "') AND (D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "') " & _
                  " AND (D.SampleDate >= '" & Format$((dtSampleDate - DeltaDaysBackLimit), "dd/MMM/yyyy") & "')  " & _
                  DateQry & _
                  " AND Hospital = '" & cmbHospital & "' " & _
                  " AND R.Code = '" & AnalyteCode & "'" & _
                  " ORDER BY D.SampleDate,D.sampleid ASC "

12400         Set sn = New Recordset
12410         RecOpenServer 0, sn, sql
12420         If Not sn.EOF Then
12430             FindCoagPrvTestValue = AnalyteName & "  " & sn!SampleID & " ( " & sn!Rundate & " )"
12440         End If
12450     End If


12460     Exit Function


FindCoagPrvTestValue_Error:

          Dim strES As String
          Dim intEL As Integer

12470     intEL = Erl
12480     strES = Err.Description
12490     LogError "frmEditAll", "FindCoagPrvTestValue", intEL, strES, sql

End Function

Public Sub LoadDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleDate As String
          Dim RooH As Boolean
          Dim Dxs As New Demographics
          Dim dx As Demographic

12500     On Error GoTo LoadDemographics_Error

12510     UrgentTest = False
12520     RooH = IsRoutine()
12530     cRooH(0) = RooH
12540     cRooH(1) = Not RooH
12550     bViewBB.Enabled = False
12560     txtAge = ""
12570     lblAge = ""

12580     lblDemogValid = "Demographics Not Valid"
12590     lblDemogValid.BackColor = vbRed
12600     lblDemogValid.ForeColor = vbYellow
12610     cmdValidateDemographics.Visible = False

12620     If Trim$(txtSampleID) = "" Then Exit Sub

12630     lRandom = "Random Sample"

12640     Screen.MousePointer = 11
12650     If FormLoaded Then txtLabNo = ""   'Val(FndMaxID("demographics", "LabNo", ""))

12660     Dxs.Load txtSampleID
12670     If Dxs.Count = 0 Then
12680         mNewRecord = True
12690         dtRunDate = Format$(Now, "dd/mm/yyyy")
12700         dtSampleDate = Format$(Now, "dd/mm/yyyy")
12710         txtChart = ""
12720         txtSurName = ""
12730         txtForeName = ""
12740         txtAddress(0) = ""
12750         txtAddress(1) = ""
12760         txtExtSampleID = ""
12770         txtSex = ""
12780         txtDoB = ""
12790         txtAge = ""
12800         cmbWard = "GP"
12810         cmbClinician = ""
12820         cmbGP = ""
12830         cmbHospital = HospName(0)
12840         txtDemographicComment = ""
12850         cmbDemogComment = ""
12860         tSampleTime.Mask = ""
12870         tSampleTime.Text = ""
12880         tSampleTime.Mask = "##:##"
12890         lblChartNumber.Caption = HospName(0) & " Chart #"
12900         lblChartNumber.BackColor = &H8000000F
12910         lblChartNumber.ForeColor = vbBlack
12920         cClDetails = ""
12930     Else
12940         Set dx = Dxs(1)
12950         If Trim$(dx.Hospital & "") <> "" Then
12960             cmbHospital = Trim$(dx.Hospital)
12970             lblChartNumber = Trim$(dx.Hospital) & " Chart #"
12980             If UCase$(dx.Hospital) = UCase$(HospName(0)) Then
12990                 lblChartNumber.BackColor = &H8000000F
13000                 lblChartNumber.ForeColor = vbBlack
13010             Else
13020                 lblChartNumber.BackColor = vbRed
13030                 lblChartNumber.ForeColor = vbYellow
13040             End If
13050         Else
13060             cmbHospital = HospName(0)
13070             lblChartNumber.Caption = HospName(0) & " Chart #"
13080             lblChartNumber.BackColor = &H8000000F
13090             lblChartNumber.ForeColor = vbBlack
13100         End If
13110         If IsDate(dx.Rundate) Then
13120             dtRunDate = Format$(dx.Rundate, "dd/mm/yyyy")
13130         Else
13140             dtRunDate = Format$(Now, "dd/mm/yyyy")
13150         End If
13160         StatusBar1.Panels(4).Text = dtRunDate
13170         mNewRecord = False
13180         cRooH(0) = IIf(dx.RooH, True, False)
13190         cRooH(1) = Not cRooH(0)
13200         txtChart = dx.Chart
13210         txtSurName = SurName(dx.PatName)
13220         txtForeName = ForeName(dx.PatName)
13230         txtAddress(0) = dx.Addr0
13240         txtAddress(1) = dx.Addr1
13250         txtExtSampleID = dx.ExtSampleID
13260         If Val(dx.LabNo) <> 0 Then
13270             txtLabNo = dx.LabNo
13280         End If
13290         Select Case Left$(Trim$(UCase$(dx.Sex)), 1)
                  Case "M": txtSex = "Male"
13300             Case "F": txtSex = "Female"
13310             Case Else: txtSex = ""
13320         End Select
13330         If IsDate(dx.DoB) Then
13340             txtDoB = Format$(dx.DoB, "dd/mm/yyyy")
13350         Else
13360             txtDoB = ""
13370         End If
13380         txtAge = dx.Age

13390         cmbWard = ConvertNull(dx.Ward, "") & ""
13400         DoEvents
13410         DoEvents
13420         cmbClinician = dx.Clinician
13430         cmbGP = dx.GP
13440         cClDetails = dx.ClDetails
13450         If IsDate(dx.SampleDate) Then
13460             dtSampleDate = Format$(dx.SampleDate, "dd/mm/yyyy")
13470             If Format$(dx.SampleDate, "hh:mm") <> "00:00" Then
13480                 tSampleTime = Format$(dx.SampleDate, "hh:mm")
13490             Else
13500                 tSampleTime.Mask = ""
13510                 tSampleTime.Text = ""
13520                 tSampleTime.Mask = "##:##"
13530             End If
13540         Else
13550             dtSampleDate = Format$(Now, "dd/mm/yyyy")
13560             tSampleTime.Mask = ""
13570             tSampleTime.Text = ""
13580             tSampleTime.Mask = "##:##"
13590         End If
13600         lblSampleDate = dtSampleDate
13610         If dx.Fasting Then
13620             lRandom = "Fasting Sample"
13630             chkFasting.Value = vbChecked
13640         Else
13650             chkFasting.Value = vbUnchecked
13660         End If
13670         If dx.Valid Then
13680             lblDemogValid = "Demographics Valid"
13690             lblDemogValid.BackColor = &H80FF80
13700             lblDemogValid.ForeColor = vbBlack
13710             cmdValidateDemographics.Visible = True

13720             EnableDemographicEntry False
13730         Else
13740             lblDemogValid = "Demographics Not Valid"
13750             lblDemogValid.BackColor = vbRed
13760             lblDemogValid.ForeColor = vbYellow
13770             cmdValidateDemographics.Visible = False
13780             EnableDemographicEntry True
13790         End If
13800         If IsDate(dx.RecDate) Then
13810             dtRecDate = Format$(dx.RecDate, "dd/mm/yyyy")
13820             If Format$(dx.RecDate, "hh:mm") <> "00:00" Then
13830                 tRecTime = Format$(dx.RecDate, "hh:mm")
13840             Else
13850                 tRecTime.Mask = ""
13860                 tRecTime.Text = ""
13870                 tRecTime.Mask = "##:##"
13880             End If
13890         Else
13900             dtRecDate = dtSampleDate
13910             tRecTime.Mask = ""
13920             tRecTime.Text = ""
13930             tRecTime.Mask = "##:##"
13940         End If
13950         If frmOptUrgent Then
13960             If dx.Urgent Then
13970                 lblUrgent.Visible = True
13980                 chkUrgent.Value = 1
13990                 UrgentTest = True
14000             Else
14010                 chkUrgent.Value = 0
14020                 UrgentTest = False
14030             End If
14040         End If
14050         If dx.Urgent Then
14060             lblUrgent.Visible = True
14070             chkUrgent.Value = 1
14080             UrgentTest = True
14090         Else
14100             chkUrgent.Value = 0
14110             UrgentTest = False
14120         End If
            
              'Abubaker+++ 29-11-2023 (cmb ward and lbl name not loading on first try issue )
14130         DoEvents
14140         DoEvents
14150         cmbWard.Text = ConvertNull(dx.Ward, "") & ""
14160         lblName.Caption = Trim$(txtSurName.Text & " " & txtForeName.Text)
              'Abubaker--- 29-11-2023
          
14170     End If
14180     cmdSaveHoldDemographics.Enabled = False
14190     cmdSaveDemographics.Enabled = False

14200     If frmOptBloodBank Then
14210         If Trim$(txtChart) <> "" Then
14220             sql = "Select  * from PatientDetails where " & _
                      "PatNum = '" & txtChart & "'"
14230             Set tb = New Recordset
14240             RecOpenClientBB 0, tb, sql
14250             bViewBB.Enabled = Not tb.EOF
14260         End If
14270     End If
14280     SetViewScans txtSampleID, cmdViewScan
14290     Screen.MousePointer = 0



14300     Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

14310     intEL = Erl
14320     strES = Err.Description
14330     LogError "frmEditAll", "LoadDemographics", intEL, strES, sql

End Sub
Public Sub ClearDemographics()
14340     On Error GoTo ErrorHandler

14350     lblUrgent.Visible = False
14360     mNewRecord = True
14370     dtRunDate = Format$(Now, "dd/mm/yyyy")
          '40    lblRunDate = dtRunDate
14380     dtSampleDate = Format$(Now, "dd/mm/yyyy")
14390     lblSampleDate = dtSampleDate
14400     dtRecDate = Format$(Now, "dd/mm/yyyy")
14410     lblDemogValid = "Demographics Not Valid"
14420     lblDemogValid.BackColor = vbRed
14430     lblDemogValid.ForeColor = vbYellow
14440     cmdValidateDemographics.Visible = False

14450     txtChart = ""
14460     txtSurName = ""
14470     txtForeName = ""
14480     txtAddress(0) = ""
14490     txtAddress(1) = ""
14500     StatusBar1.Panels(4).Text = ""
14510     txtSex = ""
14520     txtDoB = ""
14530     txtAge = ""
14540     lblDoB = ""
14550     lblAge = ""
14560     lblSex = ""
14570     cmbWard = "GP"
14580     cmbClinician = ""
14590     cmbGP = ""
14600     cClDetails = ""
14610     txtDemographicComment = ""
14620     tSampleTime.Mask = ""
14630     tSampleTime.Text = ""
14640     tSampleTime.Mask = "##:##"
14650     tRecTime.Mask = ""
14660     tRecTime.Text = ""
14670     tRecTime.Mask = "##:##"
14680     lblChartNumber.Caption = HospName(0) & " Chart #"
14690     lblChartNumber.BackColor = &H8000000F
14700     lblChartNumber.ForeColor = vbBlack
14710     EnableDemographicEntry True
14720     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub
Private Sub EnableDemographicEntry(ByVal Enable As Boolean)

14730     fr(0).Enabled = Enable
14740     fr(1).Enabled = Enable
14750     fr(2).Enabled = Enable
14760     txtChart.Locked = Not Enable
14770     txtSurName.Locked = Not Enable
14780     txtForeName.Locked = Not Enable
14790     txtDoB.Locked = Not Enable
14800     txtAge.Locked = Not Enable
14810     txtSex.Locked = Not Enable

14820     If Enable = False Then
14830         StatusBar1.Panels(3).Text = "Demographics Validated"
14840         StatusBar1.Panels(3).Bevel = sbrInset
14850     Else
14860         StatusBar1.Panels(3).Text = "Check Demographics"
14870         StatusBar1.Panels(3).Bevel = sbrRaised
14880     End If

End Sub

Private Sub CopyCC(ByVal strPrevID As String)

          Dim sql As String
          Dim tb As Recordset
          Dim sn As Recordset

14890     On Error GoTo CopyCC_Error

14900     If Trim$(txtSampleID) = "" Then Exit Sub

          'get SendCopyTo details for previous sample id
14910     sql = "Select * from SendCopyTo where " & _
              "SampleID = '" & Val(strPrevID) & "'"
14920     Set tb = New Recordset
14930     RecOpenServer 0, tb, sql
14940     If Not tb.EOF Then    'Save data against current sample id
14950         sql = "Select * from SendCopyTo where " & _
                  "SampleID = '" & Val(txtSampleID) & "'"
14960         Set sn = New Recordset
14970         RecOpenServer 0, sn, sql
14980         If sn.EOF Then
14990             sn.AddNew
15000             sn!SampleID = Trim$(txtSampleID)
15010             sn!Ward = tb!Ward & ""
15020             sn!Clinician = tb!Clinician & ""
15030             sn!GP = tb!GP & ""
15040             sn!Device = tb!Device & ""
15050             sn!Destination = tb!Destination & ""
15060             sn.Update
15070         End If
15080     End If

15090     Exit Sub

CopyCC_Error:

          Dim strES As String
          Dim intEL As Integer

15100     intEL = Erl
15110     strES = Err.Description
15120     LogError "frmEditAll", "CopyCC", intEL, strES, sql
End Sub

Private Sub CheckCC()

          Dim sql As String
          Dim tb As Recordset

15130     On Error GoTo CheckCC_Error

15140     cmdCopyTo.Caption = "cc"
15150     cmdCopyTo.Font.Bold = False
15160     cmdCopyTo.BackColor = &H8000000F

15170     If Trim$(txtSampleID) = "" Then Exit Sub

15180     sql = "Select * from SendCopyTo where " & _
              " CONVERT(VARCHAR, SampleID ) = '" & Val(txtSampleID) & "'"
15190     Set tb = New Recordset
15200     RecOpenServer 0, tb, sql
15210     If Not tb.EOF Then
15220         cmdCopyTo.Caption = "++ cc ++"
15230         cmdCopyTo.Font.Bold = True
15240         cmdCopyTo.BackColor = &H8080FF
15250     End If

15260     Exit Sub

CheckCC_Error:

          Dim strES As String
          Dim intEL As Integer

15270     intEL = Erl
15280     strES = Err.Description
15290     LogError "frmEditAll", "CheckCC", intEL, strES, sql

End Sub

Public Sub LoadHaematology()

          Dim tb As Recordset
          Dim sn As Recordset
          Dim n As Integer
          Dim ip As String
          Dim e As String
          Dim PrevDate As String
          Dim PrevID As String
          Dim sql As String
          Dim PrevRBC As Single
          Dim PrevHgb As Single
          Dim PrevMCV As Single
          Dim PrevHct As Single
          Dim PrevRDWCV As Single
          Dim PrevRDWSD As Single
          Dim PrevMCH As Single
          Dim PrevMCHC As Single
          Dim Prevplt As Single
          Dim PrevMPV As Single
          Dim PrevPLCR As Single
          Dim PrevPdw As Single
          Dim PrevWBC As Single
          Dim PrevLymA As Single
          Dim PrevLymP As Single
          Dim PrevMonoA As Single
          Dim PrevMonoP As Single
          Dim PrevNeutA As Single
          Dim PrevNeutP As Single
          Dim PrevEosA As Single
          Dim PrevEosP As Single
          Dim PrevBasA As Single
          Dim PrevBasP As Single
          Dim DoB As String
          Dim ThisValid As Boolean
          Dim udtH As udtHaem
          Dim t As Single
          Dim SapphireErrors() As String
          Dim DeltaDate As String

15300     On Error GoTo LoadHaematology_Error

15310     ReDim i(0 To 6) As String

15320     SetViewReports "Haematology", txtSampleID
15330     t = Timer

15340     lblDateConflict.Visible = False

15350     lblIRF.Visible = False
15360     txtIRF.Visible = False

15370     bHaemGraphs.Visible = False
15380     bViewHaemRepeat.Visible = False
15390     PreviousHaem = False
15400     SSTab1.TabCaption(1) = "Haematology"
15410     lHaemErrors.Visible = False
15420     SampleProcessingError = ""

15430     lblHaemPrinted = ""
15440     Label1(36).Visible = False

15450     HaemValBy = ""
15460     lblAnalyser = ""

15470     If Trim$(txtSampleID) = "" Then Exit Sub

15480     DoB = txtDoB

          ''If Trim$(txtChart) <> "" Then
          ' If AddTicks(txtSurName & " " & txtForeName) & "' " <> "" Then
          '    DeltaDate = Format$(dtSampleDate, "dd/MMM/yyyy")
          '    If IsDate(tSampleTime) Then
          '        DeltaDate = DeltaDate & " " & tSampleTime
          '    End If
          ''    sql = "SELECT TOP 1 SampleID, RunDate FROM Demographics WHERE " & _
          '          "Chart = '" & txtChart & "' " & _
          '          "AND Hospital = '" & cmbHospital & "' " & _
          '          "AND SampleDate < '" & DeltaDate & "' " & _
          '          "ORDER BY SampleDate desc"
          '
          ''    sql = "SELECT TOP 1 SampleID, RunDate FROM Demographics WHERE " & _
          '          "Chart = '" & txtChart & "' " & _
          '          "AND Hospital = '" & cmbHospital & "' " & _
          '          " AND D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "' " & _
          '          " AND D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "' " & _
          '          "AND SampleDate < '" & DeltaDate & "' " & _
          '          "ORDER BY SampleDate desc"
          '
          '   sql = " SELECT     TOP (1) R.Rundate,D.SampleDate,R.sampleid " & _
          '             " FROM         demographics AS D INNER JOIN HaemResults AS R ON D.SampleID = R.sampleid  " & _
          '             " WHERE     (D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "') AND (D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "') " & _
          '             "AND Hospital = '" & cmbHospital & "' " & _
          '             "  ORDER BY D.SampleDate,D.sampleid ASC "
          '
          '    Set sn = New Recordset
          '    RecOpenServer 0, sn, sql
          '    If Not sn.EOF Then
          '        PrevDate = sn!Rundate
          '        PrevID = sn!SampleID
          '        sql = "Select * from HaemResults where " & _
          '              "SampleID = '" & PrevID & "'"
          '        Set tb = New Recordset
          '        RecOpenServer 0, tb, sql
          '        If Not tb.EOF Then
          '            PreviousHaem = True
          '            PrevRBC = Val(tb!rbc & "")
          '            PrevHgb = Val(tb!Hgb & "")
          '            PrevMCV = Val(tb!MCV & "")
          '            PrevHct = Val(tb!hct & "")
          '            PrevRDWCV = Val(tb!RDWCV & "")
          '            PrevRDWSD = Val(tb!rdwsd & "")
          '            PrevMCH = Val(tb!mch & "")
          '            PrevMCHC = Val(tb!mchc & "")
          '            Prevplt = Val(tb!plt & "")
          '            PrevMPV = Val(tb!mpv & "")
          '            PrevPLCR = Val(tb!plcr & "")
          '            PrevPdw = Val(tb!pdw & "")
          '            PrevWBC = Val(tb!WBC & "")
          '            PrevLymA = Val(tb!LymA & "")
          '            PrevLymP = Val(tb!LymP & "")
          '            PrevMonoA = Val(tb!MonoA & "")
          '            PrevMonoP = Val(tb!MonoP & "")
          '            PrevNeutA = Val(tb!NeutA & "")
          '            PrevNeutP = Val(tb!NeutP & "")
          '            PrevEosA = Val(tb!EosA & "")
          '            PrevEosP = Val(tb!EosP & "")
          '            PrevBasA = Val(tb!BasA & "")
          '            PrevBasP = Val(tb!BasP & "")
          '        End If
          '    End If
          'End If
          '  ShowHistory (SSTab1.Tab)
15490     ClearHaematologyResults

15500     sql = "Select * from HaemResults where " & _
              "SampleID = '" & Val(txtSampleID) & "'"
15510     Set tb = New Recordset
15520     RecOpenServer 0, tb, sql

15530     If ExtendedIPUFlagsAvailable() Then
15540         lHaemErrors.Visible = True
15550     End If

15560     If tb.EOF Then
15570         cmdValidateHaem.Enabled = False
15580     Else

15590         If Trim$(tb!AnalyserMessage & "") <> "" Then
15600             lHaemErrors.Visible = True
15610         End If
              '*************************************
              'BLR: if haemresults exist then remove haem request for that sampleid
15620         sql = "Delete From HaemRequests Where SampleID = " & Val(txtSampleID)
15630         Cnxn(0).Execute sql
              '*************************************
15640         Select Case Trim$(tb!Analyser & "")
                  Case "1": lblAnalyser = "A"
15650             Case "2": lblAnalyser = "B"
15660             Case Else: lblAnalyser = tb!Analyser & ""
15670         End Select

15680         If Not IsNull(tb!wbccomment) Then
15690             If InStr(tb!wbccomment, "^") > 0 Then
15700                 SapphireErrors = Split(tb!wbccomment, "^")
15710                 For n = 0 To UBound(SapphireErrors)
15720                     If Val(SapphireErrors(n)) > 0 Then
15730                         lHaemErrors.Visible = True
15740                     End If
15750                 Next
15760                 SampleProcessingError = tb!wbccomment
15770             End If
15780         End If

15790         If Not IsNull(tb!LongError) Then
15800             If Val(tb!LongError) > 1 Then
15810                 lHaemErrors.Visible = True
15820                 lHaemErrors.Tag = Format$(tb!LongError)
15830             End If
15840         End If

15850         If Not IsNull(tb!gwb1) Or Not IsNull(tb!gwb2) Or Not IsNull(tb!gRBC) Or Not IsNull(tb!gplt) Or Not IsNull(tb!gplth) Then
15860             bHaemGraphs.Visible = True
15870         End If

15880         pdelta.Cls

15890         lblWVF = tb!WVF & ""

15900         lWIC = tb!WIC & ""
15910         lWOC = tb!WOC & ""
15920         cFilm = 0
15930         If Not IsNull(tb!cFilm) Then
15940             cFilm = IIf(tb!cFilm, 1, 0)
15950         End If
15960         t = Timer
15970         If Not IsNull(tb!rbc) Then
15980             CheckIfMustPhone "Haematology", "RBC", tb!rbc
15990             udtH = GetHaemInfo("RBC", txtSex, DoB)
16000             ColouriseHaem tRBC, tb!rbc, udtH
                  'If PreviousHaem Then DeltaCheckHaem "RBC", tb!rbc, FindHaemPrvTestValue("RBC"), PrevDateHaem, PrevIDHaem, udtH
16010             DeltaCheckHaem "RBC", tb!rbc, FindHaemPrvTestValue("RBC", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16020         End If

16030         If Not IsNull(tb!Hgb) Then
16040             CheckIfMustPhone "Haematology", "Hgb", tb!Hgb
16050             udtH = GetHaemInfo("Hgb", txtSex, DoB)
16060             ColouriseHaem tHgb, tb!Hgb, udtH
16070             DeltaCheckHaem "Hgb", tb!Hgb, FindHaemPrvTestValue("Hgb", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16080         End If

16090         If Not IsNull(tb!MCV) Then
16100             CheckIfMustPhone "Haematology", "MCV", tb!MCV
16110             udtH = GetHaemInfo("MCV", txtSex, DoB)
16120             ColouriseHaem tMCV, tb!MCV, udtH
16130             DeltaCheckHaem "MCV", tb!MCV, FindHaemPrvTestValue("MCV", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16140         End If

16150         If Not IsNull(tb!hct) Then
16160             CheckIfMustPhone "Haematology", "Hct", tb!hct
16170             udtH = GetHaemInfo("Hct", txtSex, DoB)
16180             ColouriseHaem tHct, tb!hct, udtH
16190             DeltaCheckHaem "Hct", tb!hct, FindHaemPrvTestValue("Hct", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16200         End If

16210         If Not IsNull(tb!RDWCV) Then
16220             CheckIfMustPhone "Haematology", "RDWCV", tb!RDWCV
16230             udtH = GetHaemInfo("RDWCV", txtSex, DoB)
16240             ColouriseHaem tRDWCV, tb!RDWCV, udtH
16250             DeltaCheckHaem "RDWCV", tb!RDWCV, FindHaemPrvTestValue("RDWCV", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16260         End If

16270         If Not IsNull(tb!rdwsd) Then
16280             CheckIfMustPhone "Haematology", "RDWSD", tb!rdwsd
16290             udtH = GetHaemInfo("RDWSD", txtSex, DoB)
16300             ColouriseHaem tRDWSD, tb!rdwsd, udtH
16310             DeltaCheckHaem "RDWSD", tb!rdwsd, FindHaemPrvTestValue("RDWSD", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16320         End If

16330         If Not IsNull(tb!mch) Then
16340             CheckIfMustPhone "Haematology", "MCH", tb!mch
16350             udtH = GetHaemInfo("MCH", txtSex, DoB)
16360             ColouriseHaem tMCH, tb!mch, udtH
16370             DeltaCheckHaem "MCH", tb!mch, FindHaemPrvTestValue("MCH", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16380         End If

16390         If Not IsNull(tb!mchc) Then
16400             CheckIfMustPhone "Haematology", "MCHC", tb!mchc
16410             udtH = GetHaemInfo("MCHC", txtSex, DoB)
16420             ColouriseHaem tMCHC, tb!mchc, udtH
16430             DeltaCheckHaem "MCHC", tb!mchc, FindHaemPrvTestValue("MCHC", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16440         End If

16450         If Not IsNull(tb!plt) Then
16460             CheckIfMustPhone "Haematology", "Plt", tb!plt
16470             udtH = GetHaemInfo("Plt", txtSex, DoB)
16480             ColouriseHaem tPlt, tb!plt, udtH
16490             DeltaCheckHaem "Plt", tb!plt, FindHaemPrvTestValue("Plt", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16500         End If

16510         If Not IsNull(tb!mpv) Then
16520             CheckIfMustPhone "Haematology", "MPV", tb!mpv
16530             udtH = GetHaemInfo("MPV", txtSex, DoB)
16540             ColouriseHaem tMPV, tb!mpv, udtH
16550             DeltaCheckHaem "MPV", tb!mpv, FindHaemPrvTestValue("MPV", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16560         End If

16570         If Not IsNull(tb!WBC) Then
16580             CheckIfMustPhone "Haematology", "WBC", tb!WBC
16590             udtH = GetHaemInfo("WBC", txtSex, DoB)
16600             ColouriseHaem tWBC, tb!WBC, udtH
                  '        PrevWBC = FindHaemPrvTestValue("WBC")
16610             DeltaCheckHaem "WBC", tb!WBC, FindHaemPrvTestValue("WBC", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16620         End If

16630         If Not IsNull(tb!LymA) Then
16640             CheckIfMustPhone "Haematology", "LymA", tb!LymA
16650             udtH = GetHaemInfo("LymA", txtSex, DoB)
16660             ColouriseHaem tLymA, tb!LymA, udtH
16670             DeltaCheckHaem "LymA", tb!LymA, FindHaemPrvTestValue("LymA", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16680         End If

16690         If Not IsNull(tb!LymP) Then
16700             CheckIfMustPhone "Haematology", "LymP", tb!LymP
16710             udtH = GetHaemInfo("LymP", txtSex, DoB)
16720             ColouriseHaem tLymP, tb!LymP, udtH
16730             DeltaCheckHaem "LymP", tb!LymP, FindHaemPrvTestValue("LymP", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16740         End If

16750         If Not IsNull(tb!MonoA) Then
16760             CheckIfMustPhone "Haematology", "MonoA", tb!MonoA
16770             udtH = GetHaemInfo("MonoA", txtSex, DoB)
16780             ColouriseHaem tMonoA, tb!MonoA, udtH
16790             DeltaCheckHaem "MonoA", tb!MonoA, FindHaemPrvTestValue("MonoA", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16800         End If

16810         If Not IsNull(tb!MonoP) Then
16820             CheckIfMustPhone "Haematology", "MonoP", tb!MonoP
16830             udtH = GetHaemInfo("MonoP", txtSex, DoB)
16840             ColouriseHaem tMonoP, tb!MonoP, udtH
16850             DeltaCheckHaem "MonoP", tb!MonoP, FindHaemPrvTestValue("MonoP", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16860         End If

16870         If Not IsNull(tb!NeutA) Then
16880             CheckIfMustPhone "Haematology", "NeutA", tb!NeutA
16890             udtH = GetHaemInfo("NeutA", txtSex, DoB)
16900             ColouriseHaem tNeutA, tb!NeutA, udtH
16910             If PreviousHaem Then DeltaCheckHaem "NeutA", tb!NeutA, FindHaemPrvTestValue("NeutA", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16920         End If

16930         If Not IsNull(tb!NeutP) Then
16940             CheckIfMustPhone "Haematology", "NeutP", tb!NeutP
16950             udtH = GetHaemInfo("NeutP", txtSex, DoB)
16960             ColouriseHaem tNeutP, tb!NeutP, udtH
16970             DeltaCheckHaem "NeutP", tb!NeutP, FindHaemPrvTestValue("NeutP", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
16980         End If

16990         If Not IsNull(tb!EosA) Then
17000             CheckIfMustPhone "Haematology", "EosA", tb!EosA
17010             udtH = GetHaemInfo("EosA", txtSex, DoB)
17020             ColouriseHaem tEosA, tb!EosA, udtH
17030             DeltaCheckHaem "EosA", tb!EosA, FindHaemPrvTestValue("EosA", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
17040         End If

17050         If Not IsNull(tb!EosP) Then
17060             CheckIfMustPhone "Haematology", "EosP", tb!EosP
17070             udtH = GetHaemInfo("EosP", txtSex, DoB)
17080             ColouriseHaem tEosP, tb!EosP, udtH
17090             DeltaCheckHaem "EosP", tb!EosP, FindHaemPrvTestValue("EosP", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
17100         End If

17110         If Not IsNull(tb!BasA) Then
17120             CheckIfMustPhone "Haematology", "BasA", tb!BasA
17130             udtH = GetHaemInfo("BasA", txtSex, DoB)
17140             ColouriseHaem tBasA, tb!BasA, udtH
17150             DeltaCheckHaem "BasA", tb!BasA, FindHaemPrvTestValue("BasA", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
17160         End If

17170         If Not IsNull(tb!BasP) Then
17180             CheckIfMustPhone "Haematology", "BasP", tb!BasP
17190             udtH = GetHaemInfo("BasP", txtSex, DoB)
17200             ColouriseHaem tBasP, tb!BasP, udtH
17210             DeltaCheckHaem "BasP", tb!BasP, FindHaemPrvTestValue("BasP", udtH.DeltaDaysBackLimit), PrevDateHaem, PrevIDHaem, udtH
17220         End If

17230         If Trim$(tb!CD3A & tb!CD3P & tb!CD4A & tb!CD4P & tb!CD8A & tb!CD8P & tb!CD48 & "") <> "" Then
17240             pbCD.Visible = True
17250             txtCD3A.Text = tb!CD3A & ""
17260             txtCD3P.Text = tb!CD3P & ""
17270             txtCD4A.Text = tb!CD4A & ""
17280             txtCD4P.Text = tb!CD4P & ""
17290             txtCD8A.Text = tb!CD8A & ""
17300             txtCD8P.Text = tb!CD8P & ""
17310             txtCD48.Text = tb!CD48 & ""
17320         End If


17330         If Trim$(tb!IRF & "") <> "" Then
17340             lblIRF.Visible = True
17350             txtIRF.Visible = True
17360             txtIRF = tb!IRF
17370         End If

17380         tnrbcP = tb!nrbcP & ""
17390         tnrbcA = tb!nrbcA & ""

17400         If Not IsNull(tb!cMonospot) Then
17410             cMonospot = IIf(tb!cMonospot, 1, 0)
17420         Else
17430             cMonospot = 0
17440         End If
17450         Select Case Trim$(tb!MonoSpot & "")
                  Case "P": tMonospot = "Positive"
17460             Case "N": tMonospot = "Negative"
17470             Case Else:
17480                 If cMonospot = 0 Then
17490                     tMonospot = ""
17500                 Else
17510                     tMonospot = "?"
17520                 End If
17530         End Select

17540         If Not IsNull(tb!cESR) Then
17550             cESR = IIf(tb!cESR, 1, 0)
17560         Else
17570             cESR = 0
17580         End If
17590         If cESR = 1 And Trim$(tb!ESR & "") = "" Then
17600             tESR = "?"
17610         Else
17620             tESR = tb!ESR & ""
17630         End If

17640         If Not IsNull(tb!cRA) Then
17650             chkRA = IIf(tb!cRA, 1, 0)
17660         Else
17670             chkRA = 0
17680         End If
17690         Select Case Trim$(tb!RA & "")
                  Case "P": lblRA = "Positive"
17700             Case "N": lblRA = "Negative"
17710             Case "?": lblRA = "?"
17720             Case Else: lblRA = ""
17730         End Select

17740         If Not IsNull(tb!cRetics) Then
17750             cRetics = IIf(tb!cRetics, 1, 0)
17760         Else
17770             cRetics = 0
17780         End If
17790         If cRetics = 1 And Trim$(tb!RetA & "") = "" Then
17800             tRetA = "?"
17810         Else
17820             tRetA = tb!RetA & ""
17830         End If
17840         If cRetics = 1 And Trim$(tb!RetP & "") = "" Then
17850             tRetP = "?"
17860         Else
17870             tRetP = tb!RetP & ""
17880         End If

17890         If Not IsNull(tb!cMalaria) Then
17900             chkMalaria = IIf(tb!cMalaria, 1, 0)
17910         Else
17920             chkMalaria = 0
17930         End If
17940         If chkMalaria = 1 And Trim$(tb!Malaria & "") = "" Then
17950             lblMalaria = "?"
17960         Else
17970             lblMalaria = tb!Malaria & ""
17980         End If

17990         If Not IsNull(tb!cSickledex) Then
18000             chkSickledex = IIf(tb!cSickledex, 1, 0)
18010         Else
18020             chkSickledex = 0
18030         End If
18040         If chkSickledex = 1 And Trim$(tb!Sickledex & "") = "" Then
18050             lblSickledex = "?"
18060         Else
18070             lblSickledex = tb!Sickledex & ""
18080         End If

18090         tWarfarin = tb!Warfarin & ""

18100         ip = Left$(tb!ipmessage & "000000", 6)
18110         For n = 0 To 5
18120             ipflag(n).Enabled = Mid$(ip, n + 1, 1) = "1"
18130         Next

18140         e = tb!negposerror & ""

18150         buildinterp tb, i()
18160         For n = 0 To 6
18170             pdelta.Print i(n)
18180         Next

18190         ThisValid = False
18200         If Not IsNull(tb!Valid) Then
18210             ThisValid = IIf(tb!Valid, True, False)
18220         End If
18230         If ThisValid Then
18240             HaemValBy = tb!Operator & ""
18250         End If
18260         cmdValidateHaem.Enabled = Not ThisValid
18270         Label1(36).Visible = ThisValid
18280         If Not IsNull(tb!Printed) Then
18290             If tb!Printed Then
18300                 lblHaemPrinted = "Already Printed"
18310             Else
18320                 lblHaemPrinted = "Not Printed"
18330             End If
18340         Else
18350             lblHaemPrinted = "Not Printed"
18360         End If

18370         sql = "Select * from HaemRepeats where " & _
                  "SampleID = '" & Val(txtSampleID) & "'"
18380         Set tb = New Recordset
18390         RecOpenClient 0, tb, sql
18400         If Not tb.EOF Then
18410             bViewHaemRepeat.Visible = True
18420         End If

18430         SSTab1.TabCaption(1) = ">>Haematology<<"

18440         If Trim$(tESR) <> "" Then
18450             CheckRunSampleDates
18460         End If

18470     End If
18480     LoadOutstandingHaem
18490     cmdSaveHaem.Enabled = False
18500     Screen.MousePointer = 0
18510     Debug.Print "Load haem " & Timer - t

18520     Exit Sub

LoadHaematology_Error:

          Dim strES As String
          Dim intEL As Integer

18530     intEL = Erl
18540     strES = Err.Description
18550     LogError "frmEditAll", "LoadHaematology", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindHaemPrvTestValue
' Author    : XPMUser
' Date      : 06/Aug/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FindHaemPrvTestValue(AnalyteName As String, DeltaDaysBackLimit As Single) As Single

18560     On Error GoTo FindHaemPrvTestValue_Error
          Dim sql As String
          Dim sn As Recordset
          Dim DateQry As String


18570     PrevDateHaem = ""
18580     PrevIDHaem = ""
18590     PreviousHaem = False

18600     If IsDate(tSampleTime) Then
18610         DateQry = " AND (D.SampleDate < '" & Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm") & "')"
18620     Else
18630         DateQry = " AND (D.SampleDate < '" & Format$((dtSampleDate), "dd/MMM/yyyy") & "')"
18640     End If


18650     If AddTicks(txtSurName & " " & txtForeName) & "' " <> "" Then
18660         sql = " SELECT     TOP (1) R.Rundate,D.SampleDate,R.sampleid " & _
                  "," & AnalyteName & _
                  " FROM         demographics AS D INNER JOIN HaemResults AS R ON D.SampleID = R.sampleid  " & _
                  " WHERE     (D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "') AND (D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "') " & _
                  " AND (D.SampleDate >= '" & Format$((dtSampleDate - DeltaDaysBackLimit), "dd/MMM/yyyy") & "')  " & _
                  DateQry & _
                  " AND Hospital = '" & cmbHospital & "' " & _
                  " AND " & AnalyteName & " IS NOT NULL " & _
                  " ORDER BY D.SampleDate DESC,D.sampleid ASC "

18670         Set sn = New Recordset
18680         RecOpenServer 0, sn, sql
18690         If Not sn.EOF Then
18700             PrevDateHaem = sn!Rundate
18710             PrevIDHaem = sn!SampleID
18720             PreviousHaem = True

18730             FindHaemPrvTestValue = Val(sn("" & AnalyteName & "") & "")
18740         End If
18750     End If

18760     Exit Function


FindHaemPrvTestValue_Error:

          Dim strES As String
          Dim intEL As Integer

18770     intEL = Erl
18780     strES = Err.Description
18790     LogError "frmEditAll", "FindHaemPrvTestValue", intEL, strES, sql
End Function


Private Sub DeltaCheckHaem(ByVal Analyte As String, _
          ByVal Value As String, _
          ByVal PreviousValue As String, _
          ByVal PreviousDate As String, _
          ByVal PreviousID As String, _
          ByRef udtH As udtHaem)

18800     If Val(PreviousValue) = 0 Then Exit Sub

18810     If udtH.DoDelta Then
18820         If Abs(Val(PreviousValue) - Val(Value)) > Val(udtH.DeltaValue) Then
18830             pdelta.Print Left$(Format$(PreviousDate, "dd/mm/yyyy") & _
                      "(" & PreviousID & ") " & _
                      Analyte & ":" & Space(25), 25); PreviousValue
18840         End If
18850     End If

End Sub

Private Sub ColouriseHaem(ByVal Destination As TextBox, _
          ByVal strValue As String, _
          ByRef udtH As udtHaem)

          Dim Value As Single

18860     Value = Val(strValue)

18870     Destination.Text = strValue
18880     If Trim$(strValue) = "" Then
18890         Destination.BackColor = &HFFFFFF
18900         Destination.ForeColor = &H0&
18910         Exit Sub
18920     End If

18930     If Value < udtH.PlausibleLow Or Value > udtH.PlausibleHigh Then
18940         Destination.BackColor = vbBlack
18950         Destination.ForeColor = vbWhite
18960     ElseIf Value > udtH.High Then
18970         Destination.BackColor = &HFFFF&
18980         Destination.ForeColor = &HFF&
18990     ElseIf Value < udtH.Low Then
19000         Destination.BackColor = &HFFFF00
19010         Destination.ForeColor = &HC00000
19020     Else
19030         Destination.BackColor = &HFFFFFF
19040         Destination.ForeColor = &H0&
19050     End If

End Sub


Private Sub ClearHaematologyResults()

          Dim n As Integer

19060     ClearHaemExceptHgb

19070     tHgb = ""
19080     tHgb.BackColor = &HFFFFFF
19090     tHgb.ForeColor = &H0&

19100     pdelta.Cls

19110     cESR = 0
19120     cRetics = 0
19130     cMonospot = 0
19140     chkMalaria = 0
19150     chkSickledex = 0
19160     chkRA = 0
19170     tESR = ""
19180     tRetA = ""
19190     tRetP = ""
19200     tMonospot = ""
19210     lblMalaria = ""
19220     lblSickledex = ""
19230     lblRA = ""

19240     cFilm = 0

          'cCoag = 0

19250     tWarfarin = ""

19260     For n = 0 To 5
19270         ipflag(n).Visible = False
19280     Next

19290     txtCD3A.Text = ""
19300     txtCD4A.Text = ""
19310     txtCD8A.Text = ""
19320     txtCD3P.Text = ""
19330     txtCD4P.Text = ""
19340     txtCD8P.Text = ""
19350     txtCD48.Text = ""
19360     pbCD.Visible = False

19370     lblWVF.Caption = ""

End Sub


Private Sub FillcmbSampleType()

19380     On Error GoTo FillcmbSampleType_Error

19390     FillGenericList cmbSampleType, "ST"

19400     If cmbSampleType.ListCount > 0 Then
19410         cmbSampleType.ListIndex = 0
19420     End If

19430     Exit Sub

FillcmbSampleType_Error:

          Dim strES As String
          Dim intEL As Integer

19440     intEL = Erl
19450     strES = Err.Description
19460     LogError "frmEditAll", "FillcmbSampleType", intEL, strES

End Sub

Private Sub FillcmbAdd()

          Dim tb As Recordset
          Dim sql As String

19470     On Error GoTo FillcmbAdd_Error

19480     sql = "SELECT DISTINCT B.ShortName, B.PrintPriority, B.Code " & _
              "FROM BioTestDefinitions B, Lists L " & _
              "WHERE B.SampleType = L.Code " & _
              "AND L.ListType = 'ST' " & _
              "AND L.Text LIKE '" & cmbSampleType & "' " & _
              "AND B.InUse = 1 " & _
              "ORDER BY B.PrintPriority"
19490     Set tb = New Recordset
19500     RecOpenServer 0, tb, sql

19510     cmbAdd.Clear
19520     lstAdd.Clear
19530     Do While Not tb.EOF
19540         cmbAdd.AddItem tb!ShortName
19550         lstAdd.AddItem tb!Code
19560         tb.MoveNext
19570     Loop

19580     Exit Sub

FillcmbAdd_Error:

          Dim strES As String
          Dim intEL As Integer

19590     intEL = Erl
19600     strES = Err.Description
19610     LogError "frmEditAll", "FillcmbAdd", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

19620     pBar = 0

End Sub

Private Sub Form_Paint()
    '
    'Dim TabNumber As Integer
    '
    'If Activated Then Exit Sub
    '
    'Activated = True
    '
    'TabNumber = Val(GetSetting("NetAcquire", "StartUp", "LastDepartment", "0"))
    '
    'If SSTab1.TabVisible(TabNumber) Then
    '  SSTab1.Tab = TabNumber
    'End If
    '
    'Debug.Print SSTab1.Tab

End Sub

Private Sub Form_Unload(Cancel As Integer)

19630     If Val(txtSampleID) > Val(GetSetting("NetAcquire", "StartUp", "LastUsed", "1")) Then
19640         SaveSetting "NetAcquire", "StartUp", "LastUsed", txtSampleID
19650     End If
19660     Debug.Print SSTab1.Tab

19670     SaveSetting "NetAcquire", "StartUp", "LastDepartment", CStr(SSTab1.Tab)

19680     pPrintToPrinter = ""

19690     m_StartInDepartment = ""

19700     Activated = False

End Sub





Private Sub fr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

19710     pBar = 0

End Sub

Private Sub g_Click()

          Dim S As String
          Dim Prompt As String
          Dim x As Integer
          Dim y As Integer
          Dim ySave As Integer
19720     On Error GoTo ErrorHandler
19730     If g.MouseRow = 0 Then Exit Sub

19740     ySave = g.row

19750     For y = 1 To g.Rows - 1
19760         g.row = y
19770         For x = 1 To 5
19780             g.Col = x
19790             g.CellBackColor = 0
19800         Next
19810     Next

19820     g.row = ySave
19830     For x = 1 To 5
19840         g.Col = x
19850         g.CellBackColor = vbYellow
19860     Next


19870     If g.MouseCol = 1 Then
19880         Prompt = "Enter result for " & g.TextMatrix(g.row, 0)
19890         S = iBOX(Prompt, , g.TextMatrix(g.row, 1))
19900         If S <> "" Then
19910             g.TextMatrix(g.row, 1) = S
19920             g.TextMatrix(g.row, 5) = Format(Now, "dd/mmm/yyyy")
19930         End If
19940     End If
19950     cmdSaveExt.Enabled = True
19960     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub gBio_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

          Static S As String
          Dim tb As Recordset
          Dim sql As String
          Static SampleIDTest As String

19970     On Error GoTo gBio_MouseMove_Error

19980     pBar = 0

19990     If gBio.MouseRow = 0 Then
20000         S = ""
20010         SampleIDTest = ""
20020     ElseIf lblChartNumber.BackColor = vbRed Then
20030         S = ""
20040         SampleIDTest = ""
20050     ElseIf gBio.MouseCol = 0 Then
20060         If SampleIDTest <> Trim$(txtSampleID) & gBio.TextMatrix(gBio.MouseRow, 0) Then

20070             SampleIDTest = Trim$(txtSampleID) & gBio.TextMatrix(gBio.MouseRow, 0)
20080             If txtChart = "" Then
20090                 S = ""
20100             Else
20110                 sql = "select top 1 D.SampleID, D.RunDate, R.Result, B.DP " & _
                          "from Demographics as D, BioResults as R, BioTestDefinitions as B where " & _
                          "D.Chart = '" & txtChart & "' " & _
                          "and D.SampleID < " & Val(txtSampleID) & " " & _
                          "and D.SampleID = R.SampleID " & _
                          "and B.Code = R.Code " & _
                          "AND R.Code = '" & gBio.TextMatrix(gBio.MouseRow, 9) & "'" & _
                          "ORDER BY D.SampleID desc"
20120                 Set tb = New Recordset
20130                 RecOpenServer 0, tb, sql
20140                 If Not tb.EOF Then
20150                     S = Format$(tb!Rundate, "dd/mm/yy") & _
                              " (" & tb!SampleID & ") " & _
                              gBio.TextMatrix(gBio.MouseRow, 0) & ": "
20160                     If IsNumeric(tb!Result & "") Then
20170                         S = S & FormatNumber(tb!Result & "", tb!DP)
20180                     Else
20190                         S = S & tb!Result & ""
20200                     End If
20210                 Else
20220                     S = ""
20230                 End If
20240             End If
20250         End If
20260     ElseIf gBio.MouseCol = 4 Then
20270         Select Case gBio.TextMatrix(gBio.MouseRow, 4)
                  Case "L": S = "Low"
20280             Case "H": S = "High"
20290             Case "X": S = "Implausible"
20300             Case Else: S = ""
20310         End Select
20320     ElseIf gBio.MouseCol = 5 Then
20330         S = gBio.TextMatrix(gBio.MouseRow, 5)
20340         SampleIDTest = ""
20350     ElseIf gBio.MouseCol = 7 Then
20360         S = gBio.TextMatrix(gBio.MouseRow, 7)
20370         SampleIDTest = ""
20380     Else
20390         S = ""
20400         SampleIDTest = ""
20410     End If

20420     gBio.ToolTipText = S

20430     Exit Sub

gBio_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

20440     intEL = Erl
20450     strES = Err.Description
20460     LogError "frmEditAll", "gBio_MouseMove", intEL, strES, sql


End Sub

Private Function SavePrintInhibit() As Boolean
          'Returns True if there is something to print

          Dim sql As String
          Dim y As Integer
          Dim Discipline As String
          Dim g As MSFlexGrid

20470     On Error GoTo SavePrintInhibit_Error

20480     Discipline = ""

20490     Select Case SSTab1.Tab
              Case 2: SavePrintInhibit = False: Discipline = "Bio": Set g = gBio: g.Col = 8
20500         Case 3: SavePrintInhibit = False: Discipline = "Coa": Set g = gCoag: g.Col = 5
20510         Case Else: SavePrintInhibit = True: Discipline = ""
20520     End Select

20530     If Discipline = "" Then Exit Function

20540     sql = "DELETE FROM PrintInhibit WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "AND Discipline = '" & Discipline & "' " & _
              "INSERT INTO PrintInhibit " & _
              "(SampleID, Discipline, Parameter) " & _
              "SELECT DISTINCT '" & txtSampleID & "', '" & Discipline & "', D.ShortName " & _
              "FROM " & Discipline & "TestDefinitions D " & _
              "JOIN " & Discipline & "Results R " & _
              "ON D.Code = R.Code " & _
              "WHERE R.SampleID = '" & txtSampleID & "'"
20550     Cnxn(0).Execute sql

20560     For y = 1 To g.Rows - 1
20570         g.row = y
20580         If g.CellPicture = imgGreenTick.Picture Then
20590             sql = "DELETE FROM PrintInhibit " & _
                      "WHERE SampleID = '" & txtSampleID & "' " & _
                      "AND Discipline = '" & Discipline & "' " & _
                      "AND Parameter = '" & g.TextMatrix(y, 0) & "'"
20600             Cnxn(0).Execute sql
20610             SavePrintInhibit = True
20620         End If
20630     Next

20640     Select Case SSTab1.Tab
              Case 2: If Len(Trim$(txtBioComment)) > 0 Then SavePrintInhibit = True    'Bio
20650         Case 3: If Len(Trim$(txtCoagComment)) > 0 Then SavePrintInhibit = True    'Coag
20660     End Select

20670     Exit Function

SavePrintInhibit_Error:

          Dim strES As String
          Dim intEL As Integer

20680     intEL = Erl
20690     strES = Err.Description
20700     LogError "frmEditAll", "SavePrintInhibit", intEL, strES, sql

End Function
Private Sub gBio_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

          Dim sql As String
          Dim AnalyteToRemove As String
          Dim AnalyteCodeToRemove As String
          Dim f As Form
          Dim n As Integer

20710     On Error GoTo gBio_MouseUp_Error

20720     If gBio.MouseRow = 0 Then Exit Sub
20730     If gBio.TextMatrix(gBio.row, 0) = "" Then Exit Sub

20740     gBio.row = gBio.MouseRow

20750     If gBio.MouseCol = 7 Then
20760         Set f = New frmComment
20770         f.Comment = gBio
20780         f.Show 1
20790         sql = "UPDATE BioResults " & _
                  "SET Comment = '" & f.Comment & "' " & _
                  "WHERE SampleID = '" & txtSampleID & "' " & _
                  "AND Code = '" & gBio.TextMatrix(gBio.row, 9) & "'"
20800         Cnxn(0).Execute sql
20810         Unload f
20820         Set f = Nothing
20830         LoadBiochemistry
20840         Exit Sub
20850     ElseIf gBio.MouseCol = 8 Then
20860         If gBio.CellPicture = imgGreenTick.Picture Then
20870             Set gBio.CellPicture = imgRedCross.Picture
20880         Else
20890             Set gBio.CellPicture = imgGreenTick.Picture
20900         End If
20910         Exit Sub
20920     End If

20930     If gBio.TextMatrix(gBio.row, 0) = "HbA1c" Then
20940         frmViewFullData.SampleID = txtSampleID
20950         frmViewFullData.Show 1
20960         LoadBiochemistry
20970         Exit Sub
20980     End If

          Dim fr As New frmRemoveBioTest
20990     With fr
21000         .Discipline = "Bio"
21010         .Analyte = gBio.TextMatrix(gBio.row, 0)
21020         .AnalyteCode = gBio.TextMatrix(gBio.row, 9)
21030         .SampleID = Val(txtSampleID)
21040         .Show vbModal
21050         AnalyteCodeToRemove = .AnalyteCode
21060         AnalyteToRemove = .Analyte
21070     End With
21080     Unload fr
21090     Set fr = Nothing

21100     If AnalyteToRemove = "All" Then
21110         RemoveFromPhoneAlert txtSampleID, "Biochemistry", "All"
21120         sql = "Delete from BioResults where " & _
                  "SampleID = '" & txtSampleID & "'"
21130         Cnxn(0).Execute sql
21140     ElseIf AnalyteCodeToRemove = gBio.TextMatrix(gBio.row, 9) Then
21150         cmbNewResult = gBio.TextMatrix(gBio.row, 1)
21160         cmbUnits = gBio.TextMatrix(gBio.row, 3)
21170         For n = 0 To cmbAdd.ListCount - 1
21180             If AnalyteToRemove = cmbAdd.List(n) Then
21190                 cmbAdd.ListIndex = n
21200                 lstAdd.ListIndex = n
21210                 Exit For
21220             End If
21230         Next
21240         RemoveFromPhoneAlert txtSampleID, "Biochemistry", AnalyteToRemove

21250         sql = "DELETE FROM BioResults WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Code = '" & AnalyteCodeToRemove & "'"
21260         Cnxn(0).Execute sql
21270         cmbNewResult.SetFocus
21280     End If

21290     If AnalyteCodeToRemove = GetOptionSetting("BioCodeForCreatinine", "") Then
21300         sql = "DELETE FROM BioResults WHERE " & _
                  "Code = '" & GetOptionSetting("BioCodeForEGFR", "") & "'"
21310         Cnxn(0).Execute sql
21320     End If

21330     LoadBiochemistry

21340     Exit Sub

gBio_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

21350     intEL = Erl
21360     strES = Err.Description
21370     LogError "frmEditAll", "gBio_MouseUp", intEL, strES, sql

End Sub


Private Sub gCoag_Click()

          Dim sql As String

21380     On Error GoTo gCoag_Click_Error

21390     With gCoag

21400         If .MouseRow = 0 Then Exit Sub

21410         If .TextMatrix(.row, 0) = "" Then Exit Sub

21420         Select Case .Col

                  Case 0:
21430                 If iMsg("Delete " & .Text & "?", vbQuestion + vbYesNo) = vbYes Then
21440                     cParameter = .Text
21450                     tResult = .TextMatrix(.row, 1)
21460                     If .Rows = 2 Then
21470                         .AddItem ""
21480                         .RemoveItem 1
21490                     Else
21500                         .RemoveItem .row
21510                     End If
21520                     sql = "Delete from CoagResults where " & _
                              "SampleID = '" & txtSampleID & "' " & _
                              "and Code in " & _
                              "( Select Code from CoagTestDefinitions where " & _
                              "  TestName = '" & cParameter & "')"

21530                     cmdSaveCoag.Enabled = True
21540                     Cnxn(0).Execute sql
21550                 End If

21560             Case 1:
21570                 .Text = iBOX("Enter new Value for " & .TextMatrix(.row, 0), , .Text)
21580                 cmdSaveCoag.Enabled = True

21590             Case 4:
21600                 .Text = IIf(.Text = "", "V", "")
21610                 cmdSaveCoag.Enabled = True

21620             Case 5:
21630                 If .CellPicture = imgGreenTick.Picture Then
21640                     Set .CellPicture = imgRedCross.Picture
21650                 Else
21660                     Set .CellPicture = imgGreenTick.Picture
21670                 End If

21680         End Select

21690     End With

21700     Exit Sub

gCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

21710     intEL = Erl
21720     strES = Err.Description
21730     LogError "frmEditAll", "gCoag_Click", intEL, strES, sql


End Sub

Private Sub gCoag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

21740     pBar = 0

End Sub

Private Sub gOutstandingCoag_Click()

          Dim tb As Recordset
          Dim sql As String

21750     On Error GoTo gOutstandingCoag_Click_Error

21760     With gOutstandingCoag
21770         If .MouseRow = 0 Then Exit Sub
21780         If .Text = "" Then Exit Sub
21790         If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
21800             sql = "Delete from CoagRequests where " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "and code in " & _
                      "( Select Code from CoagTestDefinitions where " & _
                      "  TestName = '" & .Text & "')"
21810             Set tb = New Recordset
21820             RecOpenClient 0, tb, sql
21830             If .Rows > 2 Then
21840                 .RemoveItem .row
21850             Else
21860                 .AddItem ""
21870                 .RemoveItem 1
21880             End If
21890         End If
21900     End With

21910     Exit Sub

gOutstandingCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

21920     intEL = Erl
21930     strES = Err.Description
21940     LogError "frmEditAll", "gOutstandingCoag_Click", intEL, strES, sql


End Sub

Private Sub grdOutstanding_Click()

          Dim sql As String

21950     On Error GoTo grdOutstanding_Click_Error

21960     With grdOutstanding
21970         If .MouseRow = 0 Then Exit Sub
21980         .row = .MouseRow
21990         If .Text = "" Then Exit Sub
22000         If iMsg("Remove " & .TextMatrix(.row, 0) & " from Requests?", vbQuestion + vbYesNo) = vbYes Then

22010             sql = "DELETE FROM BioRequests " & _
                      "WHERE SampleID = '" & txtSampleID & "' " & _
                      "AND Code = '" & .TextMatrix(.row, 1) & "'"
22020             Cnxn(0).Execute sql

22030             If .Rows > 2 Then
22040                 .RemoveItem .row
22050             Else
22060                 .AddItem ""
22070                 .RemoveItem 1
22080             End If

22090         End If

22100     End With

22110     Exit Sub

grdOutstanding_Click_Error:

          Dim strES As String
          Dim intEL As Integer

22120     intEL = Erl
22130     strES = Err.Description
22140     LogError "frmEditAll", "grdOutstanding_Click", intEL, strES, sql

End Sub

Private Sub grdOutstanding_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

22150     pBar = 0

End Sub

Private Sub imgLast_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim strDept As String
          Dim strSplitSelect As String

22160     On Error GoTo imgLast_Click_Error

22170     If Val(txtSampleID) > 100000000 Then
22180         txtSampleID = "1"
22190     End If

22200     Select Case SSTab1.Tab
              Case 0:
22210             txtSampleID = Format$(Val(txtSampleID) + 1)
22220             Debug.Print "imglast_click 70"
22230             LoadAllDetails

22240             cmdSaveHoldDemographics.Enabled = False
22250             cmdSaveDemographics.Enabled = False
22260             cmdSaveHaem.Enabled = False
22270             cmdSaveBio.Enabled = False
22280             cmdSaveCoag.Enabled = False
22290             Exit Sub

22300         Case 1: strDept = "Haem"
22310         Case 2: strDept = "Bio"
22320         Case 3: strDept = "Coag"
22330         Case 6: strDept = "Ext"
22340     End Select

22350     sql = "Select top 1 SampleID from " & strDept & "Results " & _
              "WHERE SampleID < 1000000000 Order by SampleID desc"

22360     If strDept = "Bio" And lblSplitView.Tag <> "0" Then
22370         strSplitSelect = LoadSplitList(Val(lblSplitView.Tag))
22380         If strSplitSelect <> "" Then
22390             sql = "Select top 1 SampleID from " & strDept & "Results where " & _
                      "(" & strSplitSelect & ") " & _
                      "AND SampleID < 1000000000 Order by SampleID desc"
22400         End If
22410     End If

22420     Set tb = New Recordset
22430     RecOpenClient 0, tb, sql
22440     If Not tb.EOF Then
22450         txtSampleID = tb!SampleID & ""
22460     End If

22470     If Val(txtSampleID) > 1000000000 Then
22480         txtSampleID = "1"
22490     End If

22500     Debug.Print "imglast 340"
22510     LoadAllDetails

22520     cmdSaveHoldDemographics.Enabled = False
22530     cmdSaveDemographics.Enabled = False
22540     cmdSaveHaem.Enabled = False
22550     cmdSaveBio.Enabled = False
22560     cmdSaveCoag.Enabled = False
22570     cmdSaveExt.Enabled = False

22580     Exit Sub

imgLast_Click_Error:

          Dim strES As String
          Dim intEL As Integer

22590     intEL = Erl
22600     strES = Err.Description
22610     LogError "frmEditAll", "imgLast_Click", intEL, strES, sql

End Sub

Private Sub iRecDate_Click(Index As Integer)

22620     If Index = 0 Then
22630         dtRecDate = DateAdd("d", -1, dtRecDate)
22640     Else
22650         If DateDiff("d", dtRecDate, Now) > 0 Then
22660             dtRecDate = DateAdd("d", 1, dtRecDate)
22670         End If
22680     End If

22690     cmdSaveDemographics.Enabled = True
22700     cmdSaveHoldDemographics.Enabled = True

End Sub

Private Sub irelevant_Click(Index As Integer)

22710     On Error GoTo irelevant_Click_Error

          'txtLabNo = Val(FndMaxID("demographics", "LabNo", "")) + 1
22720     MoveToNextRelevant Index

22730     Exit Sub

irelevant_Click_Error:

          Dim strES As String
          Dim intEL As Integer

22740     intEL = Erl
22750     strES = Err.Description
22760     LogError "frmEditAll", "irelevant_Click", intEL, strES


End Sub

Private Sub iRunDate_Click(Index As Integer)

22770     If Index = 0 Then
22780         dtRunDate = DateAdd("d", -1, dtRunDate)
22790         If DateDiff("d", dtRunDate, dtSampleDate) > 0 Then
22800             dtRunDate = dtSampleDate
22810         End If
22820     Else
22830         If DateDiff("d", dtRunDate, Now) > 0 Then
22840             dtRunDate = DateAdd("d", 1, dtRunDate)
22850         End If
22860     End If

22870     cmdSaveDemographics.Enabled = True
22880     cmdSaveHoldDemographics.Enabled = True

End Sub

Private Sub iSampleDate_Click(Index As Integer)

22890     If Index = 0 Then
22900         dtSampleDate = DateAdd("d", -1, dtSampleDate)
22910     Else
22920         If DateDiff("d", dtSampleDate, Now) > 0 Then
22930             dtSampleDate = DateAdd("d", 1, dtSampleDate)
22940         End If
22950     End If

22960     cmdSaveDemographics.Enabled = True
22970     cmdSaveHoldDemographics.Enabled = True

End Sub


Private Sub iToday_Click(Index As Integer)

22980     If Index = 0 Then
22990         dtRunDate = Format$(Now, "dd/mm/yyyy")
23000     ElseIf Index = 1 Then
23010         If DateDiff("d", dtRunDate, Now) > 0 Then
23020             dtSampleDate = dtRunDate
23030         Else
23040             dtSampleDate = Format$(Now, "dd/mm/yyyy")
23050         End If
23060     ElseIf Index = 2 Then
23070         If DateDiff("d", dtRunDate, Now) > 0 Then
23080             dtRecDate = dtRunDate
23090         Else
23100             dtRecDate = Format$(Now, "dd/mm/yyyy")
23110         End If
23120     End If

23130     cmdSaveDemographics.Enabled = True
23140     cmdSaveHoldDemographics.Enabled = True

End Sub




Private Sub lblResultOrRequest_Click()

23150     If SSTab1.Tab <> 0 Then
23160         Select Case lblResultOrRequest
                  Case "Results": lblResultOrRequest = "Request"
23170             Case "Request": lblResultOrRequest = "Not Val"
23180             Case "Not Val": lblResultOrRequest = "Results"
23190         End Select
23200     End If

End Sub

Private Sub lblHaemolysed_Click()

23210     With lblHaemolysed()
23220         Select Case .Caption
                  Case "", "0": .Caption = "1+"
23230             Case "1+": .Caption = "2+"
23240             Case "2+": .Caption = "3+"
23250             Case "3+": .Caption = "4+"
23260             Case "4+": .Caption = "5+"
23270             Case "5+": .Caption = "6+"
23280             Case "6+": .Caption = ""
23290         End Select
23300     End With

23310     cmdSaveBio.Enabled = True

End Sub

Private Sub lblSplit_Click(Index As Integer)

          Dim n As Integer

23320     For n = 0 To 6
23330         lblSplit(n).BackColor = SSTab1.BackColor
23340         lblSplit(n).ForeColor = vbBlack
23350     Next

23360     If Index = 0 Then
23370         lblSplitView.BackColor = SSTab1.BackColor
23380         lblSplitView.FontBold = False
23390         lblSplitView.ForeColor = vbBlack
23400     Else
23410         lblSplitView.BackColor = vbRed
23420         lblSplitView.FontBold = True
23430         lblSplitView.ForeColor = vbYellow

23440         lblSplit(Index).BackColor = vbRed
23450         lblSplit(Index).ForeColor = vbYellow

23460     End If

23470     lblSplitView.Caption = "Viewing " & lblSplit(Index).Caption
23480     lblSplitView.Tag = Format$(Index)

23490     SaveOptionSetting "LastUsedBioSplitName", lblSplit(Index).Caption
23500     SaveOptionSetting "LastUsedBioSplitIndex", Format$(Index)

23510     LoadBiochemistry

End Sub

Private Sub lblUrgent_Click()

          Dim sql As String

23520     On Error GoTo lblUrgent_Click_Error

23530     sql = "Update Demographics " & _
              "Set Urgent = 0 where " & _
              "SampleID = '" & txtSampleID & "'"
23540     Cnxn(0).Execute sql

23550     lblUrgent.Visible = False

23560     Exit Sub

lblUrgent_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23570     intEL = Erl
23580     strES = Err.Description
23590     LogError "frmEditAll", "lblUrgent_Click", intEL, strES, sql


End Sub

Private Sub lHaemErrors_Click()

23600     With frmHaemErrors
23610         .SampleID = Val(txtSampleID)
23620         .ErrorNumber = lHaemErrors.Tag
23630         .Show 1
23640     End With

End Sub

Private Sub lRandom_Click()

23650     If lRandom = "Random Sample" Then
23660         lRandom = "Fasting Sample"
23670     Else
23680         lRandom = "Random Sample"
23690     End If

23700     LoadBiochemistry

23710     cmdSaveBio.Enabled = True

End Sub

Private Sub mnuAutoVal_Click()

23720     With frmTestAutoValidate
23730         .Discipline = "Bio"
23740         .SampleType = "S"
23750         .Show 1
23760     End With

End Sub

Private Sub mnuDelta_Click()

23770     With frmTestDelta
23780         .Discipline = "Bio"
23790         .SampleType = "S"
23800         .Show 1
23810     End With

End Sub

Private Sub mnuInUse_Click()

23820     With frmTestInUse
23830         .Discipline = "Bio"
23840         .SampleType = "S"
23850         .Show 1
23860     End With

End Sub

Private Sub mnuKnownToAnalyser_Click()

23870     With frmTestKnownToAnalyser
23880         .Discipline = "Bio"
23890         .SampleType = "S"
23900         .Show 1
23910     End With

End Sub


Private Sub mnuMasks_Click()

23920     With frmTestMasks
23930         .Discipline = "Bio"
23940         .Show 1
23950     End With

End Sub

Private Sub mnuNewResult_Click()

23960     With frmListsGeneric
23970         .ListType = "NewResult"
23980         .ListTypeNames = "New Results"
23990         .ListTypeName = "New Result"
24000         .Show 1
24010     End With

24020     FillGenericList cmbNewResult, "NewResult"

End Sub

Private Sub mnuNormalFlag_Click()

24030     With frmTestNormalRange
24040         .Discipline = "Bio"
24050         .SampleType = "S"
24060         .Show 1
24070     End With

End Sub

Private Sub mnuPlausible_Click()

24080     With frmTestPlausible
24090         .Discipline = "Bio"
24100         .SampleType = "S"
24110         .Show 1
24120     End With

End Sub

Private Sub mnuSplits_Click()

          Dim n As Integer

24130     frmBioSplitList.Show 1

24140     For n = 1 To 6
24150         lblSplit(n).Caption = GetOptionSetting("BioSplitName" & Format$(n), "Split " & Format$(n))
24160     Next

End Sub



Private Sub mnuTestCodeMappingBio_Click()

24170     On Error GoTo mnuTestCodeMappingBio_Click_Error

24180     frmTestCodeMapping.Discipline = "BIO"
24190     frmTestCodeMapping.Show 1

24200     Exit Sub

mnuTestCodeMappingBio_Click_Error:

          Dim strES As String
          Dim intEL As Integer

24210     intEL = Erl
24220     strES = Err.Description
24230     LogError "frmEditAll", "mnuTestCodeMappingBio_Click", intEL, strES

End Sub

Private Sub mnuTestCodeMappingCoag_Click()

24240     On Error GoTo mnuTestCodeMappingCoag_Click_Error

24250     frmTestCodeMapping.Discipline = "COAG"
24260     frmTestCodeMapping.Show 1

24270     Exit Sub

mnuTestCodeMappingCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

24280     intEL = Erl
24290     strES = Err.Description
24300     LogError "frmEditAll", "mnuTestCodeMappingCoag_Click", intEL, strES

End Sub

Private Sub mnuTestCodeMappingHaem_Click()

24310     On Error GoTo mnuTestCodeMappingHaem_Click_Error

24320     frmTestCodeMapping.Discipline = "HAEM"
24330     frmTestCodeMapping.Show 1

24340     Exit Sub

mnuTestCodeMappingHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

24350     intEL = Erl
24360     strES = Err.Description
24370     LogError "frmEditAll", "mnuTestCodeMappingHaem_Click", intEL, strES

End Sub

Private Sub mnuTestSequence_Click()

24380     With frmTestSequence
24390         .Discipline = "Bio"
24400         .Show 1
24410     End With

End Sub

Private Sub mnuUnitsPrecision_Click()

24420     With frmTestUnitsPrecision
24430         .Discipline = "Bio"
24440         .SampleType = "S"
24450         .Show 1
24460     End With

End Sub

Private Sub tMCH_KeyPress(KeyAscii As Integer)

24470     cmdSaveHaem.Enabled = True

End Sub


Private Sub tMCHC_KeyPress(KeyAscii As Integer)

24480     cmdSaveHaem.Enabled = True

End Sub


Private Sub tmrUpDown_Timer()

24490     txtSampleID = Val(txtSampleID) + UpDownDirection
24500     If Val(txtSampleID) < 1 Then
24510         txtSampleID = "1"
24520     ElseIf Val(txtSampleID) > 9999999 Then
24530         txtSampleID = "9999999"
24540     End If

End Sub

Private Sub tRecTime_GotFocus()

24550     tRecTime.SelStart = 0
24560     tRecTime.SelLength = 0

End Sub

Private Sub tRecTime_KeyPress(KeyAscii As Integer)

24570     pBar = 0

24580     cmdSaveHoldDemographics.Enabled = True
24590     cmdSaveDemographics.Enabled = True

End Sub


Private Sub chkOld_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

24600     cmdSaveBio.Enabled = True

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
24610     On Error GoTo ErrorHandler

24620     pBar = 0
24630     cmdViewReports.Visible = False

24640     Debug.Print "sstab1.click " & Now

24650     ShowMenuLists

24660     Select Case PreviousTab
              Case 0
                  '          If CheckDemographics = False Then
                  '              MsgBox "Please check demographics.", vbInformation
                  '              SSTab1.Tab = 0
                  '              Exit Sub
                  '          Else
24670             If cmdSaveHoldDemographics.Enabled Then
24680                 cmdSaveHoldDemographics_Click
24690             End If
                  '          End If
24700         Case 1
24710             If cmdSaveHaem.Enabled Then
24720                 SaveHaematology 0
24730                 SaveComments
24740                 UpdateMRU Me
24750                 cmdSaveHaem.Enabled = False
24760             End If
24770         Case 2
24780             imgLast.ToolTipText = "Last Biochemistry Record"
24790             If cmdSaveBio.Enabled Then
24800                 cmdSavebio_Click
24810             End If
24820         Case 3
24830             If cmdSaveCoag.Enabled Then
24840                 cmdSaveCoag_Click
24850             End If

24860     End Select

24870     Select Case SSTab1.Tab
              Case 0:    'Demographics
24880             lblResultOrRequest = "SampleID"
24890             imgLast.ToolTipText = "Next Record"
24900             bPrintHold.Enabled = False
24910             bPrint.Enabled = False
24920         Case 1:    'Haematology
24930             lblResultOrRequest = "Results"
24940             imgLast.ToolTipText = "Last Haematology Record"
24950             If Not HaemLoaded Then
24960                 Debug.Print "sstab1_click"
24970                 LoadHaematology
24980                 HaemLoaded = True
24990             End If
25000             bPrintHold.Enabled = True
25010             bPrint.Enabled = True
25020             SetViewReports "Haematology", txtSampleID

25030         Case 2:    'Biochemistry
25040             lblResultOrRequest = "Results"
25050             imgLast.ToolTipText = "Last Biochemistry Record"
25060             If Not BioLoaded Then
25070                 LoadBiochemistry
25080                 BioLoaded = True
25090             End If
25100             bPrintHold.Enabled = True
25110             bPrint.Enabled = True
25120             SetViewReports "Biochemistry", txtSampleID

25130         Case 3:    'Coagulation
25140             lblResultOrRequest = "Results"
25150             imgLast.ToolTipText = "Last Coagulation Record"
25160             If Not CoagLoaded Then
25170                 LoadCoagulation
25180                 CoagLoaded = True
25190             End If
25200             bPrintHold.Enabled = True
25210             bPrint.Enabled = True
25220             SetViewReports "Coagulation", txtSampleID

25230         Case 6:    'External
25240             lblResultOrRequest = "Results"
25250             imgLast.ToolTipText = "Last External Record"
25260             LoadExt
25270             bPrintHold.Enabled = True
25280             bPrint.Enabled = True

25290     End Select

          'SetViewHistory
25300     ShowHistory (SSTab1.Tab)
25310     LoadComments
25320     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

25330     pBar = 0

End Sub


Private Sub tBasP_KeyUp(KeyCode As Integer, Shift As Integer)

25340     If Val(tWBC) > 0 Then
25350         tBasA = Format(Val(tBasP) * Val(tWBC) / 100, "0.0")
25360     End If

End Sub

Private Sub tEosP_KeyUp(KeyCode As Integer, Shift As Integer)

25370     If Val(tWBC) > 0 Then
25380         tEosA = Format(Val(tEosP) * Val(tWBC) / 100, "0.0")
25390     End If

End Sub

Private Sub tLymP_KeyUp(KeyCode As Integer, Shift As Integer)

25400     If Val(tWBC) > 0 Then
25410         tLymA = Format(Val(tLymP) * Val(tWBC) / 100, "0.0")
25420     End If

End Sub

Private Sub tMonoP_KeyUp(KeyCode As Integer, Shift As Integer)

25430     If Val(tWBC) > 0 Then
25440         tMonoA = Format(Val(tMonoP) * Val(tWBC) / 100, "0.0")
25450     End If

End Sub


Private Sub tNeutP_KeyUp(KeyCode As Integer, Shift As Integer)

25460     If Val(tWBC) > 0 Then
25470         tNeutA = Format(Val(tNeutP) * Val(tWBC) / 100, "0.0")
25480     End If

End Sub

Private Sub tRecTime_LostFocus()
25490     SetDatesColour
End Sub

Private Sub tSampleTime_GotFocus()

25500     tSampleTime.SelStart = 0
25510     tSampleTime.SelLength = 0

End Sub

Private Sub tSampleTime_LostFocus()
25520     SetDatesColour
End Sub

Private Sub txtaddress_Change(Index As Integer)

25530     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP & ":" & cmbClinician

End Sub

Private Sub txtaddress_KeyPress(Index As Integer, KeyAscii As Integer)

25540     cmdSaveHoldDemographics.Enabled = True
25550     cmdSaveDemographics.Enabled = True

End Sub


Private Sub txtaddress_LostFocus(Index As Integer)

25560     txtAddress(Index) = Initial2Upper(txtAddress(Index))

End Sub


Private Sub txtage_Change()

25570     lblAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
25580     If lblDemogValid.Caption = "Demographics Valid" Then Exit Sub
25590     cmdSaveHoldDemographics.Enabled = True
25600     cmdSaveDemographics.Enabled = True

End Sub


Private Sub tBasA_KeyPress(KeyAscii As Integer)

25610     cmdSaveHaem.Enabled = True

End Sub


Private Sub tBasP_KeyPress(KeyAscii As Integer)

25620     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtBioComment_KeyUp(KeyCode As Integer, Shift As Integer)

25630     cmdSaveBio.Enabled = True

End Sub

Private Sub txtCD3A_KeyPress(KeyAscii As Integer)

25640     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtCD3P_KeyPress(KeyAscii As Integer)

25650     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtCD48_KeyPress(KeyAscii As Integer)

25660     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtCD4A_KeyPress(KeyAscii As Integer)

25670     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtCD4P_KeyPress(KeyAscii As Integer)

25680     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtCD8A_KeyPress(KeyAscii As Integer)

25690     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtCD8P_KeyPress(KeyAscii As Integer)

25700     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtchart_Change()

25710     lblChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)
25720     If lblDemogValid.Caption = "Demographics Valid" Then Exit Sub
25730     cmdSaveHoldDemographics.Enabled = True
25740     cmdSaveDemographics.Enabled = True

25750     If SSTab1.Tab <> 0 Then
25760         SSTab1.Tab = 0
25770     End If
End Sub


Private Sub txtchart_LostFocus()

25780     If Trim$(txtChart) = "" Then Exit Sub

25790     If lblDemogValid.Caption = "Demographics Valid" Then Exit Sub

25800     LoadPatientFromChart Me, mNewRecord

25810     cmdSaveHoldDemographics.Enabled = True
25820     cmdSaveDemographics.Enabled = True
25830     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
25840         If MatchingDemoLoaded = False Then LoadMatchingDemo
25850     End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadMatchingDemo
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub LoadMatchingDemo()
25860     On Error GoTo LoadMatchingDemo_Error
          Dim SearchConditon As String


25870     SearchConditon = " AND D.Dob = '" & Format(txtDoB, "YYYY/MMM/DD") & "'"
25880     If Val(txtLabNo & "") <> 0 Then
25890         If FndMatchingRecords(SearchConditon) > 0 Then
25900             With frmPatHistoryChart
25910                 .LabNoUpd = txtLabNo
25920                 Set .EditScreen = Me
25930                 .PatientHistory = SearchConditon
25940                 If frmPatHistoryChart.Visible = False Then
25950                     If frmPatHistoryChart.g.TextMatrix(1, 1) <> "" Then
25960                         .Show 1
25970                     End If
25980                 End If
25990                 MatchingDemoLoaded = True
26000             End With
26010             Exit Sub
26020         End If
26030     End If

26040     ClearLabNoSelection
26050     Exit Sub


LoadMatchingDemo_Error:

          Dim strES As String
          Dim intEL As Integer

26060     intEL = Erl
26070     strES = Err.Description
26080     LogError "frmEditAll", "LoadMatchingDemo", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : FndMatchingRecords
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FndMatchingRecords(Condition As String)

26090     On Error GoTo FndMatchingRecords_Error
          Dim sql As String
          Dim tb As ADODB.Recordset

26100     sql = "select Count(PatName) as Cnt from Demographics D " & _
              " WHERE D.SAMPLEID <> '-9' " & Condition & " "
          '"  GROUP BY D.PatName,D.Chart,D.Addr0,D.DoB,D.Sex "
26110     Set tb = New Recordset
26120     RecOpenClient 0, tb, sql

26130     If tb.EOF = False Then
26140         FndMatchingRecords = tb!Cnt
26150     End If


26160     Exit Function


FndMatchingRecords_Error:

          Dim strES As String
          Dim intEL As Integer

26170     intEL = Erl
26180     strES = Err.Description
26190     LogError "frmEditAll", "FndMatchingRecords", intEL, strES, sql
End Function

Private Sub txtCoagComment_KeyUp(KeyCode As Integer, Shift As Integer)

26200     cmdSaveCoag.Enabled = True

End Sub

Private Sub txtDemographicComment_KeyUp(KeyCode As Integer, Shift As Integer)

26210     cmdSaveHoldDemographics.Enabled = True
26220     cmdSaveDemographics.Enabled = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtDoB_Change
' Author    : XPMUser
' Date      : 23/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtDoB_Change()
26230     On Error GoTo txtDoB_Change_Error


26240     LabNoUpdatePrviousData = ""
26250     LabNoUpdatePrvColor
26260     lblDoB = txtDoB


26270     Exit Sub


txtDoB_Change_Error:

          Dim strES As String
          Dim intEL As Integer

26280     intEL = Erl
26290     strES = Err.Description
26300     LogError "frmEditAll", "txtDoB_Change", intEL, strES

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)

26310     If lblDemogValid.Caption = "Demographics Valid" Then Exit Sub

26320     cmdSaveHoldDemographics.Enabled = True
26330     cmdSaveDemographics.Enabled = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtDoB_LostFocus
' Author    : XPMUser
' Date      : 03/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtDoB_LostFocus()
26340     On Error GoTo txtDoB_LostFocus_Error


26350     If lblDemogValid.Caption = "Demographics Valid" Then Exit Sub
26360     txtDoB = Convert62Date(txtDoB, BACKWARD)
26370     txtAge = CalcAge(txtDoB, dtSampleDate)
26380     If Val(txtAge) < 0 Then
26390         txtDoB = ""
26400         txtAge = ""
26410     End If
26420     LoadBiochemistry
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
26430     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
26440         If MatchingDemoLoaded = False Then LoadMatchingDemo
26450     End If

26460     Exit Sub


txtDoB_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

26470     intEL = Erl
26480     strES = Err.Description
26490     LogError "frmEditAll", "txtDoB_LostFocus", intEL, strES
End Sub


Private Sub tEosA_KeyPress(KeyAscii As Integer)

26500     cmdSaveHaem.Enabled = True

End Sub


Private Sub tEosP_KeyPress(KeyAscii As Integer)

26510     cmdSaveHaem.Enabled = True

End Sub


Private Sub tESR_Change()

26520     If Trim$(tESR) <> "" Then
26530         cESR = 1
26540     Else
26550         cESR = 0
26560     End If

End Sub

Private Sub tESR_KeyPress(KeyAscii As Integer)

26570     cmdSaveHaem.Enabled = True

End Sub


Private Sub tHct_KeyPress(KeyAscii As Integer)

26580     cmdSaveHaem.Enabled = True

End Sub


Private Sub tHgb_KeyPress(KeyAscii As Integer)

26590     cmdSaveHaem.Enabled = True

End Sub


Private Sub TimerBar_Timer()

26600     pBar = pBar + 1

26610     If pBar = pBar.max Then
26620         Unload Me
26630         Exit Sub
26640     End If

End Sub


Private Sub tLymA_KeyPress(KeyAscii As Integer)

26650     cmdSaveHaem.Enabled = True

End Sub


Private Sub tLymP_KeyPress(KeyAscii As Integer)

26660     cmdSaveHaem.Enabled = True

End Sub


Private Sub tMCV_KeyPress(KeyAscii As Integer)

26670     cmdSaveHaem.Enabled = True

End Sub


Private Sub tMonoA_KeyPress(KeyAscii As Integer)

26680     cmdSaveHaem.Enabled = True

End Sub


Private Sub tMonoP_KeyPress(KeyAscii As Integer)

26690     cmdSaveHaem.Enabled = True

End Sub


Private Sub tMonospot_Change()

26700     If Trim$(tMonospot) <> "" Then
26710         cMonospot = 1
26720     Else
26730         cMonospot = 0
26740     End If

End Sub

Private Sub tMonospot_Click()

          Dim f As Form

26750     cmdSaveHaem.Enabled = True

26760     If Trim$(tMonospot) = "" Or tMonospot = "?" Then
26770         tMonospot = "Negative"
26780     ElseIf tMonospot = "Negative" Then
26790         tMonospot = "Positive"
26800     Else
26810         tMonospot = ""
26820     End If

26830     If tMonospot <> "" Then
26840         If Not CheckReagentLotNumber("Monospot", txtSampleID) Then
26850             Set f = New frmCheckReagentLotNumber
26860             With f
26870                 .Analyte = "Monospot"
26880                 .SampleID = txtSampleID
26890                 .Show 1
26900             End With
26910             Unload f
26920             Set f = Nothing
26930         End If
26940     End If

End Sub

Private Sub tMonospot_KeyPress(KeyAscii As Integer)

          Dim f As Form
          Dim sql As String

26950     On Error GoTo tMonospot_KeyPress_Error

26960     cmdSaveHaem.Enabled = True

26970     If Trim$(tMonospot) = "" Then
26980         tMonospot = "Negative"
26990     ElseIf tMonospot = "Negative" Then
27000         tMonospot = "Positive"
27010     Else
27020         tMonospot = ""
27030     End If

27040     If tMonospot <> "" Then
27050         If Not CheckReagentLotNumber("Monospot", txtSampleID) Then
27060             Set f = New frmCheckReagentLotNumber
27070             With f
27080                 .Analyte = "Monospot"
27090                 .SampleID = txtSampleID
27100                 .Show 1
27110             End With
27120             Unload f
27130             Set f = Nothing
27140         End If
27150     Else
27160         sql = "Delete from ReagentLotNumbers where " & _
                  "Analyte = 'Monospot' " & _
                  "and SampleID = " & Val(txtSampleID)
27170         Cnxn(0).Execute sql
27180     End If

27190     Exit Sub

tMonospot_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

27200     intEL = Erl
27210     strES = Err.Description
27220     LogError "frmEditAll", "tMonospot_KeyPress", intEL, strES, sql


End Sub


Private Sub tMPV_KeyPress(KeyAscii As Integer)

27230     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtExtSampleID_Change()
27240     cmdSaveHoldDemographics.Enabled = True
27250     cmdSaveDemographics.Enabled = True
End Sub



'---------------------------------------------------------------------------------------
' Procedure : txtExtSampleID_LostFocus
' Author    : Masood
' Date      : 08/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtExtSampleID_LostFocus()
27260     On Error GoTo txtExtSampleID_LostFocus_Error
27270     If txtExtSampleID <> "" Then
27280         With frmGporders
27290             SavedDemoFromGPCom = False
27300             CancelFromGpCom = False
27310             Set .EditScreen = Me
27320             .SampleID = txtSampleID
27330             .SampleIDExt = txtExtSampleID
27340             .ClinicalDetails = cClDetails
27350             .DisiplinesQuery = " AND P.Department IN ('Biochemistry','Coagulation','External','Haematology')"
27360             .Show 1

27370             If CancelFromGpCom = False Then
27380                 If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
                          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
27390                     If MatchingDemoLoaded = False Then LoadMatchingDemo
27400                     cmdSaveDemographics.Value = True
27410                 End If
27420             End If
27430         End With
27440     End If

27450     Exit Sub


txtExtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

27460     intEL = Erl
27470     strES = Err.Description
27480     LogError "frmEditAll", "txtExtSampleID_LostFocus", intEL, strES
End Sub

Private Sub txtFilmComment_KeyPress(KeyAscii As Integer)

27490     cmdSaveHaem.Enabled = True

End Sub


Private Sub txtIRF_KeyPress(KeyAscii As Integer)

27500     cmdSaveHaem.Enabled = True

End Sub

Private Sub txtsampleid_KeyPress(KeyAscii As Integer)

27510     If KeyAscii = 13 Then txtsampleid_LostFocus

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSex_LostFocus
' Author    : Masood
' Date      : 04/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSex_LostFocus()
27520     On Error GoTo txtSex_LostFocus_Error


          '20    Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)


27530     Exit Sub


txtSex_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

27540     intEL = Erl
27550     strES = Err.Description
27560     LogError "frmEditAll", "txtSex_LostFocus", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSurName_Change
' Author    : XPMUser
' Date      : 23/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSurName_Change()
27570     On Error GoTo txtSurName_Change_Error

27580     LabNoUpdatePrviousData = ""
27590     LabNoUpdatePrvColor
27600     lblName = Trim$(txtSurName & " " & txtForeName)


27610     Exit Sub


txtSurName_Change_Error:

          Dim strES As String
          Dim intEL As Integer

27620     intEL = Erl
27630     strES = Err.Description
27640     LogError "frmEditAll", "txtSurName_Change", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtForeName_Change
' Author    : XPMUser
' Date      : 23/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtForeName_Change()
27650     On Error GoTo txtForeName_Change_Error

27660     LabNoUpdatePrviousData = ""
27670     LabNoUpdatePrvColor
27680     lblName = Trim$(txtSurName & " " & txtForeName)


27690     Exit Sub


txtForeName_Change_Error:

          Dim strES As String
          Dim intEL As Integer

27700     intEL = Erl
27710     strES = Err.Description
27720     LogError "frmEditAll", "txtForeName_Change", intEL, strES

End Sub

Private Sub txtSurName_KeyPress(KeyAscii As Integer)

27730     If lblDemogValid.Caption <> "Demographics Valid" Then
27740         cmdSaveHoldDemographics.Enabled = True
27750         cmdSaveDemographics.Enabled = True
27760     End If

27770     If SSTab1.Tab <> 0 Then
27780         SSTab1.Tab = 0
27790     End If

End Sub


Private Sub txtForeName_KeyPress(KeyAscii As Integer)

27800     If lblDemogValid.Caption <> "Demographics Valid" Then
27810         cmdSaveHoldDemographics.Enabled = True
27820         cmdSaveDemographics.Enabled = True
27830     End If

27840     If SSTab1.Tab <> 0 Then
27850         SSTab1.Tab = 0
27860     End If

End Sub



'---------------------------------------------------------------------------------------
' Procedure : txtSurname_LostFocus
' Author    : XPMUser
' Date      : 03/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSurname_LostFocus()

          Dim strSex As String
          Dim strSurName As String
          Dim strForeName As String

27870     On Error GoTo txtSurname_LostFocus_Error


27880     strSex = txtSex
27890     strSurName = Trim$(txtSurName)
27900     strForeName = Trim$(txtForeName)

27910     NameLostFocus strSurName, strForeName, strSex

27920     txtSurName = strSurName
27930     txtForeName = strForeName
27940     txtSex = strSex

          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
27950     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
27960         If MatchingDemoLoaded = False Then LoadMatchingDemo
27970     End If

27980     Exit Sub


txtSurname_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

27990     intEL = Erl
28000     strES = Err.Description
28010     LogError "frmEditAll", "txtSurname_LostFocus", intEL, strES
End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtForeName_LostFocus
' Author    : XPMUser
' Date      : 03/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtForeName_LostFocus()

          Dim strSex As String
          Dim strSurName As String
          Dim strForeName As String

28020     On Error GoTo txtForeName_LostFocus_Error


28030     strSex = txtSex
28040     strSurName = Trim$(txtSurName)
28050     strForeName = Trim$(txtForeName)

28060     NameLostFocus strSurName, strForeName, strSex

28070     txtSurName = strSurName
28080     txtForeName = strForeName
28090     txtSex = strSex
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)
28100     If txtSurName <> "" And txtForeName <> "" And txtDoB <> "" Then
28110         If MatchingDemoLoaded = False Then LoadMatchingDemo
28120     End If


28130     Exit Sub


txtForeName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

28140     intEL = Erl
28150     strES = Err.Description
28160     LogError "frmEditAll", "txtForeName_LostFocus", intEL, strES
End Sub

Private Sub tNeutA_KeyPress(KeyAscii As Integer)

28170     cmdSaveHaem.Enabled = True

End Sub


Private Sub tNeutP_KeyPress(KeyAscii As Integer)

28180     cmdSaveHaem.Enabled = True

End Sub


Private Sub tnrbcA_KeyPress(KeyAscii As Integer)

28190     cmdSaveHaem.Enabled = True

End Sub

Private Sub tnrbcP_KeyPress(KeyAscii As Integer)

28200     cmdSaveHaem.Enabled = True

End Sub

Private Sub tPlt_KeyPress(KeyAscii As Integer)

28210     cmdSaveHaem.Enabled = True

End Sub


Private Sub tRBC_KeyPress(KeyAscii As Integer)

28220     cmdSaveHaem.Enabled = True

End Sub


Private Sub tRDWCV_KeyPress(KeyAscii As Integer)

28230     cmdSaveHaem.Enabled = True

End Sub

Private Sub tRDWSD_KeyPress(KeyAscii As Integer)

28240     cmdSaveHaem.Enabled = True

End Sub

Private Sub tRetA_Change()

28250     If Trim$(tRetA) <> "" Then
28260         cRetics = 1
28270     Else
28280         cRetics = 0
28290     End If

End Sub

Private Sub tRetA_KeyPress(KeyAscii As Integer)

28300     cmdSaveHaem.Enabled = True

End Sub


Private Sub tRetP_Change()

28310     If Trim$(tRetP) <> "" Then
28320         cRetics = 1
28330     Else
28340         cRetics = 0
28350     End If

End Sub


Private Sub tRetP_KeyPress(KeyAscii As Integer)

28360     cmdSaveHaem.Enabled = True

End Sub



Public Sub txtsampleid_LostFocus()

28370     On Error GoTo txtsampleid_LostFocus_Error

28380     If Trim$(txtSampleID) = "" Then Exit Sub

28390     If Val(txtSampleID) > 100000000 Then
28400         txtSampleID = "1"
28410     End If

28420     txtSampleID = Format(Val(txtSampleID))

          'txtLabNo = Val(FndMaxID("demographics", "LabNo", "")) + 1
          'LabNoClear
28430     LoadAllDetails

          'Abubaker +++ 09/10/2023 (On Meting the situation where name not load we execute this function to reload the lost focus procedure and then store data to BugLog table)
          '          m_Counter added __ Abubaker 16-11-2023
          '          If m_Counter > 1 Then
          '            m_Counter = 0
          '          Else
          '            If lblName.Caption = "" Or cmbWard.Text = "" Then
          '                Call ReloadAndLogData
          '                m_Counter = m_Counter + 1
          '            Else
          '                m_Counter = 0
          '            End If
          '          End If
          'Abubaker --- 09/10/2023

28440     cmdSaveHoldDemographics.Enabled = False
28450     cmdSaveDemographics.Enabled = False
28460     cmdSaveHaem.Enabled = False
28470     cmdSaveBio.Enabled = False
28480     cmdSaveCoag.Enabled = False

28490     Exit Sub

txtsampleid_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

28500     intEL = Erl
28510     strES = Err.Description
28520     LogError "frmEditAll", "txtSampleID_LostFocus", intEL, strES

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
'    If txtSampleID.Text <> "" Then
'        Sql = "Select SampleID from Demographics Where SampleID = '" & txtSampleID.Text & "'"
'        Set tb = New Recordset
'        RecOpenServer 0, tb, Sql
'        If Not tb Is Nothing Then
'            If Not tb.EOF Then
'                currentTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
'                sampleIDValue = txtSampleID.Text
'                lblNameValue = lblName.Caption
'                Sql = "INSERT INTO BugLog (DateTime, SampleID, LblName) VALUES ('" & currentTime & "', '" & sampleIDValue & "', '" & lblNameValue & "');"
'                Cnxn(0).Execute Sql
'                Call txtsampleid_LostFocus
'            End If
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


'---------------------------------------------------------------------------------------
' Procedure : ClearLabNoSelection
' Author    : Masood
' Date      : 18/Feb/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ClearLabNoSelection()

28530     LabNoUpdatePrviousData = ""
28540     frmEditAll.txtMultiSeltdDemoForLabNoUpd = ""
28550     frmEditAll.gMDemoLabNoUpd.Clear
28560     gMDemoLabNoUpd.Rows = 1
End Sub

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

28570     cmdSaveHoldDemographics.Enabled = True
28580     cmdSaveDemographics.Enabled = True

End Sub


Private Sub txtSex_Change()

28590     lblSex = txtSex

End Sub

Private Sub txtsex_Click()

28600     If StatusBar1.Panels(3).Text = "Demographics Validated" Then
28610         Exit Sub
28620     End If

28630     Select Case Trim$(txtSex)
              Case "": txtSex = "Male"
28640         Case "Male": txtSex = "Female"
28650         Case "Female": txtSex = ""
28660         Case Else: txtSex = ""
28670     End Select

28680     LoadBiochemistry

28690     cmdSaveHoldDemographics.Enabled = True
28700     cmdSaveDemographics.Enabled = True

End Sub


Private Sub txtsex_KeyPress(KeyAscii As Integer)

28710     KeyAscii = 0
28720     txtsex_Click

End Sub


Private Sub tWBC_KeyPress(KeyAscii As Integer)

28730     cmdSaveHaem.Enabled = True

End Sub

Private Sub txtCoagComment_LostFocus()

          Dim S As Variant
          Dim n As Integer
          Dim Temp As String

28740     If Trim$(txtCoagComment) = "" Then Exit Sub

28750     S = Split(txtCoagComment, " ")

28760     For n = 0 To UBound(S)
28770         Temp = ListTextFor("CO", S(n))
28780         If Temp <> "" Then
28790             S(n) = Temp
28800             Temp = ""
28810         End If
28820     Next

28830     txtCoagComment = Join(S, " ")

End Sub

Private Sub txtDemographicComment_LostFocus()
    '
    'Dim Lx As List
    'Dim s As Variant
    'Dim n As Integer
    '
    'If Trim$(txtDemographicComment) = "" Then Exit Sub
    '
    's = Split(txtDemographicComment, " ")
    '
    'For n = 0 To UBound(s)
    '  Set Lx = colLists("DE", s(n))
    '  If Not Lx Is Nothing Then
    '    s(n) = Lx.Text
    '  End If
    'Next
    '
    'txtDemographicComment = Join(s, " ")
    '
End Sub


Private Sub txtHaemComment_KeyPress(KeyAscii As Integer)

28840     cmdSaveHaem.Enabled = True

End Sub




Public Property Let StartInDepartment(ByVal strNewValue As String)

28850     m_StartInDepartment = strNewValue

End Property
Public Property Let PrintToPrinter(ByVal strNewValue As String)

28860     pPrintToPrinter = strNewValue

End Property

Public Property Get PrintToPrinter() As String

28870     PrintToPrinter = pPrintToPrinter

End Property

Private Sub SetDatesColour()

28880     On Error GoTo SetDatesColour_Error

28890     If CheckDateSequence(dtSampleDate, dtRecDate, dtRunDate, tSampleTime, tRecTime) Then
28900         fr(1).ForeColor = vbButtonText
28910         fr(1).Font.Bold = False
              '50        Label1(54).ForeColor = vbButtonText
              '60        Label1(54).Font.Bold = False
28920         Label1(55).ForeColor = vbButtonText
28930         Label1(55).Font.Bold = False
28940         Label1(45).ForeColor = vbButtonText
28950         Label1(45).Font.Bold = False
28960         lblDateError.Visible = False
28970     Else
28980         fr(1).ForeColor = vbRed
28990         fr(1).Font.Bold = True
              '150       Label1(54).ForeColor = vbRed
              '160       Label1(54).Font.Bold = True
29000         Label1(55).ForeColor = vbRed
29010         Label1(55).Font.Bold = True
29020         Label1(45).ForeColor = vbRed
29030         Label1(45).Font.Bold = True
29040         lblDateError.Visible = True
29050     End If

29060     Exit Sub

SetDatesColour_Error:

          Dim strES As String
          Dim intEL As Integer

29070     intEL = Erl
29080     strES = Err.Description
29090     LogError "basShared", "SetDatesColour", intEL, strES

End Sub

Private Sub FillCommentTemplates(cmb As ComboBox, Department As String)

          Dim tb As Recordset
          Dim sql As String


29100     On Error GoTo FillCommentTemplates_Error

29110     sql = "Select * From CommentsTemplate " & _
              "Where Department = '" & Department & "' And Inactive = 0 Order By CommentName"
29120     Set tb = New Recordset
29130     RecOpenClient 0, tb, sql
29140     If Not tb.EOF Then

29150         With cmb
29160             .Clear
29170             While Not tb.EOF
29180                 .AddItem tb!CommentName & ""
29190                 .ItemData(.NewIndex) = tb!CommentID
29200                 tb.MoveNext

29210             Wend
29220             .Text = "*** Insert Comment Template ***"
29230         End With

29240     End If

29250     Exit Sub

FillCommentTemplates_Error:

          Dim strES As String
          Dim intEL As Integer

29260     intEL = Erl
29270     strES = Err.Description
29280     LogError "frmEditAll", "FillCommentTemplates", intEL, strES, sql


End Sub

Private Sub cmdScan_Click()


29290     On Error GoTo cmdScan_Click_Error

29300     With frmScan
29310         .txtSampleID = txtSampleID
29320         .Show 1
29330     End With
29340     SetViewScans txtSampleID, cmdViewScan

29350     Exit Sub

cmdScan_Click_Error:

          Dim strES As String
          Dim intEL As Integer

29360     intEL = Erl
29370     strES = Err.Description
29380     LogError "frmEditAll", "cmdScan_Click", intEL, strES


End Sub

'---------------------------------------------------------------------------------------
' Procedure : LabNoUpdatePrvData
' Author    : XPMUser
' Date      : 25/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LabNoUpdatePrvData(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String)
29390     On Error GoTo LabNoUpdatePrvData_Error
          Dim sql As String

29400     If txtMultiSeltdDemoForLabNoUpd <> "" Then
29410         sql = txtMultiSeltdDemoForLabNoUpd
29420         Cnxn(0).Execute (sql)

29430     End If
          
29440     ClearLabNoSelection
29450     Exit Sub


LabNoUpdatePrvData_Error:

          Dim strES As String
          Dim intEL As Integer

29460     intEL = Erl
29470     strES = Err.Description
29480     LogError "frmEditAll", "LabNoUpdatePrvData", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : FindLatestAddress
' Author    : XPMUser
' Date      : 25/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FindLatestAddress(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String) As String
          Dim sql As String

29490     On Error GoTo FindLatestAddress_Error

          Dim tb As New ADODB.Recordset
29500     sql = "Select Addr0 from demographics  "
29510     sql = sql & " WHERE "
29520     sql = sql & " UPPER(PatName) ='" & UCase(Name) & "'"
29530     sql = sql & " AND DoB ='" & DoB & "'"
29540     sql = sql & " AND UPPER(Sex) ='" & UCase(Sex) & "'"
29550     sql = sql & " AND UPPER(Chart) ='" & UCase(ChartNo) & "'"
29560     sql = sql & " ORDER BY DateTimeDemographics DESC  "
29570     Set tb = New Recordset
29580     RecOpenServer 0, tb, sql

29590     If Not tb.EOF Then
29600         FindLatestAddress = tb!Addr0
29610     End If


29620     Exit Function


FindLatestAddress_Error:

          Dim strES As String
          Dim intEL As Integer

29630     intEL = Erl
29640     strES = Err.Description
29650     LogError "frmEditAll", "FindLatestAddress", intEL, strES, sql
End Function


'---------------------------------------------------------------------------------------
' Procedure : ShowHistory
' Author    : XPMUser
' Date      : 11/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ShowHistory(TabNo As Integer)
29660     On Error GoTo ShowHistory_Error
          Dim sql As String
          Dim DispName As String
          Dim sn As New ADODB.Recordset
          Dim tb As New ADODB.Recordset



          '10    Select Case SSTab1.Tab
          '      Case 0: cmdHistory.Visible = False
          '20    Case 1: cmdHistory.Visible = PreviousHaem
          '30    Case 2: cmdHistory.Visible = PreviousBio
          '40    Case 3: cmdHistory.Visible = PreviousCoag
          '50    End Select


29670     cmdHistory.Visible = False

29680     If SSTab1.Tab = 0 Then
29690         Exit Sub
29700     ElseIf SSTab1.Tab = 1 Then
29710         DispName = "Haem"
29720     ElseIf SSTab1.Tab = 2 Then
29730         DispName = "Bio"
29740     ElseIf SSTab1.Tab = 3 Then
29750         DispName = "Coag"
29760     Else
29770         Exit Sub
29780     End If

29790     sql = " SELECT     TOP (1) R.Rundate,D.SampleDate,R.sampleid " & _
              " FROM         demographics AS D INNER JOIN " & DispName & "Results AS R ON D.SampleID = R.sampleid  " & _
              " WHERE     (D.DoB = '" & Format(txtDoB, "dd/mmm/yyyy") & "') AND (D.PatName = '" & AddTicks(txtSurName & " " & txtForeName) & "') " & _
              "AND Hospital = '" & cmbHospital & "' " & _
              "  ORDER BY D.SampleDate,D.sampleid ASC "


29800     Set sn = New Recordset
29810     RecOpenServer 0, sn, sql
29820     If Not sn.EOF Then
29830         If Not IsNull(sn!SampleID) Then
29840             cmdHistory.Visible = True
29850         End If
29860     End If
29870     Exit Sub


ShowHistory_Error:

          Dim strES As String
          Dim intEL As Integer

29880     intEL = Erl
29890     strES = Err.Description
29900     LogError "frmEditAll", "ShowHistory", intEL, strES, sql
End Sub


Private Function DemographicsUniLabNoSelect(PatName As String, DoB As String, Sex As String, Chart As String, LabNo As String) As Double

29910     On Error GoTo DemographicsUniLabNoSelect_Error
          Dim sql As String
          Dim tb As New ADODB.Recordset
29920     If PatName = "" Or DoB = "" Or Sex = "" Then
29930         Exit Function
29940     End If


29950     sql = "select Top 1 ISNULL(LabNo,0) as LabNo  from DemographicsUniLabNo As D  " & _
              " WHERE ISNULL(LabNo,0)  <> 0 AND  D.PatName='" & AddTicks(PatName) & "' AND DoB = '" & Format(DoB, "dd/MMM/yyyy") & "'" & _
              " ORDER BY DateTimeOfRecord DESC "

29960     Set tb = New Recordset
29970     RecOpenClient 0, tb, sql

29980     If tb.EOF = False Then
29990         DemographicsUniLabNoSelect = tb!LabNo
30000     Else
30010         LabNo = Val(FndMaxID("demographics", "LabNo", ""))
30020         Call DemographicsUniLabNoInsertValues("", UserName, PatName, DoB, Sex, Chart, LabNo)
30030         DemographicsUniLabNoSelect = LabNo
30040     End If

30050     txtLabNo = DemographicsUniLabNoSelect

30060     Exit Function


DemographicsUniLabNoSelect_Error:

          Dim strES As String
          Dim intEL As Integer

30070     intEL = Erl
30080     strES = Err.Description
30090     LogError "frmEditMicrobiology", "DemographicsUniLabNoSelect", intEL, strES
End Function

Private Function ExtendedIPUFlagsAvailable() As Boolean

          Dim sql As String
          Dim tb As Recordset

30100     On Error GoTo ExtendedIPUFlagsAvailable_Error

30110     sql = "SELECT Count(*) AS Cnt FROM HaemFlags WHERE SampleID = " & txtSampleID
30120     Set tb = New Recordset
30130     RecOpenServer 0, tb, sql
30140     ExtendedIPUFlagsAvailable = (tb!Cnt > 0)


30150     Exit Function

ExtendedIPUFlagsAvailable_Error:

          Dim strES As String
          Dim intEL As Integer

30160     intEL = Erl
30170     strES = Err.Description
30180     LogError "frmEditAll", "ExtendedIPUFlagsAvailable", intEL, strES, sql

End Function


Private Function CheckACR(ByVal BRs As BIEResults) As Boolean
          'returns True if ACR added

          Dim BR As BIEResult
          Dim Code As String

          Dim ACRValue As String

          Dim Rundate As String
          Dim RunTime As String
          Dim bNew As BIEResult
          Dim sql As String
          Dim tb As Recordset

          Dim CodeForUrinaryCreatinine As String
          Dim CodeForAlbumin As String
          Dim CodeForACR As String
          Dim ResultsValue As String

          Dim IsBothCodeInSample As Boolean
          Dim DP As String


30190     On Error GoTo CheckACR_Error


30200     CheckACR = False
          '+++ Junaid
30210     Exit Function
          '--- Junaid
30220     If BRs Is Nothing Then Exit Function


30230     CodeForUrinaryCreatinine = GetOptionSetting("BioCodeForUCreat", "1096")
30240     CodeForAlbumin = GetOptionSetting("BioCodeForUAlb", "2839")
30250     CodeForACR = GetOptionSetting("BioCodeForACR", "ACR")

30260     If Not (CheckCodeExistsInResult(BRs, CodeForUrinaryCreatinine) = True And CheckCodeExistsInResult(BRs, CodeForAlbumin) = True) Then
30270         Exit Function
30280     End If

30290     DP = FindFeildValue("BioTestDefinitions", "DP", " WHERE Code = '" & CodeForACR & "'")

30300     For Each BR In BRs
30310         Code = UCase$(Trim$(BR.Code))
30320         If Code = CodeForUrinaryCreatinine Then
                  '+++Junaid
30330             ACRValue = 0 'Val(CalculateACR(CheckResultsValue(BRs, CodeForUrinaryCreatinine), CheckResultsValue(BRs, CodeForAlbumin)))
                  '---Junaid
30340             If IsDate(BR.Rundate) Then
30350                 Rundate = BR.Rundate
30360             Else
30370                 Rundate = Format$(BR.RunTime, "dd/mmm/yyyy")
30380             End If

30390             RunTime = BR.RunTime
30400             sql = "SELECT * FROM BioResults WHERE " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "AND Code = '" & CodeForACR & "'"
30410             Set tb = New Recordset
30420             RecOpenClient 0, tb, sql
30430             If tb.EOF Then
30440                 tb.AddNew
30450             End If
30460             tb!SampleID = txtSampleID

30470             tb!Rundate = Rundate
30480             tb!RunTime = RunTime

30490             tb!Code = CodeForACR
30500             tb!Result = ACRValue
30510             tb!Units = "mg/mmol"
30520             tb!Printed = BR.Printed
30530             tb!Valid = BR.Valid
30540             tb!FAXed = 0
30550             tb!Analyser = ""
30560             tb!SampleType = "U"
30570             tb.Update

30580             tb.Close

30590             sql = "Select Top 1 * From BioTestDefinitions " & _
                      "Where Code = '" & CodeForACR & "'"
30600             Set tb = New Recordset
30610             RecOpenClient 0, tb, sql
30620             If Not tb.EOF Then

30630                 Set bNew = New BIEResult
30640                 bNew.SampleID = txtSampleID
30650                 bNew.Code = CodeForACR
30660                 bNew.Rundate = Rundate
30670                 bNew.RunTime = RunTime
30680                 bNew.Result = ACRValue
30690                 bNew.Units = "mg/mmol"
30700                 bNew.Printed = BR.Printed
30710                 bNew.Valid = BR.Valid
30720                 bNew.SampleType = tb!SampleType
30730                 bNew.LongName = tb!LongName
30740                 bNew.ShortName = tb!ShortName
30750                 bNew.PlausibleLow = tb!PlausibleLow
30760                 bNew.PlausibleHigh = tb!PlausibleHigh
30770                 bNew.FlagLow = 0    'tb!FlagLow
30780                 bNew.Printformat = tb!DP
30790                 bNew.FlagHigh = 9999    'tb!FlagHigh

30800                 BRs.Add bNew

30810                 CheckACR = True


30820             End If
30830             Exit For
30840         End If
30850     Next


30860     Exit Function


CheckACR_Error:

          Dim strES As String
          Dim intEL As Integer

30870     intEL = Erl
30880     strES = Err.Description
30890     LogError "frmEditAll", "CheckACR", intEL, strES, sql

End Function




'---------------------------------------------------------------------------------------
' Procedure : CheckUrine24hr
' Author    : Masood
' Date      : 18/Feb/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CheckUrine24hr(ByVal BRs As BIEResults, Urine24HrCodeInOptions As String, Urine24HrCodeInOptionsDefult As String, _
          UrineCodeInOptions As String, UrineCodeInOptionsDefult As String) As Boolean
          'returns True if ACR added

          Dim BR As BIEResult

          Dim Urine24HrCode As String
          Dim Urine24HrResult As String
          Dim UrineCode As String
          Dim UrineResult As String
          Dim UrineVolumeCode As String
          Dim UrineVolumeResult As String

          'Dim Code As String
          'Dim Value As String
          '
          'Dim ValueForUrineVolume As String
          'Dim Code24HrForCalculate As String
          'Dim Code24HrValue As String


30900     On Error GoTo CheckUrine24hr_Error


30910     CheckUrine24hr = False

30920     If BRs Is Nothing Then Exit Function

30930     Urine24HrCode = GetOptionSetting(Urine24HrCodeInOptions, Urine24HrCodeInOptionsDefult)
30940     Urine24HrResult = CheckResultsValue(BRs, Urine24HrCode)
30950     If Urine24HrResult <> "" Then Exit Function

30960     UrineCode = GetOptionSetting(UrineCodeInOptions, Urine24HrCodeInOptionsDefult)
30970     UrineResult = CheckResultsValue(BRs, UrineCode)
30980     If UrineResult = "" Then Exit Function

30990     UrineVolumeCode = GetOptionSetting("BioCodeForUrineVolume", "TUV")
31000     UrineVolumeResult = CheckResultsValue(BRs, UrineVolumeCode)
31010     If UrineVolumeResult = "" Then Exit Function

31020     Urine24HrResult = Val(UrineVolumeResult) * Val(UrineResult)
31030     AddTestInDbnCollection BRs, Urine24HrCode, Urine24HrResult

31040     Exit Function

CheckUrine24hr_Error:

          Dim strES As String
          Dim intEL As Integer

31050     intEL = Erl
31060     strES = Err.Description
31070     LogError "frmEditAll", "CheckUrine24hr", intEL, strES

End Function



Private Function CheckCreatinineClearance(ByVal BRs As BIEResults) As Boolean
          'returns True if ACR added

          Dim BR As BIEResult
          Dim BRsPrev As New BIEResults

          Dim CreatinineClearanceCode As String
          Dim CreatinineClearanceResult As String

          Dim SerumCreatinineCode As String
          Dim SerumCreatinineResult As String

          Dim UrineVolumeCode As String
          Dim UrineVolumeResult As String

          Dim UrineCreatinineCode As String
          Dim UrineCreatinineResult As String

          Dim tb As Recordset
          Dim sql As String


31080     On Error GoTo CheckUrineCreatinineRatio_Error


31090     CheckCreatinineClearance = False

31100     If BRs Is Nothing Then Exit Function

31110     CreatinineClearanceCode = GetOptionSetting("BioCodeForCreatinineClearance", "CC")
31120     CreatinineClearanceResult = CheckResultsValue(BRs, CreatinineClearanceCode)
31130     If CreatinineClearanceResult <> "" Then Exit Function


31140     UrineCreatinineCode = GetOptionSetting("BIOCODEFORUCREAT", "1096")
31150     UrineCreatinineResult = CheckResultsValue(BRs, UrineCreatinineCode)
31160     If UrineCreatinineResult = "" Then Exit Function

31170     UrineVolumeCode = GetOptionSetting("BioCodeForUrineVolume", "TUV")
31180     UrineVolumeResult = CheckResultsValue(BRs, UrineVolumeCode)
31190     If UrineVolumeResult = "" Then Exit Function

31200     SerumCreatinineCode = GetOptionSetting("BIOCODEFORCREAT", "1068")

31210     sql = "SELECT * FROM Demographics WHERE SampleID = " & Val(txtSampleID) - 1
31220     Set tb = New Recordset
31230     RecOpenServer 0, tb, sql
31240     If tb.EOF Then Exit Function

31250     If UCase(txtSurName & " " & txtForeName) <> UCase(tb!PatName & "") Or txtDoB <> tb!DoB Then
31260         Exit Function
31270     End If

31280     Set BRsPrev = BRsPrev.Load("Bio", Val(txtSampleID) - 1, "Results", gDONTCARE, gDONTCARE, , , Trim$(txtSex), Trim$(txtDoB))
31290     For Each BR In BRsPrev
31300         If UCase(BR.Code) = SerumCreatinineCode Then
31310             SerumCreatinineResult = BR.Result
31320             Exit For
31330         End If
31340     Next
31350     If SerumCreatinineResult = "" Then Exit Function







31360     CreatinineClearanceResult = (Val(UrineCreatinineResult) * 1000 * Val(UrineVolumeResult)) / (Val(SerumCreatinineResult) * 1440)
31370     Call AddTestInDbnCollection(BRs, CreatinineClearanceCode, CreatinineClearanceResult)
31380     CheckCreatinineClearance = True


31390     Exit Function


CheckUrineCreatinineRatio_Error:

          Dim strES As String
          Dim intEL As Integer

31400     intEL = Erl
31410     strES = Err.Description
31420     LogError "frmEditAll", "CheckUrineCreatinineRatio", intEL, strES

End Function



'---------------------------------------------------------------------------------------
' Procedure : CheckUrineCreatinineRatio
' Author    : Masood
' Date      : 03/Mar/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CheckUrineCreatinineRatio(ByVal BRs As BIEResults) As Boolean
          'returns True if ACR added

          Dim BR As BIEResult

          'Dim Code As String
          'Dim CodeValue As String

          Dim UrineCreatinineRatioCode As String
          Dim UrineCreatinineRatioResult As String

          Dim CaUCode As String
          Dim CaUResult As String

          Dim UrineCreatinineCode As String
          Dim UrineCreatinineResult As String



31430     On Error GoTo CheckUrineCreatinineRatio_Error


31440     CheckUrineCreatinineRatio = False


          'If no results available then exit
31450     If BRs Is Nothing Then Exit Function

          'get code for and check if value exists for that code, if not then exit
31460     UrineCreatinineRatioCode = GetOptionSetting("BioCodeForUrineCreatinineRatio", "CCR")
31470     UrineCreatinineRatioResult = CheckResultsValue(BRs, UrineCreatinineRatioCode)
31480     If UrineCreatinineRatioResult <> "" Then Exit Function

          'get code for and check if value doesn't exists for that code, if not then exit
31490     CaUCode = GetOptionSetting("BioCodeForUrineCa", "1097")
31500     CaUResult = CheckResultsValue(BRs, CaUCode)
31510     If CaUResult = "" Then Exit Function

          'get code for and check if value doesn't exists for that code, if not then exit
31520     UrineCreatinineCode = GetOptionSetting("BIOCODEFORUCREAT", "1096")
31530     UrineCreatinineResult = CheckResultsValue(BRs, UrineCreatinineCode)
31540     If UrineCreatinineResult = "" Then Exit Function


31550     UrineCreatinineRatioResult = (Val(CaUResult) * 1000) / Val(UrineCreatinineResult)

31560     Call AddTestInDbnCollection(BRs, UrineCreatinineRatioCode, UrineCreatinineRatioResult)




31570     Exit Function


CheckUrineCreatinineRatio_Error:

          Dim strES As String
          Dim intEL As Integer

31580     intEL = Erl
31590     strES = Err.Description
31600     LogError "frmEditAll", "CheckUrineCreatinineRatio", intEL, strES

End Function








'---------------------------------------------------------------------------------------
' Procedure : AddTestInDbnCollection
' Author    : Masood
' Date      : 23/Feb/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function AddTestInDbnCollection(ByVal BRs As BIEResults, Code As String, Result As String) As Boolean

          Dim tbAd As Recordset
          Dim tb As Recordset
          Dim BR As BIEResult
          Dim sql As String
          Dim Rundate As String
          Dim RunTime As String
          Dim bNew As BIEResult

31610     On Error GoTo AddTestInDbnCollection_Error

31620     For Each BR In BRs
31630         sql = "Select Top 1 * From BioTestDefinitions " & _
                  "Where Code = '" & Code & "'"
31640         Set tbAd = New Recordset
31650         RecOpenClient 0, tbAd, sql
31660         If Not tbAd.EOF Then
31670             If IsDate(BR.Rundate) Then
31680                 Rundate = BR.Rundate
31690             Else
31700                 Rundate = Format$(BR.RunTime, "dd/mmm/yyyy")
31710             End If
31720             RunTime = BR.RunTime
31730             sql = "SELECT * FROM BioResults WHERE " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "AND Code = '" & Code & "'"
31740             Set tb = New Recordset
31750             RecOpenClient 0, tb, sql
31760             If tb.EOF Then
31770                 tb.AddNew
31780             End If
31790             tb!SampleID = txtSampleID
31800             tb!Rundate = Rundate
31810             tb!RunTime = RunTime
31820             tb!Code = tbAd!Code
31830             tb!Result = Result
31840             tb!SampleType = tbAd!SampleType
31850             tb!Units = tbAd!Units
31860             tb!Printed = BR.Printed
31870             tb!Valid = BR.Valid
31880             tb!FAXed = 0
31890             tb!Analyser = ""
31900             tb.Update

31910             Set bNew = New BIEResult
31920             bNew.SampleID = txtSampleID
31930             bNew.Code = tbAd!Code
31940             bNew.Rundate = Rundate
31950             bNew.RunTime = RunTime
31960             bNew.Result = Result
31970             bNew.Units = IIf(IsNull(tbAd!Units), "", tbAd!Units)
31980             bNew.Printed = BR.Printed
31990             bNew.Valid = BR.Valid
32000             bNew.SampleType = tbAd!SampleType
32010             bNew.LongName = tbAd!LongName
32020             bNew.ShortName = tbAd!ShortName
32030             bNew.PlausibleLow = tbAd!PlausibleLow
32040             bNew.PlausibleHigh = tbAd!PlausibleHigh
32050             bNew.FlagLow = 0    'tb!FlagLow
32060             bNew.FlagHigh = 9999    'tb!FlagHigh
32070             bNew.Printformat = tbAd!DP

32080             BRs.Add bNew

32090             AddTestInDbnCollection = True
32100             Exit For
32110         End If
32120     Next



32130     Exit Function


AddTestInDbnCollection_Error:

          Dim strES As String
          Dim intEL As Integer

32140     intEL = Erl
32150     strES = Err.Description
32160     LogError "frmEditAll", "AddTestInDbnCollection", intEL, strES, sql

End Function



'---------------------------------------------------------------------------------------
' Procedure : CheckCodeExistsInResult
' Author    : Masood
' Date      : 17/Feb/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CheckCodeExistsInResult(ByVal BRs As BIEResults, Code As String) As Boolean
32170     On Error GoTo CheckCodeExistsInResult_Error

          Dim BR As BIEResult

32180     For Each BR In BRs
32190         If Code = UCase$(Trim$(BR.Code)) Then
32200             CheckCodeExistsInResult = True
32210             Exit Function
32220         End If
32230     Next


32240     Exit Function


CheckCodeExistsInResult_Error:

          Dim strES As String
          Dim intEL As Integer

32250     intEL = Erl
32260     strES = Err.Description
32270     LogError "frmEditAll", "CheckCodeExistsInResult", intEL, strES
End Function




'---------------------------------------------------------------------------------------
' Procedure : CheckResultsValue
' Author    : Masood
' Date      : 17/Feb/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CheckResultsValue(ByVal BRs As BIEResults, Code As String) As String
32280     On Error GoTo CheckResultsValue_Error

          Dim BR As BIEResult

32290     For Each BR In BRs
32300         If Code = UCase$(Trim$(BR.Code)) Then
32310             CheckResultsValue = BR.Result
32320             Exit Function
32330         End If
32340     Next


32350     Exit Function


CheckResultsValue_Error:

          Dim strES As String
          Dim intEL As Integer

32360     intEL = Erl
32370     strES = Err.Description
32380     LogError "frmEditAll", "CheckResultsValue", intEL, strES
End Function


'---------------------------------------------------------------------------------------
' Procedure : CalculateACR
' Author    : Masood
' Date      : 11/Feb/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : CalculateACR
' Author    : Masood
' Date      : 17/Feb/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CalculateACR(ByVal UrinaryCreatinine As String, ByVal Albumin As String) As String
32390     On Error GoTo CalculateACR_Error
          '+++ Junaid
32400     Exit Function
          '--- Junaid
32410     If Val(UrinaryCreatinine) = 0 Then
32420         CalculateACR = "XXXX"
32430     Else
32440         CalculateACR = Val(Albumin) / (Val(UrinaryCreatinine) / 1000)
32450     End If

32460     Exit Function


CalculateACR_Error:

          Dim strES As String
          Dim intEL As Integer

32470     intEL = Erl
32480     strES = Err.Description
32490     LogError "frmEditAll", "CalculateACR", intEL, strES
End Function





Public Property Get SavedDemoFromGPCom() As Boolean

32500     SavedDemoFromGPCom = m_bSavedDemoFromGPCom

End Property

Public Property Let SavedDemoFromGPCom(ByVal bSavedDemoFromGPCom As Boolean)

32510     m_bSavedDemoFromGPCom = bSavedDemoFromGPCom

End Property

Public Property Get CancelFromGpCom() As Boolean

32520     CancelFromGpCom = m_bCancelFromGpCom

End Property

Public Property Let CancelFromGpCom(ByVal bCancelFromGpCom As Boolean)

32530     m_bCancelFromGpCom = bCancelFromGpCom

End Property
Private Sub LoadRejectedSample()
32540     On Error GoTo ErrorHandler

          Dim sql As String
          Dim tb As Recordset
32550     sql = "Select * from BioResults where Code='REJ' and SampleID = '" & txtSampleID & "'"
32560     Set tb = New Recordset
32570     RecOpenServer 0, tb, sql
32580     If Not tb.EOF Then
32590         chkBioReject.Value = 1
32600         chkBioReject.BackColor = vbYellow
32610         chkBioReject.Enabled = False
32620     End If
32630     sql = "Select * from CoagResults where Code='REJ' and SampleID = '" & txtSampleID & "'"
32640     Set tb = New Recordset
32650     RecOpenServer 0, tb, sql
32660     If Not tb.EOF Then
32670         chkCoagReject.Value = 1
32680         chkCoagReject.BackColor = vbYellow
32690         chkCoagReject.Enabled = False
32700     End If
32710     sql = "Select * from Observations where Discipline='Haematology' AND comment='Haematology sample rejected' and SampleID = '" & txtSampleID & "'"
32720     Set tb = New Recordset
32730     RecOpenServer 0, tb, sql
32740     If Not tb.EOF Then
32750         chkHaemReject.Value = 1
32760         chkHaemReject.BackColor = vbYellow
32770         chkHaemReject.Enabled = False
32780     End If
32790     sql = "Select * from ExtResults where Analyte='REJ' and SampleID = '" & txtSampleID & "'"
32800     Set tb = New Recordset
32810     RecOpenServer 0, tb, sql
32820     If Not tb.EOF Then
32830         ChkExtReject.Value = 1
32840         ChkExtReject.BackColor = vbYellow
32850         ChkExtReject.Enabled = False
32860     End If
32870     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description
End Sub

Private Sub ShowSamples()
32880     On Error GoTo ShowSamples_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          Dim tbD As ADODB.Recordset
          Dim tbR As ADODB.Recordset
          Dim l_RequestID As String
          
32890     cmbOtherSamples.Clear
32900     sql = "Select IsNULL(RequestID,'') RequestID from ocmRequestDetails Where SampleID = '" & txtSampleID.Text & "'"
32910     Set tbR = New Recordset
32920     RecOpenServer 0, tbR, sql
32930     If Not tbR Is Nothing Then
32940         If Not tbR.EOF Then
32950             l_RequestID = tbR!RequestID
32960         End If
32970     End If
32980     sql = "Select IsNULL(SampleID,'') SampleID from ocmRequestDetails "
32990     sql = sql & "Where RequestID = '" & l_RequestID & "'"
33000     Set tb = New Recordset
33010     RecOpenServer 0, tb, sql
33020     If Not tb Is Nothing Then
33030         If Not tb.EOF Then
33040             cmbOtherSamples.AddItem ("Sample ID    Sample Date                    Received Date")
33050             While Not tb.EOF
33060                 sql = "Select SampleID, SampleDate, RecDate from Demographics Where SampleID = '" & tb!SampleID & "'"
33070                 Set tbD = New Recordset
33080                 RecOpenServer 0, tbD, sql
33090                 If Not tbD Is Nothing Then
33100                     If Not tbD.EOF Then
33110                         cmbOtherSamples.AddItem (Right("0000000" & tbD!SampleID, 7) & " ---- " & tbD!SampleDate & " ---- " & tbD!RecDate)
33120                     End If
33130                 End If
33140                 tb.MoveNext
33150             Wend
33160         End If
33170     End If
33180     Exit Sub

ShowSamples_Error:
          
          Dim strES As String
          Dim intEL As Integer

33190     intEL = Erl
33200     strES = Err.Description
33210     LogError "frmEditAll", "ShowSamples", intEL, strES, sql
End Sub

Private Sub cmbOtherSamples_Change()
33220     On Error GoTo cmbOtherSamples_Change_Error

33230     If Left(cmbOtherSamples.Text, 6) = "Sample" Then
33240         Call txtsampleid_LostFocus
33250     Else
33260         txtSampleID.Text = Left(cmbOtherSamples.Text, 7)
33270         Call txtsampleid_LostFocus
33280     End If

33290     Exit Sub

cmbOtherSamples_Change_Error:
          
          Dim strES As String
          Dim intEL As Integer

33300     intEL = Erl
33310     strES = Err.Description
33320     LogError "frmEditAll", "cmbOtherSamples_Change", intEL, strES
End Sub

Private Sub FormatGrid()
33330     On Error GoTo ERROR_FormatGrid
          
33340     flxQuestions.Rows = 1
33350     flxQuestions.row = 0
          
33360     flxQuestions.ColWidth(fcLine_NO) = 100
          
33370     flxQuestions.TextMatrix(0, fcQus) = "Questions"
33380     flxQuestions.ColWidth(fcQus) = 3000
33390     flxQuestions.ColAlignment(fcQus) = flexAlignLeftCenter
          
33400     flxQuestions.TextMatrix(0, fcAns) = "Answers"
33410     flxQuestions.ColWidth(fcAns) = 3100
33420     flxQuestions.ColAlignment(fcAns) = flexAlignLeftCenter
          
33430     flxQuestions.TextMatrix(0, fcRID) = ""
33440     flxQuestions.ColWidth(fcRID) = 0
33450     flxQuestions.ColAlignment(fcRID) = flexAlignLeftCenter
          
          
33460     Exit Sub
ERROR_FormatGrid:
          Dim strES As String
          Dim intEL As Integer

33470     intEL = Erl
33480     strES = Err.Description
33490     LogError "frmEditAll", "FormatGrid", intEL, strES
End Sub

Private Sub ShowQuestions()
33500     On Error GoTo ShowQuestions_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          Dim tbR As ADODB.Recordset
          Dim l_RequestID As String
          Dim l_str As String
          
33510     flxQuestions.Rows = 1
33520     flxQuestions.row = 0
          
33530     sql = "Select IsNULL(RequestID,'') RequestID from ocmRequestDetails Where SampleID = '" & txtSampleID.Text & "'"
33540     Set tbR = New Recordset
33550     RecOpenServer 0, tbR, sql
33560     If Not tbR Is Nothing Then
33570         If Not tbR.EOF Then
33580             l_RequestID = tbR!RequestID
33590         End If
33600     End If
33610     sql = "Select * from ocmQuestions "
33620     sql = sql & "Where RID = '" & l_RequestID & "'"
33630     Set tb = New Recordset
33640     RecOpenServer 0, tb, sql
33650     If Not tb Is Nothing Then
33660         If Not tb.EOF Then
33670             While Not tb.EOF
33680                 l_str = "" & vbTab & tb!Question & vbTab & tb!Answer & vbTab & tb!RID
33690                 flxQuestions.AddItem (l_str)
33700                 tb.MoveNext
33710             Wend
33720         End If
33730     End If
33740     Exit Sub

ShowQuestions_Error:
          
          Dim strES As String
          Dim intEL As Integer

33750     intEL = Erl
33760     strES = Err.Description
33770     LogError "frmEditAll", "ShowQuestions", intEL, strES, sql
End Sub

Private Sub DeleteOutstandings()
33780     On Error GoTo DeleteOutstandings_Error

          Dim sql As String
          
33790     sql = "Delete from HaeRequests Where SampleID = '" & txtSampleID.Text & "'"
33800     Cnxn(0).Execute sql
          
33810     Exit Sub

DeleteOutstandings_Error:
          
          Dim strES As String
          Dim intEL As Integer

33820     intEL = Erl
33830     strES = Err.Description
33840     LogError "frmEditAll", "DeleteOutstandings", intEL, strES, sql
End Sub

Private Function GetRequestID(p_SampleID As String) As String
33850     On Error GoTo GetRequestID_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          
33860     GetRequestID = ""
33870     sql = "Select Distinct IsNULL(RequestID,'') RequestID from ocmRequestDetails Where SampleID = '" & p_SampleID & "'"
33880     Set tb = New Recordset
33890     RecOpenServer 0, tb, sql
33900     If Not tb Is Nothing Then
33910         If Not tb.EOF Then
33920             GetRequestID = "Request ID: " & tb!RequestID
33930         End If
33940     End If
          
33950     Exit Function

GetRequestID_Error:
          
          Dim strES As String
          Dim intEL As Integer
          
33960     GetRequestID = ""
33970     intEL = Erl
33980     strES = Err.Description
33990     LogError "frmEditAll", "GetRequestID", intEL, strES, sql
End Function

'+++ Junaid 29-02-2024
Private Function CheckDemographics() As Boolean
34000     On Error GoTo CheckDemographics_Error

34010     CheckDemographics = True
          
34020     If txtSampleID.Text = "" Then
34030         CheckDemographics = False
34040         Exit Function
34050     End If
          
34060     If txtForeName.Text = "" Then
34070         CheckDemographics = False
34080         Exit Function
34090     End If
          
34100     If txtSurName.Text = "" Then
34110         CheckDemographics = False
34120         Exit Function
34130     End If
          
34140     If txtDoB.Text = "" Then
34150         CheckDemographics = False
34160         Exit Function
34170     End If
          
34180     If txtAge.Text = "" Then
34190         CheckDemographics = False
34200         Exit Function
34210     End If
          
34220     If txtSex.Text = "" Then
34230         CheckDemographics = False
34240         Exit Function
34250     End If
          
34260     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
34270         CheckDemographics = False
34280         Exit Function
34290     End If
          
34300     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then CheckDemographics = False
          
34310     Exit Function

CheckDemographics_Error:
          
          Dim strES As String
          Dim intEL As Integer
          
34320     CheckDemographics = False
34330     intEL = Erl
34340     strES = Err.Description
34350     LogError "frmEditAll", "CheckDemographics", intEL, strES
End Function
'--- Junaid

