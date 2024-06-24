VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOldEditMicrobiology 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   765
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   583
      ToolTipText     =   "Log as Phoned"
      Top             =   390
      Width           =   1275
   End
   Begin VB.TextBox txtNOPAS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13860
      TabIndex        =   577
      Top             =   1230
      Width           =   1245
   End
   Begin VB.TextBox txtAandE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13800
      TabIndex        =   576
      Top             =   360
      Width           =   1245
   End
   Begin VB.CommandButton cmdSaveHold 
      Caption         =   "Save && Hold"
      Height          =   645
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   487
      Top             =   5100
      Width           =   1275
   End
   Begin VB.CommandButton cmdValidateMicro 
      Caption         =   "&Validate && Print"
      Height          =   765
      Left            =   12150
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmOldEditBacteriology.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   415
      Top             =   6570
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveMicro 
      Caption         =   "&Save Details"
      Enabled         =   0   'False
      Height          =   705
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   414
      Top             =   5760
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   615
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   11970
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   390
      TabIndex        =   58
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.CommandButton bPrintHold 
      Caption         =   "Print && Hold"
      Height          =   615
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":1502
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2850
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2085
      Index           =   0
      Left            =   390
      TabIndex        =   29
      Top             =   150
      Width           =   11565
      Begin VB.CommandButton cmdAddToConsultantList 
         Caption         =   "Remove from  Consultant List"
         Height          =   255
         Left            =   60
         TabIndex        =   582
         Top             =   1410
         Width           =   2325
      End
      Begin VB.ComboBox cmbConsultantVal 
         Height          =   315
         Left            =   570
         TabIndex        =   580
         Text            =   "cmbConsultantVal"
         Top             =   1680
         Width           =   1605
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
         Left            =   2730
         MaxLength       =   8
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox txtName 
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
         Left            =   4380
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "tName"
         Top             =   570
         Width           =   4215
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   3
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   9240
         MaxLength       =   4
         TabIndex        =   4
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   9240
         MaxLength       =   6
         TabIndex        =   5
         Top             =   990
         Width           =   1545
      End
      Begin VB.Frame Frame6 
         Height          =   1395
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   2385
         Begin VB.ComboBox cMRU 
            Height          =   315
            Left            =   570
            TabIndex        =   59
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
            Left            =   150
            MaxLength       =   12
            TabIndex        =   0
            Top             =   510
            Width           =   1785
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   480
            Left            =   1920
            TabIndex        =   33
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   847
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtSampleID"
            BuddyDispid     =   196628
            OrigLeft        =   1920
            OrigTop         =   540
            OrigRight       =   2160
            OrigBottom      =   1020
            Max             =   99999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "MRU"
            Height          =   195
            Left            =   150
            TabIndex        =   60
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image iRelevant 
            Height          =   480
            Index           =   1
            Left            =   1500
            Picture         =   "frmOldEditBacteriology.frx":1B6C
            Top             =   120
            Width           =   480
         End
         Begin VB.Image iRelevant 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   120
            Picture         =   "frmOldEditBacteriology.frx":1E76
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Sample ID"
            Height          =   195
            Left            =   720
            TabIndex        =   34
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Searc&h"
         Height          =   345
         Left            =   7920
         TabIndex        =   31
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   285
         Left            =   10800
         TabIndex        =   30
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblABsInUse 
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Left            =   9240
         TabIndex        =   223
         Top             =   1350
         Width           =   2235
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2730
         TabIndex        =   222
         Top             =   1710
         Width           =   5865
      End
      Begin VB.Label lblSiteDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2730
         TabIndex        =   218
         Top             =   1410
         Width           =   5865
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   2730
         TabIndex        =   64
         ToolTipText     =   "Click to change Location"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label lAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2730
         TabIndex        =   63
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
         Left            =   6120
         TabIndex        =   61
         Top             =   210
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   4440
         TabIndex        =   38
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   0
         Left            =   8790
         TabIndex        =   37
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   8880
         TabIndex        =   36
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   8910
         TabIndex        =   35
         Top             =   1020
         Width           =   270
      End
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Height          =   885
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":2180
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1230
      Width           =   1275
   End
   Begin VB.CommandButton bHistory 
      Caption         =   "&History"
      Height          =   675
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":248A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":28CC
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "bprint"
      Top             =   3480
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "FAX"
      Height          =   825
      Index           =   0
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":2F36
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4230
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   12150
      Picture         =   "frmOldEditBacteriology.frx":3378
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8010
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   390
      TabIndex        =   15
      Top             =   2340
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   15
      TabsPerRow      =   15
      TabHeight       =   529
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmOldEditBacteriology.frx":39E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
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
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Urine"
      TabPicture(1)   =   "frmOldEditBacteriology.frx":39FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPregnancy"
      Tab(1).Control(1)=   "fraUrineSpecific"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "frDipStick"
      Tab(1).Control(4)=   "txtUrineComment"
      Tab(1).Control(5)=   "Label10"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Identification"
      TabPicture(2)   =   "frmOldEditBacteriology.frx":3A1A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameExtras(4)"
      Tab(2).Control(1)=   "FrameExtras(3)"
      Tab(2).Control(2)=   "FrameExtras(2)"
      Tab(2).Control(3)=   "FrameExtras(1)"
      Tab(2).Control(4)=   "lblMoreID"
      Tab(2).Control(5)=   "imgMoreID"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Faeces"
      TabPicture(3)   =   "frmOldEditBacteriology.frx":3A36
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "F0157"
      Tab(3).Control(2)=   "FrOva"
      Tab(3).Control(3)=   "frRotaAdeno"
      Tab(3).Control(4)=   "frOccult"
      Tab(3).Control(5)=   "frEPC"
      Tab(3).Control(6)=   "Frame1(1)"
      Tab(3).Control(7)=   "Frame2(1)"
      Tab(3).Control(8)=   "fCulture"
      Tab(3).Control(9)=   "FToxinA"
      Tab(3).Control(10)=   "frCampylobacter"
      Tab(3).Control(11)=   "frAPI"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "C && S"
      TabPicture(4)   =   "frmOldEditBacteriology.frx":3A52
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmbQualifier(4)"
      Tab(4).Control(1)=   "cmbQualifier(3)"
      Tab(4).Control(2)=   "cmbQualifier(2)"
      Tab(4).Control(3)=   "cmbQualifier(1)"
      Tab(4).Control(4)=   "txtConsultantComment(0)"
      Tab(4).Control(5)=   "cmdUnlock(4)"
      Tab(4).Control(6)=   "cmdUnlock(3)"
      Tab(4).Control(7)=   "cmdUnlock(2)"
      Tab(4).Control(8)=   "cmdUnlock(1)"
      Tab(4).Control(9)=   "cmdUseSecondary(4)"
      Tab(4).Control(10)=   "cmdRemoveSecondary(4)"
      Tab(4).Control(11)=   "cmdUseSecondary(3)"
      Tab(4).Control(12)=   "cmdRemoveSecondary(3)"
      Tab(4).Control(13)=   "cmdUseSecondary(2)"
      Tab(4).Control(14)=   "cmdRemoveSecondary(2)"
      Tab(4).Control(15)=   "cmdUseSecondary(1)"
      Tab(4).Control(16)=   "cmdRemoveSecondary(1)"
      Tab(4).Control(17)=   "cmbABSelect(4)"
      Tab(4).Control(18)=   "cmbABSelect(3)"
      Tab(4).Control(19)=   "cmbABSelect(2)"
      Tab(4).Control(20)=   "cmbABSelect(1)"
      Tab(4).Control(21)=   "cmbOrgName(4)"
      Tab(4).Control(22)=   "cmbOrgName(3)"
      Tab(4).Control(23)=   "cmbOrgName(2)"
      Tab(4).Control(24)=   "cmbOrgName(1)"
      Tab(4).Control(25)=   "cmbOrgGroup(4)"
      Tab(4).Control(26)=   "txtCSComment(0)"
      Tab(4).Control(27)=   "cmbOrgGroup(1)"
      Tab(4).Control(28)=   "cmbOrgGroup(2)"
      Tab(4).Control(29)=   "cmbOrgGroup(3)"
      Tab(4).Control(30)=   "grdAB(4)"
      Tab(4).Control(31)=   "grdAB(3)"
      Tab(4).Control(32)=   "grdAB(2)"
      Tab(4).Control(33)=   "grdAB(1)"
      Tab(4).Control(34)=   "Label48(0)"
      Tab(4).Control(35)=   "lblSetAllR(4)"
      Tab(4).Control(36)=   "lblSetAllS(4)"
      Tab(4).Control(37)=   "lblSetAllR(3)"
      Tab(4).Control(38)=   "lblSetAllS(3)"
      Tab(4).Control(39)=   "lblSetAllR(2)"
      Tab(4).Control(40)=   "lblSetAllS(2)"
      Tab(4).Control(41)=   "lblMoreCS"
      Tab(4).Control(42)=   "imgMoreCS"
      Tab(4).Control(43)=   "Label46(3)"
      Tab(4).Control(44)=   "Label46(2)"
      Tab(4).Control(45)=   "Label46(1)"
      Tab(4).Control(46)=   "Label46(0)"
      Tab(4).Control(47)=   "lblSetAllS(1)"
      Tab(4).Control(48)=   "lblSetAllR(1)"
      Tab(4).Control(49)=   "imgSquareTick"
      Tab(4).Control(50)=   "imgSquareCross"
      Tab(4).Control(51)=   "Label20(0)"
      Tab(4).ControlCount=   52
      TabCaption(5)   =   "Identification 5/8"
      TabPicture(5)   =   "frmOldEditBacteriology.frx":3A6E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameExtras(7)"
      Tab(5).Control(1)=   "FrameExtras(6)"
      Tab(5).Control(2)=   "FrameExtras(5)"
      Tab(5).Control(3)=   "FrameExtras(8)"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "C && S 5/8"
      TabPicture(6)   =   "frmOldEditBacteriology.frx":3A8A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmbQualifier(8)"
      Tab(6).Control(1)=   "cmbQualifier(7)"
      Tab(6).Control(2)=   "cmbQualifier(6)"
      Tab(6).Control(3)=   "cmbQualifier(5)"
      Tab(6).Control(4)=   "txtConsultantComment(1)"
      Tab(6).Control(5)=   "cmdUnlock(8)"
      Tab(6).Control(6)=   "cmdUnlock(7)"
      Tab(6).Control(7)=   "cmdUnlock(6)"
      Tab(6).Control(8)=   "cmdUnlock(5)"
      Tab(6).Control(9)=   "cmdUseSecondary(8)"
      Tab(6).Control(10)=   "cmdRemoveSecondary(8)"
      Tab(6).Control(11)=   "cmdUseSecondary(7)"
      Tab(6).Control(12)=   "cmdRemoveSecondary(7)"
      Tab(6).Control(13)=   "cmdUseSecondary(6)"
      Tab(6).Control(14)=   "cmdRemoveSecondary(6)"
      Tab(6).Control(15)=   "cmdUseSecondary(5)"
      Tab(6).Control(16)=   "cmdRemoveSecondary(5)"
      Tab(6).Control(17)=   "cmbABSelect(8)"
      Tab(6).Control(18)=   "cmbABSelect(7)"
      Tab(6).Control(19)=   "cmbABSelect(6)"
      Tab(6).Control(20)=   "cmbABSelect(5)"
      Tab(6).Control(21)=   "cmbOrgName(8)"
      Tab(6).Control(22)=   "cmbOrgName(7)"
      Tab(6).Control(23)=   "cmbOrgName(6)"
      Tab(6).Control(24)=   "cmbOrgName(5)"
      Tab(6).Control(25)=   "cmbOrgGroup(8)"
      Tab(6).Control(26)=   "cmbOrgGroup(7)"
      Tab(6).Control(27)=   "cmbOrgGroup(6)"
      Tab(6).Control(28)=   "txtCSComment(1)"
      Tab(6).Control(29)=   "cmbOrgGroup(5)"
      Tab(6).Control(30)=   "grdAB(5)"
      Tab(6).Control(31)=   "grdAB(6)"
      Tab(6).Control(32)=   "grdAB(7)"
      Tab(6).Control(33)=   "grdAB(8)"
      Tab(6).Control(34)=   "Label48(1)"
      Tab(6).Control(35)=   "lblSetAllR(8)"
      Tab(6).Control(36)=   "lblSetAllS(8)"
      Tab(6).Control(37)=   "lblSetAllR(7)"
      Tab(6).Control(38)=   "lblSetAllS(7)"
      Tab(6).Control(39)=   "lblSetAllR(6)"
      Tab(6).Control(40)=   "lblSetAllS(6)"
      Tab(6).Control(41)=   "lblSetAllR(5)"
      Tab(6).Control(42)=   "lblSetAllS(5)"
      Tab(6).Control(43)=   "Label46(7)"
      Tab(6).Control(44)=   "Label46(6)"
      Tab(6).Control(45)=   "Label46(5)"
      Tab(6).Control(46)=   "Label20(3)"
      Tab(6).Control(47)=   "Label46(4)"
      Tab(6).ControlCount=   48
      TabCaption(7)   =   "MRSA"
      TabPicture(7)   =   "frmOldEditBacteriology.frx":3AA6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "FOB"
      TabPicture(8)   =   "frmOldEditBacteriology.frx":3AC2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame8"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Rota/Adeno"
      TabPicture(9)   =   "frmOldEditBacteriology.frx":3ADE
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame9"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "C.diff"
      TabPicture(10)  =   "frmOldEditBacteriology.frx":3AFA
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame11"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "O/P"
      TabPicture(11)  =   "frmOldEditBacteriology.frx":3B16
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame15"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "Identification"
      TabPicture(12)  =   "frmOldEditBacteriology.frx":3B32
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lblMoreIdentity"
      Tab(12).Control(1)=   "imgMoreIdentity"
      Tab(12).Control(2)=   "fraIdentification(0)"
      Tab(12).Control(3)=   "fraIdentification(1)"
      Tab(12).Control(4)=   "fraIdentification(2)"
      Tab(12).Control(5)=   "fraIdentification(3)"
      Tab(12).ControlCount=   6
      TabCaption(13)  =   "Identification 5/8"
      TabPicture(13)  =   "frmOldEditBacteriology.frx":3B4E
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "fraIdentification(4)"
      Tab(13).Control(1)=   "fraIdentification(5)"
      Tab(13).Control(2)=   "fraIdentification(6)"
      Tab(13).Control(3)=   "fraIdentification(7)"
      Tab(13).ControlCount=   4
      TabCaption(14)  =   "RSV"
      TabPicture(14)  =   "frmOldEditBacteriology.frx":3B6A
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Frame16"
      Tab(14).ControlCount=   1
      Begin VB.Frame Frame16 
         Caption         =   "RSV"
         Height          =   1275
         Left            =   -71070
         TabIndex        =   638
         Top             =   1620
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
            TabIndex        =   639
            Top             =   450
            Width           =   2205
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 8"
         Height          =   5445
         Index           =   7
         Left            =   -66900
         TabIndex        =   635
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   8
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   637
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   636
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 7"
         Height          =   5445
         Index           =   6
         Left            =   -69480
         TabIndex        =   632
         Top             =   540
         Width           =   2505
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   634
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   7
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   633
            Top             =   600
            Width           =   2355
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 6"
         Height          =   5445
         Index           =   5
         Left            =   -72060
         TabIndex        =   629
         Top             =   540
         Width           =   2505
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   631
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   6
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   630
            Top             =   600
            Width           =   2355
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 5"
         Height          =   5445
         Index           =   4
         Left            =   -74670
         TabIndex        =   626
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   5
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   628
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   627
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 4"
         Height          =   5445
         Index           =   3
         Left            =   -66900
         TabIndex        =   623
         Top             =   540
         Width           =   2505
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   625
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   4
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   624
            Top             =   600
            Width           =   2355
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 3"
         Height          =   5445
         Index           =   2
         Left            =   -69480
         TabIndex        =   620
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   3
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   622
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   621
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 2"
         Height          =   5445
         Index           =   1
         Left            =   -72060
         TabIndex        =   617
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   2
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   619
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   618
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraIdentification 
         Caption         =   "Organism 1"
         Height          =   5445
         Index           =   0
         Left            =   -74670
         TabIndex        =   614
         Top             =   540
         Width           =   2505
         Begin VB.ComboBox cmbIdentification 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   616
            Text            =   "cmbIdentification"
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtIdentification 
            Height          =   4695
            Index           =   1
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   615
            Top             =   600
            Width           =   2355
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Ova / Parasites"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   -72270
         TabIndex        =   607
         Top             =   1440
         Width           =   6195
         Begin VB.ComboBox cmbOva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   510
            TabIndex        =   612
            Top             =   2640
            Width           =   5055
         End
         Begin VB.ComboBox cmbOva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   510
            TabIndex        =   611
            Top             =   2070
            Width           =   5055
         End
         Begin VB.ComboBox cmbOva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   510
            TabIndex        =   610
            Top             =   1500
            Width           =   5055
         End
         Begin VB.Label lblCrypto 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2730
            TabIndex        =   609
            Top             =   780
            Width           =   2805
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Cryptosporidium"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   540
            TabIndex        =   608
            Top             =   810
            Width           =   1710
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "C.diff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   -71820
         TabIndex        =   602
         Top             =   1650
         Width           =   4755
         Begin VB.Label lblToxinB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   606
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label lblToxinA 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   605
            Top             =   780
            Width           =   1905
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Toxin B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   810
            TabIndex        =   604
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Toxin A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   810
            TabIndex        =   603
            Top             =   780
            Width           =   780
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Rota/Adeno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -71310
         TabIndex        =   597
         Top             =   1500
         Width           =   3735
         Begin VB.TextBox txtAdeno 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1500
            TabIndex        =   601
            Top             =   1440
            Width           =   1395
         End
         Begin VB.TextBox txtRota 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1500
            TabIndex        =   600
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Adeno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   480
            TabIndex        =   599
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Rota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   630
            TabIndex        =   598
            Top             =   690
            Width           =   525
         End
      End
      Begin VB.Frame Frame8 
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
         Left            =   -71430
         TabIndex        =   590
         Top             =   1710
         Width           =   4245
         Begin VB.CheckBox chkFOB 
            Alignment       =   1  'Right Justify
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   660
            TabIndex        =   595
            Top             =   1560
            Width           =   405
         End
         Begin VB.CheckBox chkFOB 
            Alignment       =   1  'Right Justify
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   593
            Top             =   1020
            Width           =   405
         End
         Begin VB.CheckBox chkFOB 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   591
            Top             =   510
            Width           =   405
         End
         Begin VB.Label lblFOB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   1350
            TabIndex        =   596
            Top             =   1560
            Width           =   2025
         End
         Begin VB.Label lblFOB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1350
            TabIndex        =   594
            Top             =   1020
            Width           =   2025
         End
         Begin VB.Label lblFOB 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1350
            TabIndex        =   592
            Top             =   480
            Width           =   2025
         End
      End
      Begin VB.Frame fraPregnancy 
         Caption         =   "Pregnancy"
         Height          =   1245
         Left            =   -70830
         TabIndex        =   584
         Top             =   780
         Width           =   3855
         Begin VB.TextBox txtHCGLevel 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   586
            Top             =   750
            Width           =   1545
         End
         Begin VB.TextBox txtPregnancy 
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            MultiLine       =   -1  'True
            TabIndex        =   585
            ToolTipText     =   "P-Positive N-Negative E-Equivocal U-Unsuitable"
            Top             =   420
            Width           =   2055
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "HCG Level"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   589
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Pregnancy Test"
            Height          =   195
            Left            =   360
            TabIndex        =   588
            Top             =   450
            Width           =   1155
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "IU/L"
            Height          =   195
            Left            =   3120
            TabIndex        =   587
            Top             =   780
            Width           =   330
         End
      End
      Begin VB.CommandButton cmdCopyFromPrevious 
         BackColor       =   &H00FF80FF&
         Caption         =   "Copy all Details from Sample # 123456789"
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   581
         Top             =   480
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   8
         Left            =   -66300
         TabIndex        =   572
         Text            =   "cmbQualifier"
         Top             =   1140
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   7
         Left            =   -69030
         TabIndex        =   571
         Text            =   "cmbQualifier"
         Top             =   1140
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   6
         Left            =   -71760
         TabIndex        =   570
         Text            =   "cmbQualifier"
         Top             =   1140
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   5
         Left            =   -74490
         TabIndex        =   569
         Text            =   "cmbQualifier"
         Top             =   1140
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   4
         Left            =   -66270
         TabIndex        =   568
         Text            =   "cmbQualifier"
         Top             =   1320
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   3
         Left            =   -69000
         TabIndex        =   567
         Text            =   "cmbQualifier"
         Top             =   1320
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   2
         Left            =   -71730
         TabIndex        =   566
         Text            =   "cmbQualifier"
         Top             =   1320
         Width           =   2205
      End
      Begin VB.ComboBox cmbQualifier 
         Height          =   315
         Index           =   1
         Left            =   -74460
         TabIndex        =   565
         Text            =   "cmbQualifier"
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 7"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   7
         Left            =   -69480
         TabIndex        =   533
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   7
            Left            =   60
            TabIndex        =   534
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   7
            Left            =   885
            TabIndex        =   550
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   7
            Left            =   885
            TabIndex        =   549
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   7
            Left            =   915
            TabIndex        =   548
            Top             =   690
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   7
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   547
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   7
            Left            =   915
            TabIndex        =   546
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   7
            Left            =   915
            TabIndex        =   545
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   7
            Left            =   915
            TabIndex        =   544
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   7
            Left            =   915
            TabIndex        =   543
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   7
            Left            =   885
            TabIndex        =   542
            Top             =   5220
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   7
            Left            =   885
            TabIndex        =   541
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   7
            Left            =   885
            TabIndex        =   540
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   7
            Left            =   885
            TabIndex        =   539
            Top             =   3930
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   7
            Left            =   885
            MaxLength       =   10
            TabIndex        =   538
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   7
            Left            =   885
            MaxLength       =   10
            TabIndex        =   537
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   7
            Left            =   900
            TabIndex        =   536
            Top             =   2220
            Width           =   1515
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   7
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   535
            Top             =   2580
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   563
            Top             =   3090
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   562
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   1
            Left            =   450
            TabIndex        =   561
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   560
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   559
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   558
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   557
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   556
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   555
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   554
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   553
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   552
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   551
            Top             =   2250
            Width           =   465
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 4"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   4
         Left            =   -66900
         TabIndex        =   502
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   4
            Left            =   60
            TabIndex        =   503
            Top             =   3000
            Width           =   2355
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   4
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   519
            Top             =   2580
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   518
            Top             =   2220
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   4
            Left            =   900
            MaxLength       =   10
            TabIndex        =   517
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   4
            Left            =   885
            MaxLength       =   10
            TabIndex        =   516
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   4
            Left            =   885
            TabIndex        =   515
            Top             =   3930
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   4
            Left            =   885
            TabIndex        =   514
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   4
            Left            =   885
            TabIndex        =   513
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   4
            Left            =   885
            TabIndex        =   512
            Top             =   5220
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   511
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   510
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   509
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   508
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   4
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   507
            Top             =   360
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   4
            Left            =   915
            TabIndex        =   506
            Top             =   690
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   4
            Left            =   885
            TabIndex        =   505
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   4
            Left            =   885
            TabIndex        =   504
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   7
            Left            =   390
            TabIndex        =   532
            Top             =   2250
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   8
            Left            =   420
            TabIndex        =   531
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   8
            Left            =   390
            TabIndex        =   530
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   7
            Left            =   405
            TabIndex        =   529
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   7
            Left            =   300
            TabIndex        =   528
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   527
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   526
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   10
            Left            =   255
            TabIndex        =   525
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   9
            Left            =   300
            TabIndex        =   524
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   7
            Left            =   195
            TabIndex        =   523
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   7
            Left            =   450
            TabIndex        =   522
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   7
            Left            =   75
            TabIndex        =   521
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   520
            Top             =   3090
            Width           =   705
         End
      End
      Begin VB.TextBox txtConsultantComment 
         BackColor       =   &H80000018&
         Height          =   855
         Index           =   0
         Left            =   -69030
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   499
         Top             =   5370
         Width           =   4995
      End
      Begin VB.TextBox txtConsultantComment 
         BackColor       =   &H80000018&
         Height          =   855
         Index           =   1
         Left            =   -69060
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   498
         Top             =   5190
         Width           =   4995
      End
      Begin VB.CheckBox chkPregnant 
         Alignment       =   1  'Right Justify
         Caption         =   "Pregnant"
         Height          =   225
         Left            =   4350
         TabIndex        =   497
         Top             =   960
         Width           =   945
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   8
         Left            =   -66330
         Picture         =   "frmOldEditBacteriology.frx":3B86
         Style           =   1  'Graphical
         TabIndex        =   496
         Top             =   4620
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   7
         Left            =   -69060
         Picture         =   "frmOldEditBacteriology.frx":3FC8
         Style           =   1  'Graphical
         TabIndex        =   495
         Top             =   4620
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   6
         Left            =   -71790
         Picture         =   "frmOldEditBacteriology.frx":440A
         Style           =   1  'Graphical
         TabIndex        =   494
         Top             =   4620
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   5
         Left            =   -74520
         Picture         =   "frmOldEditBacteriology.frx":484C
         Style           =   1  'Graphical
         TabIndex        =   493
         Top             =   4620
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   4
         Left            =   -66300
         Picture         =   "frmOldEditBacteriology.frx":4C8E
         Style           =   1  'Graphical
         TabIndex        =   492
         Top             =   4800
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   3
         Left            =   -69030
         Picture         =   "frmOldEditBacteriology.frx":50D0
         Style           =   1  'Graphical
         TabIndex        =   491
         Top             =   4800
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   2
         Left            =   -71760
         Picture         =   "frmOldEditBacteriology.frx":5512
         Style           =   1  'Graphical
         TabIndex        =   490
         Top             =   4800
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUnlock 
         Height          =   555
         Index           =   1
         Left            =   -74490
         Picture         =   "frmOldEditBacteriology.frx":5954
         Style           =   1  'Graphical
         TabIndex        =   489
         Top             =   4800
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   8
         Left            =   -66720
         Picture         =   "frmOldEditBacteriology.frx":5D96
         Style           =   1  'Graphical
         TabIndex        =   486
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3270
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   8
         Left            =   -66720
         Picture         =   "frmOldEditBacteriology.frx":60A0
         Style           =   1  'Graphical
         TabIndex        =   485
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2730
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   7
         Left            =   -69450
         Picture         =   "frmOldEditBacteriology.frx":63AA
         Style           =   1  'Graphical
         TabIndex        =   484
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3270
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   7
         Left            =   -69450
         Picture         =   "frmOldEditBacteriology.frx":66B4
         Style           =   1  'Graphical
         TabIndex        =   483
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2730
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   6
         Left            =   -72180
         Picture         =   "frmOldEditBacteriology.frx":69BE
         Style           =   1  'Graphical
         TabIndex        =   482
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3270
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   6
         Left            =   -72180
         Picture         =   "frmOldEditBacteriology.frx":6CC8
         Style           =   1  'Graphical
         TabIndex        =   481
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2730
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   5
         Left            =   -74910
         Picture         =   "frmOldEditBacteriology.frx":6FD2
         Style           =   1  'Graphical
         TabIndex        =   480
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3270
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   5
         Left            =   -74910
         Picture         =   "frmOldEditBacteriology.frx":72DC
         Style           =   1  'Graphical
         TabIndex        =   479
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2730
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   4
         Left            =   -66690
         Picture         =   "frmOldEditBacteriology.frx":75E6
         Style           =   1  'Graphical
         TabIndex        =   478
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   4
         Left            =   -66690
         Picture         =   "frmOldEditBacteriology.frx":78F0
         Style           =   1  'Graphical
         TabIndex        =   477
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2910
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   3
         Left            =   -69420
         Picture         =   "frmOldEditBacteriology.frx":7BFA
         Style           =   1  'Graphical
         TabIndex        =   476
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   3
         Left            =   -69420
         Picture         =   "frmOldEditBacteriology.frx":7F04
         Style           =   1  'Graphical
         TabIndex        =   475
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2910
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   2
         Left            =   -72150
         Picture         =   "frmOldEditBacteriology.frx":820E
         Style           =   1  'Graphical
         TabIndex        =   474
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   2
         Left            =   -72150
         Picture         =   "frmOldEditBacteriology.frx":8518
         Style           =   1  'Graphical
         TabIndex        =   473
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2910
         Width           =   375
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   1
         Left            =   -74880
         Picture         =   "frmOldEditBacteriology.frx":8822
         Style           =   1  'Graphical
         TabIndex        =   472
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   1
         Left            =   -74880
         Picture         =   "frmOldEditBacteriology.frx":8B2C
         Style           =   1  'Graphical
         TabIndex        =   471
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   2910
         Width           =   375
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   8
         Left            =   -66270
         TabIndex        =   463
         Text            =   "cmbABSelect"
         Top             =   4620
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   7
         Left            =   -69000
         TabIndex        =   462
         Text            =   "cmbABSelect"
         Top             =   4620
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   6
         Left            =   -71730
         TabIndex        =   461
         Text            =   "cmbABSelect"
         Top             =   4620
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   5
         Left            =   -74460
         TabIndex        =   460
         Text            =   "cmbABSelect"
         Top             =   4620
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   4
         Left            =   -66240
         TabIndex        =   459
         Text            =   "cmbABSelect"
         Top             =   4800
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   3
         Left            =   -68970
         TabIndex        =   458
         Text            =   "cmbABSelect"
         Top             =   4800
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   2
         Left            =   -71700
         TabIndex        =   457
         Text            =   "cmbABSelect"
         Top             =   4800
         Width           =   2115
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   1
         Left            =   -74430
         TabIndex        =   456
         Text            =   "cmbABSelect"
         Top             =   4800
         Width           =   2085
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   8
         Left            =   -66300
         TabIndex        =   441
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   7
         Left            =   -69030
         TabIndex        =   440
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   6
         Left            =   -71760
         TabIndex        =   439
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   5
         Left            =   -74490
         TabIndex        =   438
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   4
         Left            =   -66270
         TabIndex        =   437
         Text            =   "cmbOrgName"
         Top             =   990
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   3
         Left            =   -69000
         TabIndex        =   436
         Text            =   "cmbOrgName"
         Top             =   990
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   2
         Left            =   -71730
         TabIndex        =   435
         Text            =   "cmbOrgName"
         Top             =   990
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   1
         Left            =   -74460
         TabIndex        =   434
         Text            =   "cmbOrgName"
         Top             =   990
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   8
         Left            =   -66030
         TabIndex        =   425
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   7
         Left            =   -68760
         TabIndex        =   424
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   6
         Left            =   -71490
         TabIndex        =   423
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtCSComment 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   1
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   422
         Top             =   5190
         Width           =   5175
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   5
         Left            =   -74220
         TabIndex        =   421
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   4
         Left            =   -66000
         TabIndex        =   416
         Text            =   "cmbOrgGroup"
         Top             =   660
         Width           =   1935
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 6"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   6
         Left            =   -72060
         TabIndex        =   383
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   6
            Left            =   60
            TabIndex        =   384
            Top             =   3000
            Width           =   2355
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   6
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   400
            Top             =   2580
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   6
            Left            =   900
            TabIndex        =   399
            Top             =   2220
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   6
            Left            =   885
            MaxLength       =   10
            TabIndex        =   398
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   6
            Left            =   885
            MaxLength       =   10
            TabIndex        =   397
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   6
            Left            =   885
            TabIndex        =   396
            Top             =   3930
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   6
            Left            =   885
            TabIndex        =   395
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   6
            Left            =   885
            TabIndex        =   394
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   6
            Left            =   885
            TabIndex        =   393
            Top             =   5220
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   6
            Left            =   915
            TabIndex        =   392
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   6
            Left            =   915
            TabIndex        =   391
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   6
            Left            =   915
            TabIndex        =   390
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   6
            Left            =   915
            TabIndex        =   389
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   6
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   388
            Top             =   360
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   6
            Left            =   915
            TabIndex        =   387
            Top             =   690
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   6
            Left            =   885
            TabIndex        =   386
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   6
            Left            =   885
            TabIndex        =   385
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   6
            Left            =   390
            TabIndex        =   413
            Top             =   2250
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   7
            Left            =   420
            TabIndex        =   412
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   7
            Left            =   390
            TabIndex        =   411
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   6
            Left            =   405
            TabIndex        =   410
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   6
            Left            =   300
            TabIndex        =   409
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   408
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   407
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   406
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   8
            Left            =   300
            TabIndex        =   405
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   6
            Left            =   195
            TabIndex        =   404
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   6
            Left            =   450
            TabIndex        =   403
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   6
            Left            =   75
            TabIndex        =   402
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   401
            Top             =   3090
            Width           =   705
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 5"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   5
         Left            =   -74640
         TabIndex        =   352
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   5
            Left            =   60
            TabIndex        =   353
            Top             =   3000
            Width           =   2355
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   5
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   369
            Top             =   2580
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   5
            Left            =   900
            TabIndex        =   368
            Top             =   2220
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   5
            Left            =   885
            MaxLength       =   10
            TabIndex        =   367
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   5
            Left            =   885
            MaxLength       =   10
            TabIndex        =   366
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   5
            Left            =   885
            TabIndex        =   365
            Top             =   3930
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   5
            Left            =   885
            TabIndex        =   364
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   5
            Left            =   885
            TabIndex        =   363
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   5
            Left            =   885
            TabIndex        =   362
            Top             =   5220
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   5
            Left            =   915
            TabIndex        =   361
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   5
            Left            =   915
            TabIndex        =   360
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   5
            Left            =   915
            TabIndex        =   359
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   5
            Left            =   915
            TabIndex        =   358
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   5
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   357
            Top             =   360
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   5
            Left            =   915
            TabIndex        =   356
            Top             =   690
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   5
            Left            =   885
            TabIndex        =   355
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   5
            Left            =   885
            TabIndex        =   354
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   5
            Left            =   390
            TabIndex        =   382
            Top             =   2250
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   6
            Left            =   420
            TabIndex        =   381
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   6
            Left            =   390
            TabIndex        =   380
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   5
            Left            =   405
            TabIndex        =   379
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   378
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   377
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   376
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   8
            Left            =   255
            TabIndex        =   375
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   7
            Left            =   300
            TabIndex        =   374
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   5
            Left            =   195
            TabIndex        =   373
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   372
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   5
            Left            =   75
            TabIndex        =   371
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   370
            Top             =   3090
            Width           =   705
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 8"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   8
         Left            =   -66870
         TabIndex        =   321
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   8
            Left            =   60
            TabIndex        =   322
            Top             =   3030
            Width           =   2355
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   8
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   338
            Top             =   2580
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   8
            Left            =   900
            TabIndex        =   337
            Top             =   2220
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   8
            Left            =   885
            MaxLength       =   10
            TabIndex        =   336
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   8
            Left            =   885
            MaxLength       =   10
            TabIndex        =   335
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   8
            Left            =   885
            TabIndex        =   334
            Top             =   3930
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   8
            Left            =   885
            TabIndex        =   333
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   8
            Left            =   885
            TabIndex        =   332
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   8
            Left            =   885
            TabIndex        =   331
            Top             =   5220
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   8
            Left            =   915
            TabIndex        =   330
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   8
            Left            =   915
            TabIndex        =   329
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   8
            Left            =   915
            TabIndex        =   328
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   8
            Left            =   915
            TabIndex        =   327
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   8
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   326
            Top             =   360
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   8
            Left            =   915
            TabIndex        =   325
            Top             =   690
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   8
            Left            =   885
            TabIndex        =   324
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   8
            Left            =   885
            TabIndex        =   323
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   351
            Top             =   2250
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   5
            Left            =   420
            TabIndex        =   350
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   5
            Left            =   390
            TabIndex        =   349
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   4
            Left            =   405
            TabIndex        =   348
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   4
            Left            =   300
            TabIndex        =   347
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   6
            Left            =   90
            TabIndex        =   346
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   345
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   344
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   6
            Left            =   300
            TabIndex        =   343
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   342
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   4
            Left            =   450
            TabIndex        =   341
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   340
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   339
            Top             =   3090
            Width           =   705
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 3"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   3
         Left            =   -69510
         TabIndex        =   290
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   3
            Left            =   60
            TabIndex        =   291
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   3
            Left            =   885
            TabIndex        =   307
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   3
            Left            =   885
            TabIndex        =   306
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   3
            Left            =   915
            TabIndex        =   305
            Top             =   690
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   3
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   304
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   3
            Left            =   915
            TabIndex        =   303
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   3
            Left            =   915
            TabIndex        =   302
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   3
            Left            =   915
            TabIndex        =   301
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   3
            Left            =   915
            TabIndex        =   300
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   3
            Left            =   885
            TabIndex        =   299
            Top             =   5220
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   3
            Left            =   885
            TabIndex        =   298
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   3
            Left            =   885
            TabIndex        =   297
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   3
            Left            =   885
            TabIndex        =   296
            Top             =   3930
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   3
            Left            =   885
            MaxLength       =   10
            TabIndex        =   295
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   3
            Left            =   885
            MaxLength       =   10
            TabIndex        =   294
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   293
            Top             =   2220
            Width           =   1515
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   3
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   292
            Top             =   2580
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   320
            Top             =   3090
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   319
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   3
            Left            =   450
            TabIndex        =   318
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   317
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   316
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   315
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   314
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   313
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   312
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   3
            Left            =   405
            TabIndex        =   311
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   310
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   4
            Left            =   420
            TabIndex        =   309
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   3
            Left            =   390
            TabIndex        =   308
            Top             =   2250
            Width           =   465
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 2"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   2
         Left            =   -72090
         TabIndex        =   259
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   2
            Left            =   60
            TabIndex        =   260
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   2
            Left            =   885
            TabIndex        =   276
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   2
            Left            =   885
            TabIndex        =   275
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   2
            Left            =   915
            TabIndex        =   274
            Top             =   690
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   2
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   273
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   2
            Left            =   915
            TabIndex        =   272
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   2
            Left            =   915
            TabIndex        =   271
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   2
            Left            =   915
            TabIndex        =   270
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   2
            Left            =   915
            TabIndex        =   269
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   2
            Left            =   885
            TabIndex        =   268
            Top             =   5220
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   2
            Left            =   885
            TabIndex        =   267
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   2
            Left            =   885
            TabIndex        =   266
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   2
            Left            =   885
            TabIndex        =   265
            Top             =   3930
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   2
            Left            =   885
            MaxLength       =   10
            TabIndex        =   264
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   2
            Left            =   885
            MaxLength       =   10
            TabIndex        =   263
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   262
            Top             =   2220
            Width           =   1515
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   2
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   261
            Top             =   2580
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   289
            Top             =   3090
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   288
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   2
            Left            =   450
            TabIndex        =   287
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   286
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   4
            Left            =   300
            TabIndex        =   285
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   284
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   283
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   282
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   281
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   2
            Left            =   405
            TabIndex        =   280
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   3
            Left            =   390
            TabIndex        =   279
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   3
            Left            =   420
            TabIndex        =   278
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   2
            Left            =   390
            TabIndex        =   277
            Top             =   2250
            Width           =   465
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 1"
         ForeColor       =   &H00C000C0&
         Height          =   5655
         Index           =   1
         Left            =   -74670
         TabIndex        =   228
         Top             =   540
         Width           =   2505
         Begin VB.TextBox txtNotes 
            Height          =   2535
            Index           =   1
            Left            =   60
            TabIndex        =   229
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox txtExtraSensitivity 
            Height          =   285
            Index           =   1
            Left            =   885
            TabIndex        =   245
            Tag             =   "Ext"
            Top             =   3360
            Width           =   1515
         End
         Begin VB.TextBox txtUrineSensitivity 
            Height          =   285
            Index           =   1
            Left            =   885
            TabIndex        =   244
            Tag             =   "Uri"
            Top             =   3060
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   1
            Left            =   915
            TabIndex        =   243
            Top             =   690
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   1
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   242
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   1
            Left            =   915
            TabIndex        =   241
            Tag             =   "Rei"
            Top             =   1920
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   1
            Left            =   915
            TabIndex        =   240
            Tag             =   "Oxi"
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   1
            Left            =   915
            TabIndex        =   239
            Tag             =   "Cat"
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   1
            Left            =   915
            TabIndex        =   238
            Tag             =   "Coa"
            Top             =   1020
            Width           =   1515
         End
         Begin VB.ComboBox cmbChromogenic 
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   237
            Top             =   5220
            Width           =   1515
         End
         Begin VB.ComboBox cmbRapidec 
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   236
            Top             =   4890
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI2 
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   235
            Top             =   4560
            Width           =   1515
         End
         Begin VB.ComboBox cmbAPI1 
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   234
            Top             =   3930
            Width           =   1515
         End
         Begin VB.TextBox txtAPI2 
            Height          =   285
            Index           =   1
            Left            =   885
            MaxLength       =   10
            TabIndex        =   233
            Tag             =   "AP2"
            Top             =   4260
            Width           =   1515
         End
         Begin VB.TextBox txtAPI1 
            Height          =   285
            Index           =   1
            Left            =   885
            MaxLength       =   10
            TabIndex        =   232
            Tag             =   "AP1"
            Top             =   3660
            Width           =   1515
         End
         Begin VB.TextBox txtCrystal 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   231
            Top             =   2220
            Width           =   1515
         End
         Begin VB.CommandButton cmdNotes 
            Caption         =   "Notes"
            Height          =   315
            Index           =   1
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   230
            Top             =   2580
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sensitivity"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   258
            Top             =   3090
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extra Sens"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   257
            Top             =   3390
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   256
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   255
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   254
            Top             =   1650
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   4
            Left            =   255
            TabIndex        =   253
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   252
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   251
            Top             =   420
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Chromo"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   250
            Top             =   5280
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Staph"
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   249
            Top             =   4950
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "API (ii)"
            Height          =   195
            Index           =   2
            Left            =   390
            TabIndex        =   248
            Top             =   4320
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "API (i)"
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   247
            Top             =   3690
            Width           =   420
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Crystal"
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   246
            Top             =   2250
            Width           =   465
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Clinical Details"
         Height          =   1815
         Left            =   5580
         TabIndex        =   219
         Top             =   4500
         Width           =   4155
         Begin VB.TextBox txtClinDetails 
            Height          =   1095
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   221
            Top             =   600
            Width           =   3825
         End
         Begin VB.ComboBox cmbClinDetails 
            Height          =   315
            Left            =   150
            Sorted          =   -1  'True
            TabIndex        =   220
            Top             =   270
            Width           =   3825
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Patients Current Antibiotics"
         Height          =   1035
         Left            =   5700
         TabIndex        =   214
         Top             =   3300
         Width           =   4035
         Begin VB.CommandButton cmdABsInUse 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   217
            Top             =   420
            Width           =   375
         End
         Begin VB.ListBox lstABsInUse 
            Height          =   735
            IntegralHeight  =   0   'False
            ItemData        =   "frmOldEditBacteriology.frx":8E36
            Left            =   120
            List            =   "frmOldEditBacteriology.frx":8E38
            TabIndex        =   216
            ToolTipText     =   "Click to remove entry"
            Top             =   240
            Width           =   3345
         End
         Begin VB.ComboBox cmbABsInUse 
            Height          =   315
            Left            =   120
            TabIndex        =   215
            Text            =   "cmbABsInUse"
            Top             =   420
            Visible         =   0   'False
            Width           =   3345
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Site"
         Height          =   765
         Left            =   5700
         TabIndex        =   210
         Top             =   2490
         Width           =   5415
         Begin VB.TextBox txtSiteDetails 
            Height          =   315
            Left            =   2130
            TabIndex        =   212
            Top             =   270
            Width           =   3195
         End
         Begin VB.ComboBox cmbSite 
            Height          =   315
            Left            =   120
            TabIndex        =   211
            Text            =   "cmbSite"
            Top             =   270
            Width           =   1965
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Site Details"
            Height          =   195
            Left            =   2190
            TabIndex        =   213
            Top             =   30
            Width           =   795
         End
      End
      Begin VB.TextBox txtCSComment 
         BackColor       =   &H80000018&
         Height          =   855
         Index           =   0
         Left            =   -74490
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   207
         Top             =   5370
         Width           =   5175
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   1
         Left            =   -74190
         TabIndex        =   206
         Text            =   "cmbOrgGroup"
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   2
         Left            =   -71460
         TabIndex        =   205
         Text            =   "cmbOrgGroup"
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Index           =   3
         Left            =   -68730
         TabIndex        =   204
         Text            =   "cmbOrgGroup"
         Top             =   660
         Width           =   1935
      End
      Begin VB.Frame frAPI 
         Caption         =   "API"
         Height          =   2175
         Left            =   -70590
         TabIndex        =   188
         Top             =   630
         Width           =   3585
         Begin VB.CheckBox chkAPI 
            Caption         =   "Check5"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   203
            Top             =   1470
            Width           =   195
         End
         Begin VB.CheckBox chkAPI 
            Caption         =   "Check4"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   202
            Top             =   1170
            Width           =   195
         End
         Begin VB.CheckBox chkAPI 
            Caption         =   "Check3"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   201
            Top             =   870
            Width           =   195
         End
         Begin VB.CheckBox chkAPI 
            Caption         =   "Check2"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   200
            Top             =   570
            Width           =   195
         End
         Begin VB.CheckBox chkAPI 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   199
            Top             =   1770
            Width           =   195
         End
         Begin VB.TextBox txtAPICode 
            Height          =   285
            Index           =   0
            Left            =   390
            MaxLength       =   7
            TabIndex        =   198
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtAPICode 
            Height          =   285
            Index           =   1
            Left            =   390
            MaxLength       =   7
            TabIndex        =   197
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtAPICode 
            Height          =   285
            Index           =   2
            Left            =   390
            MaxLength       =   7
            TabIndex        =   196
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox txtAPICode 
            Height          =   285
            Index           =   3
            Left            =   390
            MaxLength       =   7
            TabIndex        =   195
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtAPIName 
            Height          =   285
            Index           =   0
            Left            =   1230
            TabIndex        =   194
            Top             =   540
            Width           =   2205
         End
         Begin VB.TextBox txtAPIName 
            Height          =   285
            Index           =   1
            Left            =   1230
            TabIndex        =   193
            Top             =   840
            Width           =   2205
         End
         Begin VB.TextBox txtAPIName 
            Height          =   285
            Index           =   2
            Left            =   1230
            TabIndex        =   192
            Top             =   1140
            Width           =   2205
         End
         Begin VB.TextBox txtAPIName 
            Height          =   285
            Index           =   3
            Left            =   1230
            TabIndex        =   191
            Top             =   1440
            Width           =   2205
         End
         Begin VB.TextBox txtAPICode 
            Height          =   285
            Index           =   4
            Left            =   390
            MaxLength       =   7
            TabIndex        =   190
            Top             =   1740
            Width           =   855
         End
         Begin VB.TextBox txtAPIName 
            Height          =   285
            Index           =   4
            Left            =   1230
            TabIndex        =   189
            Top             =   1740
            Width           =   2205
         End
      End
      Begin VB.Frame frCampylobacter 
         Caption         =   "Campylobacter"
         Height          =   2175
         Left            =   -66960
         TabIndex        =   178
         Top             =   630
         Width           =   2685
         Begin VB.CommandButton bSens 
            Caption         =   "Sensitivity"
            Height          =   285
            Index           =   2
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   210
            Width           =   915
         End
         Begin VB.TextBox txtGram 
            Height          =   285
            Left            =   630
            MaxLength       =   40
            TabIndex        =   180
            Top             =   1470
            Width           =   1905
         End
         Begin VB.TextBox txtCampCulture 
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   630
            TabIndex        =   179
            Top             =   1740
            Width           =   1905
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "PC"
            Height          =   195
            Index           =   2
            Left            =   1020
            TabIndex        =   187
            Top             =   720
            Width           =   210
         End
         Begin VB.Label lblCamp 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1290
            TabIndex        =   186
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Latex Screen"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   185
            Top             =   1110
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Gram"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   184
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Culture"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   183
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label lblCampLatex 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1290
            TabIndex        =   182
            Top             =   1080
            Width           =   1245
         End
      End
      Begin VB.Frame FToxinA 
         Caption         =   "Toxin A"
         Height          =   615
         Left            =   -70710
         TabIndex        =   173
         Top             =   5430
         Width           =   3255
         Begin VB.Label lblToxinAL 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   177
            Top             =   210
            Width           =   855
         End
         Begin VB.Label lblToxinATA 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2130
            TabIndex        =   176
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Latex"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   175
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Toxin A"
            Height          =   195
            Index           =   2
            Left            =   1500
            TabIndex        =   174
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame fCulture 
         Caption         =   "Culture"
         Height          =   2175
         Left            =   -74580
         TabIndex        =   168
         Top             =   630
         Width           =   1245
         Begin VB.CheckBox chkSeleniteDone 
            Caption         =   "Selenite"
            Height          =   195
            Left            =   60
            TabIndex        =   170
            Top             =   1530
            Width           =   1035
         End
         Begin VB.CheckBox chkPCDone 
            Caption         =   "PC Done"
            Height          =   225
            Left            =   60
            TabIndex        =   169
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label lblSelenite 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   30
            TabIndex        =   172
            Top             =   1740
            Width           =   1155
         End
         Begin VB.Label lblPC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   30
            TabIndex        =   171
            Top             =   720
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Screen"
         Height          =   2175
         Index           =   1
         Left            =   -73290
         TabIndex        =   150
         Top             =   630
         Width           =   1875
         Begin VB.CheckBox chkScreen 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   155
            Top             =   630
            Width           =   225
         End
         Begin VB.CheckBox chkScreen 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   154
            Top             =   900
            Width           =   225
         End
         Begin VB.CheckBox chkScreen 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   153
            Top             =   1200
            Width           =   225
         End
         Begin VB.CheckBox chkScreen 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   152
            Top             =   1470
            Width           =   225
         End
         Begin VB.CheckBox chkScreen 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   151
            Top             =   1770
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lact"
            Height          =   195
            Index           =   2
            Left            =   570
            TabIndex        =   167
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Urea"
            Height          =   195
            Index           =   1
            Left            =   1230
            TabIndex        =   166
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lblLact 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   165
            Top             =   600
            Width           =   555
         End
         Begin VB.Label lblLact 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   164
            Top             =   900
            Width           =   555
         End
         Begin VB.Label lblLact 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   163
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label lblLact 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   162
            Top             =   1470
            Width           =   555
         End
         Begin VB.Label lblLact 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   161
            Top             =   1770
            Width           =   555
         End
         Begin VB.Label lblUrea 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   1170
            TabIndex        =   160
            Top             =   600
            Width           =   555
         End
         Begin VB.Label lblUrea 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   1170
            TabIndex        =   159
            Top             =   900
            Width           =   555
         End
         Begin VB.Label lblUrea 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   1170
            TabIndex        =   158
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label lblUrea 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   1170
            TabIndex        =   157
            Top             =   1500
            Width           =   555
         End
         Begin VB.Label lblUrea 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   1170
            TabIndex        =   156
            Top             =   1800
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Purity"
         Height          =   2175
         Index           =   1
         Left            =   -71370
         TabIndex        =   144
         Top             =   630
         Width           =   705
         Begin VB.CheckBox chkPurity 
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   149
            Top             =   570
            Width           =   225
         End
         Begin VB.CheckBox chkPurity 
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   148
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox chkPurity 
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   147
            Top             =   1170
            Width           =   225
         End
         Begin VB.CheckBox chkPurity 
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   146
            Top             =   1470
            Width           =   225
         End
         Begin VB.CheckBox chkPurity 
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   145
            Top             =   1770
            Width           =   225
         End
      End
      Begin VB.Frame frEPC 
         Caption         =   "EPC"
         Height          =   1785
         Left            =   -72210
         TabIndex        =   135
         Top             =   4260
         Width           =   1425
         Begin VB.Label lblEPC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   750
            TabIndex        =   143
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label lblEPC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   750
            TabIndex        =   142
            Top             =   1000
            Width           =   495
         End
         Begin VB.Label lblEPC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   750
            TabIndex        =   141
            Top             =   650
            Width           =   495
         End
         Begin VB.Label lblEPC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   750
            TabIndex        =   140
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "018c:K"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   139
            Top             =   1350
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Poly 4"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   138
            Top             =   990
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Poly 3"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   137
            Top             =   660
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Poly 2"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   136
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame frOccult 
         Caption         =   "Occult Blood"
         Height          =   1095
         Left            =   -70710
         TabIndex        =   128
         Top             =   4290
         Width           =   1365
         Begin VB.CheckBox chkOccult 
            Alignment       =   1  'Right Justify
            Caption         =   "3"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   131
            Top             =   780
            Width           =   405
         End
         Begin VB.CheckBox chkOccult 
            Alignment       =   1  'Right Justify
            Caption         =   "2"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   130
            Top             =   510
            Width           =   405
         End
         Begin VB.CheckBox chkOccult 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblOccult 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   134
            Top             =   780
            Width           =   615
         End
         Begin VB.Label lblOccult 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   133
            Top             =   510
            Width           =   615
         End
         Begin VB.Label lblOccult 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   132
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frRotaAdeno 
         Caption         =   "Rota/Adeno"
         Height          =   1095
         Left            =   -69270
         TabIndex        =   123
         Top             =   4290
         Width           =   1815
         Begin VB.Label lblRota 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   570
            TabIndex        =   127
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblAdeno 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   570
            TabIndex        =   126
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Adeno"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   125
            Top             =   630
            Width           =   465
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Rota"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   124
            Top             =   300
            Width           =   345
         End
      End
      Begin VB.Frame FrOva 
         Caption         =   "Ova"
         Height          =   1755
         Left            =   -67410
         TabIndex        =   116
         Top             =   4290
         Width           =   3645
         Begin VB.ComboBox cmbOP 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   119
            Text            =   "cmbOP"
            Top             =   1230
            Width           =   3195
         End
         Begin VB.ComboBox cmbOP 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   118
            Text            =   "cmbOP"
            Top             =   870
            Width           =   3195
         End
         Begin VB.ComboBox cmbOP 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   117
            Text            =   "cmbOP"
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblAus 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2100
            TabIndex        =   122
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Auramine Stain"
            Height          =   195
            Index           =   2
            Left            =   780
            TabIndex        =   121
            Top             =   210
            Width           =   1290
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Wet"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   120
            Top             =   540
            Width           =   360
         End
      End
      Begin VB.Frame F0157 
         Caption         =   "E Coli 0157"
         Height          =   1785
         Left            =   -74580
         TabIndex        =   110
         Top             =   4260
         Width           =   2295
         Begin VB.Label lbl0157Latex 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1200
            TabIndex        =   115
            Top             =   540
            Width           =   915
         End
         Begin VB.Label lblPC0157 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1200
            TabIndex        =   114
            Top             =   270
            Width           =   915
         End
         Begin VB.Label lbl0157 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   210
            TabIndex        =   113
            Top             =   1290
            Width           =   1905
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Latex Screen"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   112
            Top             =   540
            Width           =   945
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "PC"
            Height          =   195
            Index           =   2
            Left            =   870
            TabIndex        =   111
            Top             =   270
            Width           =   210
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Salmonella/Shigella"
         Height          =   1125
         Left            =   -74580
         TabIndex        =   103
         Top             =   2880
         Width           =   7575
         Begin VB.Label lblColindale 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   109
            Top             =   660
            Width           =   6315
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Colindale"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   108
            Top             =   690
            Width           =   645
         End
         Begin VB.Label lblShigella 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4590
            TabIndex        =   107
            Top             =   270
            Width           =   2790
         End
         Begin VB.Label lblSalmonella 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   106
            Top             =   270
            Width           =   2805
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Shigella"
            Height          =   195
            Left            =   3990
            TabIndex        =   105
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Salmonella"
            Height          =   195
            Left            =   210
            TabIndex        =   104
            Top             =   300
            Width           =   765
         End
      End
      Begin VB.Frame fraUrineSpecific 
         Caption         =   "Specific"
         Height          =   1785
         Left            =   -70830
         TabIndex        =   99
         Top             =   2190
         Width           =   3855
         Begin VB.TextBox txtFatGlobules 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   94
            Top             =   1050
            Width           =   2055
         End
         Begin VB.TextBox txtBenceJones 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   92
            Top             =   450
            Width           =   2055
         End
         Begin VB.TextBox txtSG 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   93
            Top             =   750
            Width           =   2055
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Fat Globules"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   102
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Bence Jones Protein"
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   101
            Top             =   480
            Width           =   1485
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Specific Gravity"
            Height          =   195
            Left            =   390
            TabIndex        =   100
            Top             =   780
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Microscopy"
         Height          =   3165
         Left            =   -74010
         TabIndex        =   90
         Top             =   780
         Width           =   2955
         Begin VB.CommandButton cmdNADMicro 
            Caption         =   "NAD"
            Height          =   345
            Left            =   180
            TabIndex        =   564
            Top             =   390
            Width           =   615
         End
         Begin VB.TextBox txtBacteria 
            Height          =   285
            Left            =   1560
            TabIndex        =   75
            Top             =   420
            Width           =   1185
         End
         Begin VB.ComboBox cmbMisc 
            Height          =   315
            Index           =   2
            Left            =   750
            TabIndex        =   82
            Text            =   "cmbMisc"
            Top             =   2700
            Width           =   2025
         End
         Begin VB.ComboBox cmbMisc 
            Height          =   315
            Index           =   1
            Left            =   750
            TabIndex        =   81
            Text            =   "cmbMisc"
            Top             =   2370
            Width           =   2025
         End
         Begin VB.ComboBox cmbMisc 
            Height          =   315
            Index           =   0
            Left            =   750
            TabIndex        =   80
            Text            =   "cmbMisc"
            Top             =   2040
            Width           =   2025
         End
         Begin VB.ComboBox cmbCrystals 
            Height          =   315
            Left            =   750
            TabIndex        =   78
            Text            =   "cmbCrystals"
            Top             =   1380
            Width           =   2025
         End
         Begin VB.ComboBox cmbCasts 
            Height          =   315
            Left            =   750
            TabIndex        =   79
            Text            =   "cmbCasts"
            Top             =   1710
            Width           =   2025
         End
         Begin VB.TextBox txtWCC 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   76
            Top             =   780
            Width           =   780
         End
         Begin VB.TextBox txtRCC 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   77
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "/cmm"
            Height          =   195
            Left            =   2340
            TabIndex        =   488
            Top             =   810
            Width           =   435
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Bacteria"
            Height          =   195
            Left            =   900
            TabIndex        =   227
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "WCC"
            Height          =   195
            Index           =   1
            Left            =   1110
            TabIndex        =   98
            Top             =   810
            Width           =   375
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "RCC"
            Height          =   195
            Index           =   0
            Left            =   1170
            TabIndex        =   97
            Top             =   1110
            Width           =   330
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Misc"
            Height          =   195
            Left            =   390
            TabIndex        =   96
            Top             =   2100
            Width           =   330
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Casts"
            Height          =   195
            Left            =   330
            TabIndex        =   95
            Top             =   1770
            Width           =   390
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Crystals"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   91
            Top             =   1440
            Width           =   540
         End
      End
      Begin VB.Frame frDipStick 
         Caption         =   "Dip Stick"
         Height          =   3165
         Left            =   -66690
         TabIndex        =   66
         Top             =   780
         Width           =   2415
         Begin VB.CommandButton cmdNAD 
            Appearance      =   0  'Flat
            Caption         =   "NAD"
            Height          =   285
            Left            =   1050
            TabIndex        =   69
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox txtBloodHb 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   67
            Top             =   2700
            Width           =   1200
         End
         Begin VB.TextBox txtpH 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   70
            Top             =   900
            Width           =   1200
         End
         Begin VB.TextBox txtProtein 
            Height          =   285
            Left            =   1050
            MaxLength       =   20
            TabIndex        =   71
            Top             =   1200
            Width           =   1200
         End
         Begin VB.TextBox txtGlucose 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   72
            Top             =   1500
            Width           =   1200
         End
         Begin VB.TextBox txtKetones 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   73
            Top             =   1800
            Width           =   1200
         End
         Begin VB.TextBox txtBilirubin 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   68
            Top             =   2400
            Width           =   1200
         End
         Begin VB.TextBox txtUrobilinogen 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   74
            Top             =   2100
            Width           =   1200
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Blood/Hb"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   89
            Top             =   2730
            Width           =   690
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "pH"
            Height          =   195
            Index           =   0
            Left            =   780
            TabIndex        =   88
            Top             =   960
            Width           =   210
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Protein"
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   87
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Glucose"
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   86
            Top             =   1530
            Width           =   585
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Ketones"
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   85
            Top             =   1830
            Width           =   585
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Bilirubin"
            Height          =   195
            Index           =   1
            Left            =   450
            TabIndex        =   84
            Top             =   2430
            Width           =   540
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Urobilinogen"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   2130
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdOrderTests 
         Caption         =   "Order Tests"
         Height          =   945
         Left            =   10050
         Picture         =   "frmOldEditBacteriology.frx":8E3A
         Style           =   1  'Graphical
         TabIndex        =   62
         Tag             =   "bOrder"
         Top             =   3390
         Width           =   1035
      End
      Begin VB.CommandButton cmdSaveInc 
         Caption         =   "&Save"
         Height          =   795
         Left            =   10050
         Picture         =   "frmOldEditBacteriology.frx":9144
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5490
         Width           =   1035
      End
      Begin VB.TextBox txtUrineComment 
         Height          =   1185
         Left            =   -74040
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   4590
         Width           =   7065
      End
      Begin VB.Frame Frame4 
         Height          =   5655
         Left            =   330
         TabIndex        =   39
         Top             =   660
         Width           =   5265
         Begin VB.CheckBox chkPenicillin 
            Alignment       =   1  'Right Justify
            Caption         =   "Penicillin Allergy"
            Height          =   225
            Left            =   3540
            TabIndex        =   573
            Top             =   540
            Width           =   1425
         End
         Begin VB.ComboBox cmbHospital 
            Height          =   315
            Left            =   1050
            TabIndex        =   8
            Text            =   "cmbHospital"
            Top             =   2160
            Width           =   3915
         End
         Begin VB.ComboBox cmbDemogComment 
            Height          =   315
            Left            =   1050
            TabIndex        =   209
            Top             =   3570
            Width           =   3915
         End
         Begin VB.ComboBox cmbGP 
            Height          =   315
            Left            =   1050
            TabIndex        =   11
            Text            =   "cmbGP"
            Top             =   3210
            Width           =   3915
         End
         Begin VB.ComboBox cmbClinician 
            Height          =   315
            Left            =   1050
            TabIndex        =   10
            Text            =   "cmbClinician"
            Top             =   2850
            Width           =   3915
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   1
            Left            =   750
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1770
            Width           =   4215
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Index           =   0
            Left            =   750
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1500
            Width           =   4215
         End
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   1050
            TabIndex        =   9
            Text            =   "cmbWard"
            Top             =   2490
            Width           =   3915
         End
         Begin VB.TextBox txtDemographicComment 
            Height          =   1515
            Left            =   1050
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   3900
            Width           =   3885
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Hospital"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   226
            Top             =   2220
            Width           =   570
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "GP"
            Height          =   195
            Index           =   0
            Left            =   780
            TabIndex        =   54
            Top             =   3270
            Width           =   225
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Clinician"
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   53
            Top             =   2880
            Width           =   585
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1530
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Ward"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   50
            Top             =   2550
            Width           =   390
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            Height          =   195
            Index           =   0
            Left            =   3930
            TabIndex        =   49
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            Height          =   195
            Index           =   0
            Left            =   2760
            TabIndex        =   48
            Top             =   1200
            Width           =   285
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "D.o.B"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   47
            Top             =   1230
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   46
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Chart #"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   45
            Top             =   330
            Width           =   525
         End
         Begin VB.Label lblChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   44
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lblName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   750
            TabIndex        =   43
            Top             =   780
            Width           =   4215
         End
         Begin VB.Label lblDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   42
            Top             =   1170
            Width           =   1515
         End
         Begin VB.Label lblAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3180
            TabIndex        =   41
            Top             =   1170
            Width           =   585
         End
         Begin VB.Label lblSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4260
            TabIndex        =   40
            Top             =   1170
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdSaveDemographics 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   825
         Left            =   10050
         Picture         =   "frmOldEditBacteriology.frx":97AE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4560
         Width           =   1035
      End
      Begin VB.Frame Frame7 
         Caption         =   "Run Date"
         Height          =   1725
         Left            =   5700
         TabIndex        =   21
         Top             =   660
         Width           =   4095
         Begin MSComCtl2.DTPicker dtRecDate 
            Height          =   315
            Left            =   1890
            TabIndex        =   575
            Top             =   1050
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   38078
         End
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   150
            TabIndex        =   22
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   1890
            TabIndex        =   23
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   3270
            TabIndex        =   24
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
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   2760
            Picture         =   "frmOldEditBacteriology.frx":9BF0
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   1890
            Picture         =   "frmOldEditBacteriology.frx":A032
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   2370
            Picture         =   "frmOldEditBacteriology.frx":A474
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1380
            Width           =   360
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Received in Lab"
            Height          =   195
            Left            =   690
            TabIndex        =   574
            Top             =   1110
            Width           =   1170
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   2370
            Picture         =   "frmOldEditBacteriology.frx":A8B6
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   630
            Picture         =   "frmOldEditBacteriology.frx":ACF8
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   2760
            Picture         =   "frmOldEditBacteriology.frx":B13A
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   1890
            Picture         =   "frmOldEditBacteriology.frx":B57C
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   1020
            Picture         =   "frmOldEditBacteriology.frx":B9BE
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   120
            Picture         =   "frmOldEditBacteriology.frx":BE00
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sample Date"
            Height          =   195
            Index           =   2
            Left            =   1920
            TabIndex        =   25
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Height          =   795
         Left            =   9750
         TabIndex        =   26
         Top             =   660
         Width           =   1365
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   28
            Top             =   240
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3165
         Index           =   5
         Left            =   -74520
         TabIndex        =   426
         Top             =   1470
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   6
         Left            =   -71790
         TabIndex        =   464
         Top             =   1470
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   7
         Left            =   -69060
         TabIndex        =   465
         Top             =   1470
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   8
         Left            =   -66330
         TabIndex        =   466
         Top             =   1470
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   4
         Left            =   -66300
         TabIndex        =   467
         Top             =   1650
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   3
         Left            =   -69030
         TabIndex        =   468
         Top             =   1650
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   2
         Left            =   -71760
         TabIndex        =   469
         Top             =   1650
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
         Height          =   3165
         Index           =   1
         Left            =   -74490
         TabIndex        =   470
         Top             =   1650
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
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
      Begin VB.Image imgMoreIdentity 
         Height          =   480
         Left            =   -64410
         Picture         =   "frmOldEditBacteriology.frx":C242
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblMoreIdentity 
         AutoSize        =   -1  'True
         Caption         =   "More"
         Height          =   195
         Left            =   -63900
         TabIndex        =   613
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Consultant Comments"
         Height          =   195
         Index           =   1
         Left            =   -69060
         TabIndex        =   501
         Top             =   5010
         Width           =   1530
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Consultant Comments"
         Height          =   195
         Index           =   0
         Left            =   -69000
         TabIndex        =   500
         Top             =   5190
         Width           =   1530
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -66600
         TabIndex        =   455
         ToolTipText     =   "Set All Resistant"
         Top             =   3810
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -66600
         TabIndex        =   454
         ToolTipText     =   "Set All Sensitive"
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -69330
         TabIndex        =   453
         ToolTipText     =   "Set All Resistant"
         Top             =   3810
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -69330
         TabIndex        =   452
         ToolTipText     =   "Set All Sensitive"
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -72060
         TabIndex        =   451
         ToolTipText     =   "Set All Resistant"
         Top             =   3810
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -72060
         TabIndex        =   450
         ToolTipText     =   "Set All Sensitive"
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -74790
         TabIndex        =   449
         ToolTipText     =   "Set All Resistant"
         Top             =   3810
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -74790
         TabIndex        =   448
         ToolTipText     =   "Set All Sensitive"
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -66570
         TabIndex        =   447
         ToolTipText     =   "Set All Resistant"
         Top             =   3990
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -66570
         TabIndex        =   446
         ToolTipText     =   "Set All Sensitive"
         Top             =   4380
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -69300
         TabIndex        =   445
         ToolTipText     =   "Set All Resistant"
         Top             =   3990
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -69300
         TabIndex        =   444
         ToolTipText     =   "Set All Sensitive"
         Top             =   4380
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -72030
         TabIndex        =   443
         ToolTipText     =   "Set All Resistant"
         Top             =   3990
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -72030
         TabIndex        =   442
         ToolTipText     =   "Set All Sensitive"
         Top             =   4380
         Width           =   255
      End
      Begin VB.Label lblMoreID 
         AutoSize        =   -1  'True
         Caption         =   "More"
         Height          =   195
         Left            =   -64020
         TabIndex        =   433
         Top             =   870
         Width           =   360
      End
      Begin VB.Image imgMoreID 
         Height          =   480
         Left            =   -64410
         Picture         =   "frmOldEditBacteriology.frx":C684
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblMoreCS 
         AutoSize        =   -1  'True
         Caption         =   "More"
         Height          =   195
         Left            =   -63900
         TabIndex        =   432
         Top             =   1050
         Width           =   360
      End
      Begin VB.Image imgMoreCS 
         Height          =   480
         Left            =   -63960
         Picture         =   "frmOldEditBacteriology.frx":CAC6
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
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
         Index           =   7
         Left            =   -66300
         TabIndex        =   431
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
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
         Index           =   6
         Left            =   -69030
         TabIndex        =   430
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
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
         Index           =   5
         Left            =   -71760
         TabIndex        =   429
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label20 
         Caption         =   "Specimen Comments"
         Height          =   225
         Index           =   3
         Left            =   -74490
         TabIndex        =   428
         Top             =   4980
         Width           =   1515
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
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
         Index           =   4
         Left            =   -74490
         TabIndex        =   427
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label46 
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
         Index           =   3
         Left            =   -66270
         TabIndex        =   420
         Top             =   660
         Width           =   270
      End
      Begin VB.Label Label46 
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
         Index           =   2
         Left            =   -69000
         TabIndex        =   419
         Top             =   660
         Width           =   270
      End
      Begin VB.Label Label46 
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
         Index           =   1
         Left            =   -71730
         TabIndex        =   418
         Top             =   660
         Width           =   270
      End
      Begin VB.Label Label46 
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
         Index           =   0
         Left            =   -74460
         TabIndex        =   417
         Top             =   660
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -74760
         TabIndex        =   225
         ToolTipText     =   "Set All Sensitive"
         Top             =   4380
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -74760
         TabIndex        =   224
         ToolTipText     =   "Set All Resistant"
         Top             =   3990
         Width           =   270
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   -63870
         Picture         =   "frmOldEditBacteriology.frx":CF08
         Top             =   390
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareCross 
         Height          =   225
         Left            =   -63660
         Picture         =   "frmOldEditBacteriology.frx":D1DE
         Top             =   390
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Medical Scientist Comments"
         Height          =   195
         Index           =   0
         Left            =   -74460
         TabIndex        =   208
         Top             =   5160
         Width           =   1980
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Urine Specimen Comment"
         Height          =   195
         Left            =   -74040
         TabIndex        =   56
         Top             =   4410
         Width           =   1830
      End
   End
   Begin VB.Label lblNOPAS 
      AutoSize        =   -1  'True
      Caption         =   "NOPAS"
      Height          =   195
      Left            =   14250
      TabIndex        =   579
      Top             =   1050
      Width           =   555
   End
   Begin VB.Label lblAandE 
      Caption         =   "A and E"
      Height          =   225
      Left            =   14100
      TabIndex        =   578
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmOldEditMicrobiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNewRecord As Boolean

Private Activated As Boolean

Private pPrintToPrinter As String

Private UrineLoaded As Boolean
Private IdentLoaded As Boolean
Private FaecesLoaded As Boolean
Private CSLoaded As Boolean
Private FOBLoaded As Boolean
Private RotaAdenoLoaded As Boolean
Private CdiffLoaded As Boolean
Private OPLoaded As Boolean
Private IdentificationLoaded As Boolean

Private SampleIDWithOffset As Long

Private frmOptUrineSpecific As Boolean
Private Sub ClearIndividualFaeces()

Dim n As Integer

For n = 0 To 2
  chkFOB(n) = 0
  lblFOB(n) = ""
  lblFOB(n).BackColor = &H8000000F
Next

txtRota = ""
txtRota.BackColor = &H8000000F
txtAdeno = ""
txtAdeno.BackColor = &H8000000F

lblToxinA = ""
lblToxinA.BackColor = &H8000000F
lblToxinB = ""
lblToxinB.BackColor = &H8000000F

lblCrypto = ""
lblCrypto.BackColor = &H8000000F
For n = 0 To 2
  cmbOva(n) = ""
Next

End Sub

Private Sub EnableCopyFrom()

Dim sql As String
Dim tb As Recordset
Dim PrevSID As Long

cmdCopyFromPrevious.Visible = False

If sysOptAllowCopyDemographics(0) = False Then
  Exit Sub
End If

If Trim$(txtName) <> "" Or txtDoB <> "" Then
  Exit Sub
End If

PrevSID = sysOptMicroOffset(0) + Val(txtSampleID) - 1

sql = "Select * from Demographics where " & _
      "SampleID = " & PrevSID & " " & _
      "and PatName <> '' " & _
      "and PatName is not null " & _
      "and DoB is not null"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmdCopyFromPrevious.Caption = "Copy All Details from Sample # " & Format$(PrevSID - sysOptMicroOffset(0)) & " Name " & tb!PatName
  cmdCopyFromPrevious.Visible = True
End If

End Sub

Private Function CheckIfValid(ByVal Isolate As Integer) As Boolean

Dim sql As String
Dim tb As Recordset

sql = "Select count (Valid) as tot from Sensitivities where " & _
      "SampleID = '" & SampleIDWithOffset & "' " & _
      "and IsolateNumber = '" & Isolate & "' and Valid = 1"

Set tb = New Recordset
RecOpenClient 0, tb, sql

CheckIfValid = tb!tot > 0

End Function

Private Sub ClearIdent()

Dim Index As Integer

For Index = 1 To 8
  cmbGram(Index) = ""
  cmbWetPrep(Index) = ""
  txtCoagulase(Index) = ""
  txtCatalase(Index) = ""
  txtOxidase(Index) = ""
  txtAPI1(Index) = ""
  txtAPI2(Index) = ""
  cmbAPI1(Index) = ""
  cmbAPI2(Index) = ""
  cmbRapidec(Index) = ""
  cmbChromogenic(Index) = ""
  txtReincubation(Index) = ""
  txtUrineSensitivity(Index) = ""
  txtExtraSensitivity(Index) = ""
  txtNotes(Index) = ""
  cmdNotes(Index).BackColor = vbButtonFace
Next

End Sub

Private Sub ClearUrine()
  
txtBacteria = ""

txtPregnancy = ""
txtHCGLevel = ""
txtBenceJones = ""
txtSG = ""
txtFatGlobules = ""

txtPh = ""
txtProtein = ""
txtGlucose = ""
txtKetones = ""
txtUrobilinogen = ""
txtBilirubin = ""
txtBloodHb = ""
txtWCC = ""
txtRCC = ""
cmbCrystals = ""
cmbCasts = ""
cmbMisc(0) = ""
cmbMisc(1) = ""
cmbMisc(2) = ""

End Sub


Private Sub cmbSiteEffects()

Dim F As Form
Dim n As Integer

For n = 1 To 13
  ssTab1.TabVisible(n) = False
Next
ssTab1.TabVisible(4) = True 'c&S
ssTab1.TabVisible(12) = True 'Identification
ssTab1.TabVisible(13) = True 'Identification 5-8

cmdOrderTests.Enabled = False

Select Case cmbSite
  Case "Faeces":
    
    OrderFaeces
  
  Case "Urine":
    
    Set F = frmMicroUrineSite
    F.Show 1
    txtSiteDetails = F.Details
    Unload F
    Set F = Nothing

    With frmMicroOrderUrine
      .txtSampleID = txtSampleID
      .Show 1
    End With
    
    ssTab1.TabVisible(1) = True 'Urine
    ssTab1.TabVisible(2) = False
    ssTab1.TabVisible(3) = False
    ssTab1.TabVisible(4) = True
    ssTab1.TabVisible(5) = False
    ssTab1.TabVisible(8) = False
    ssTab1.TabVisible(9) = False
    ssTab1.TabVisible(10) = False
    ssTab1.TabVisible(11) = False
    ssTab1.TabVisible(12) = True
    ssTab1.TabVisible(13) = True
    
    cmdOrderTests.Enabled = True
  Case Else:
    ssTab1.TabVisible(1) = False
    ssTab1.TabVisible(2) = False
    ssTab1.TabVisible(3) = False
    ssTab1.TabVisible(4) = True
    ssTab1.TabVisible(8) = False
    ssTab1.TabVisible(9) = False
    ssTab1.TabVisible(10) = False
    ssTab1.TabVisible(11) = False
End Select

lblSiteDetails = cmbSite & " " & txtSiteDetails

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub FillABSelect(ByVal Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim ExcludeList As String


cmbABSelect(Index).Clear

ExcludeList = ""
For n = 1 To grdAB(Index).Rows - 1
  ExcludeList = ExcludeList & _
                "AntibioticName <> '" & grdAB(Index).TextMatrix(n, 0) & "' and "
Next
ExcludeList = Left$(ExcludeList, Len(ExcludeList) - 4)

sql = "Select Distinct AntibioticName, ListOrder from Antibiotics where " & _
      ExcludeList & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenClient 0, tb, sql
Do While Not tb.EOF
  cmbABSelect(Index).AddItem Trim$(tb!AntibioticName & "")
  tb.MoveNext
Loop

End Sub

Private Sub FillCurrentABs()

Dim tb As Recordset
Dim sql As String

cmbABsInUse.Clear

sql = "Select distinct AntibioticName, ListOrder " & _
      "from Antibiotics " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenClient 0, tb, sql
Do While Not tb.EOF
  cmbABsInUse.AddItem Trim$(tb!AntibioticName & "")
  tb.MoveNext
Loop

End Sub

Private Sub FillForConsultantValidation()

Dim sql As String
Dim tb As Recordset
Dim SID As Long
    
cmdAddToConsultantList.Caption = "Add to Consultant List"

cmbConsultantVal.Clear

sql = "Select * from ConsultantList " & _
      "Order by SampleID"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  SID = Val(tb!SampleID) - sysOptMicroOffset(0)
  cmbConsultantVal.AddItem Format$(SID)
  If SID = Val(txtSampleID) Then
    cmdAddToConsultantList.Caption = "Remove from Consultant List"
  End If
  tb.MoveNext
Loop

End Sub

Private Sub FillOrgNames(ByVal Index As Integer)

Dim tb As Recordset
Dim sql As String

cmbOrgName(Index).Clear

sql = "Select * from Organisms where " & _
      "GroupName = '" & cmbOrgGroup(Index).Text & "' " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenClient 0, tb, sql
Do While Not tb.EOF
  cmbOrgName(Index).AddItem tb!Name & ""
  tb.MoveNext
Loop

End Sub

Private Sub GetSampleIDWithOffset()

SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)

End Sub


Private Sub FillLists()

FillWards Me, HospName(0)
FillClinicians Me, HospName(0)
FillGPs Me, HospName(0)

FillCastsCrystalsMiscSite

End Sub


Private Function IsChild() As Boolean

IsChild = False

If Not IsDate(txtDoB) Then Exit Function

If DateDiff("yyyy", txtDoB, Now) < 10 Then
  IsChild = True
End If

End Function

Private Function IsPregnant() As Boolean

If chkPregnant = 1 Then
  IsPregnant = True
Else
  IsPregnant = False
End If

End Function

Private Function IsOutPatient() As Boolean
  
IsOutPatient = False
  
End Function

Private Sub LoadComments()

Dim Cx As New Comment
Dim Cxs As New Comments
  
'On Error Resume Next

'txtIdentificationComment = ""
txtUrineComment = ""
txtDemographicComment = ""
txtCSComment(0) = ""
txtCSComment(1) = ""

If Trim$(txtSampleID) = "" Then Exit Sub

Set Cx = Cxs.Load(SampleIDWithOffset)
If Not Cx Is Nothing Then
  txtUrineComment = Cx.MicroGeneral
  txtDemographicComment = Cx.Demographic
  txtCSComment(0) = Cx.MicroCS
  txtCSComment(1) = Cx.MicroCS
  txtConsultantComment(0) = Cx.MicroConsultant
  txtConsultantComment(1) = Cx.MicroConsultant
End If

End Sub

Private Function LoadFaeces() As Boolean
'Returns true if Faeces results present

Dim tb As Recordset
Dim sql As String
Dim n As Integer

ClearFaeces

LoadFaeces = False

sql = "Select S.SalmIdent, S.ColindaleResult, S.ShigType, F.* " & _
      "from Faeces as F left join SalmShig as S on " & _
      "S.SampleID = F.SampleID " & _
      "where F.SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
  LoadFaeces = True
  If Not IsNull(tb!pcDone) Then
    chkPCDone = IIf(tb!pcDone, 1, 0)
  Else
    chkPCDone = 0
  End If
  If tb!pc & "" = "N" Then
    lblPC = "Negative"
    lblPC.BackColor = vbGreen
  ElseIf tb!pc & "" = "P" Then
    lblPC = "Positive"
    lblPC.BackColor = vbRed
  End If

  If Not IsNull(tb!SeleniteDone) Then
    chkSeleniteDone = IIf(tb!SeleniteDone, 1, 0)
  Else
    chkSeleniteDone = 0
  End If
  If tb!selenite & "" = "N" Then
    lblSelenite = "Negative"
    lblSelenite.BackColor = vbGreen
  ElseIf tb!selenite & "" = "P" Then
    lblSelenite = "Positive"
    lblSelenite.BackColor = vbRed
  End If

  For n = 0 To 4
    chkScreen(n) = IIf(tb!Screen And 2 ^ n, 1, 0)
    chkPurity(n) = IIf(tb!Purity And 2 ^ n, 1, 0)
    chkAPI(n) = IIf(tb!api And 2 ^ n, 1, 0)
    txtAPICode(n) = tb("APICode" & Format(n)) & ""
    txtAPIName(n) = tb("APIName" & Format(n)) & ""
    If Mid(tb!Lact & Space$(5), n + 1, 1) = "N" Then
      lblLact(n) = "Neg"
      lblLact(n).BackColor = vbGreen
    ElseIf Mid(tb!Lact & Space$(5), n + 1, 1) = "P" Then
      lblLact(n) = "Pos"
      lblLact(n).BackColor = vbRed
    End If
    If Mid(tb!Urea & Space$(5), n + 1, 1) = "N" Then
      lblUrea(n) = "Neg"
      lblUrea(n).BackColor = vbGreen
    ElseIf Mid(tb!Urea & Space$(5), n + 1, 1) = "P" Then
      lblUrea(n) = "Pos"
      lblUrea(n).BackColor = vbRed
    End If
  Next

  If Left(tb!Camp & " ", 1) = "N" Then
    lblCamp = "Negative"
    lblCamp.BackColor = vbGreen
  ElseIf Left(tb!Camp & " ", 1) = "P" Then
    lblCamp = "Positive"
    lblCamp.BackColor = vbRed
  End If
  
  If Left(tb!CampLatex & " ", 1) = "N" Then
    lblCampLatex = "Negative"
    lblCampLatex.BackColor = vbGreen
  ElseIf Left(tb!CampLatex & " ", 1) = "P" Then
    lblCampLatex = "Positive"
    lblCampLatex.BackColor = vbRed
  End If
  
  txtGram = tb!Gram & ""
  txtCampCulture = tb!CampCulture & ""

  lblSalmonella = tb!SalmIdent & ""
  lblColindale = tb!ColindaleResult & ""
  lblShigella = tb!ShigType & ""

  If tb!PC0157 & "" = "N" Then
    lblPC0157 = "Negative"
    lblPC0157.BackColor = vbGreen
  ElseIf tb!PC0157 & "" = "P" Then
    lblPC0157 = "Positive"
    lblPC0157.BackColor = vbRed
  End If
  
  If tb!PC0157Latex & "" = "N" Then
    lbl0157Latex = "Negative"
    lbl0157Latex.BackColor = vbGreen
  ElseIf tb!PC0157Latex & "" = "P" Then
    lbl0157Latex = "Positive"
    lbl0157Latex.BackColor = vbRed
  End If
  lbl0157 = tb!PC0157Report & ""

  For n = 0 To 3
    If Mid$(tb!EPC & Space$(4), n + 1, 1) = "N" Then
      lblEPC(n) = "Neg"
      lblEPC(n).BackColor = vbGreen
    ElseIf Mid$(tb!EPC & Space$(4), n + 1, 1) = "P" Then
      lblEPC(n) = "Pos"
      lblEPC(n).BackColor = vbRed
    End If
  Next

  For n = 0 To 2
    chkOccult(n) = IIf(tb!chkOccult And 2 ^ n, 1, 0)
    If Mid$(tb!Occult & Space$(3), n + 1, 1) = "N" Then
      lblOccult(n) = "Neg"
      lblOccult(n).BackColor = vbGreen
    ElseIf Mid$(tb!Occult & Space$(3), n + 1, 1) = "P" Then
      lblOccult(n) = "Pos"
      lblOccult(n).BackColor = vbRed
    End If
  Next

  If tb!Rota & "" = "N" Then
    lblRota = "Negative"
    lblRota.BackColor = vbGreen
  ElseIf tb!Rota & "" = "P" Then
    lblRota = "Positive"
    lblRota.BackColor = vbRed
  End If
  
  If tb!Adeno & "" = "N" Then
    lblAdeno = "Negative"
    lblAdeno.BackColor = vbGreen
  ElseIf tb!Adeno & "" = "P" Then
    lblAdeno = "Positive"
    lblAdeno.BackColor = vbRed
  End If
  
  If tb!ToxinAL & "" = "N" Then
    lblToxinAL = "Negative"
    lblToxinAL.BackColor = vbGreen
  ElseIf tb!ToxinAL & "" = "P" Then
    lblToxinAL = "Positive"
    lblToxinAL.BackColor = vbRed
  End If
  
  If tb!ToxinATA & "" = "N" Then
    lblToxinATA = "Negative"
    lblToxinATA.BackColor = vbGreen
  ElseIf tb!ToxinATA & "" = "P" Then
    lblToxinATA = "Positive"
    lblToxinATA.BackColor = vbRed
  End If
  
  If tb!Aus & "" = "N" Then
    lblAus = "Negative"
    lblAus.BackColor = vbGreen
  ElseIf tb!Aus & "" = "P" Then
    lblAus = "Positive"
    lblAus.BackColor = vbRed
  End If

  For n = 0 To 2
    cmbOP(n) = tb("OP" & Format(n)) & ""
  Next
End If

End Function

Private Function LoadOP() As Boolean
'Returns true if OP results present

Dim tb As Recordset
Dim sql As String
Dim n As Integer

lblCrypto = ""
lblCrypto.BackColor = &H8000000F
For n = 0 To 2
  cmbOva(n) = ""
Next

LoadOP = False

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
  
  If tb!Aus & "" = "N" Then
    lblCrypto = "Negative"
    lblCrypto.BackColor = vbGreen
    LoadOP = True
  ElseIf tb!Aus & "" = "P" Then
    lblCrypto = "Positive"
    lblCrypto.BackColor = vbRed
    LoadOP = True
  End If

  For n = 0 To 2
    cmbOva(n) = Trim$(tb("OP" & Format(n)) & "")
    If cmbOva(n) <> "" Then
      LoadOP = True
    End If
  Next
End If

End Function


Private Function LoadCDiff() As Boolean
'Returns true if Cdiff results present

Dim tb As Recordset
Dim sql As String
Dim n As Integer

lblToxinA = ""
lblToxinA.BackColor = &H8000000F
lblToxinB = ""
lblToxinB.BackColor = &H8000000F

LoadCDiff = False

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
  
  If tb!ToxinAL & "" = "N" Then
    lblToxinA = "Negative"
    lblToxinA.BackColor = vbGreen
    LoadCDiff = True
  ElseIf tb!ToxinAL & "" = "P" Then
    lblToxinA = "Positive"
    lblToxinA.BackColor = vbRed
    LoadCDiff = True
  End If
  
  If tb!ToxinATA & "" = "N" Then
    lblToxinB = "Negative"
    lblToxinB.BackColor = vbGreen
    LoadCDiff = True
  ElseIf tb!ToxinATA & "" = "P" Then
    lblToxinB = "Positive"
    lblToxinB.BackColor = vbRed
    LoadCDiff = True
  End If

End If

End Function


Private Function LoadRotaAdeno() As Boolean
'Returns true if Rota/Adeno results present

Dim tb As Recordset
Dim sql As String
Dim n As Integer

txtRota = ""
txtRota.BackColor = &H8000000F
txtAdeno = ""
txtAdeno.BackColor = &H8000000F

LoadRotaAdeno = False

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then

  If tb!Rota & "" = "N" Then
    txtRota = "Negative"
    txtRota.BackColor = vbGreen
  LoadRotaAdeno = True
  ElseIf tb!Rota & "" = "P" Then
    txtRota = "Positive"
    txtRota.BackColor = vbRed
    LoadRotaAdeno = True
  End If
  
  If tb!Adeno & "" = "N" Then
    txtAdeno = "Negative"
    txtAdeno.BackColor = vbGreen
    LoadRotaAdeno = True
  ElseIf tb!Adeno & "" = "P" Then
    txtAdeno = "Positive"
    txtAdeno.BackColor = vbRed
    LoadRotaAdeno = True
  End If

End If

End Function

Private Function LoadRSV() As Boolean
'Returns true if RSV results present

Dim tb As Recordset
Dim sql As String
Dim n As Integer

lblRSV.Caption = ""
lblRSV.BackColor = &H8000000F

LoadRSV = False

sql = "Select * from GenericResults where " & _
      "SampleID = '" & SampleIDWithOffset & "' " & _
      "and TestName = 'RSV'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then

  If tb!Result & "" = "Negative" Then
    lblRSV = "Negative"
    lblRSV.BackColor = vbGreen
    LoadRSV = True
  ElseIf tb!Result & "" = "Positive" Then
    lblRSV = "Positive"
    lblRSV.BackColor = vbRed
    LoadRSV = True
  ElseIf tb!Result & "" = "Inconclusive" Then
    lblRSV = "Inconclusive"
    lblRSV.BackColor = vbYellow
    LoadRSV = True
  End If

End If

End Function

Private Function LoadFOB() As Boolean
'Returns true if FOB results present

Dim tb As Recordset
Dim sql As String
Dim n As Integer

For n = 0 To 2
  chkFOB(n) = 0
  lblFOB(n) = ""
  lblFOB(n).BackColor = &H8000000F
Next

LoadFOB = False

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
  
  For n = 0 To 2
    If tb!chkOccult And 2 ^ n Then
      chkFOB(n) = 1
      LoadFOB = True
    End If
    If Mid$(tb!Occult & Space$(3), n + 1, 1) = "N" Then
      lblFOB(n) = "Negative"
      lblFOB(n).BackColor = vbGreen
      LoadFOB = True
    ElseIf Mid$(tb!Occult & Space$(3), n + 1, 1) = "P" Then
      lblFOB(n) = "Positive"
      lblFOB(n).BackColor = vbRed
      LoadFOB = True
    End If
  Next

End If

End Function

Private Sub LoadForcedSens()

Dim tb As Recordset
Dim sql As String
Dim Index As Integer
Dim ABName As String
Dim Report As Boolean
Dim n As Integer

sql = "Select * from ForcedABReport where " & _
      "SampleID = " & sysOptMicroOffset(0) + Val(txtSampleID)
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  Index = tb!Index
  ABName = tb!ABName
  Report = tb!Report
  For n = 1 To grdAB(Index).Rows - 1
    If grdAB(Index).TextMatrix(n, 0) = ABName Then
      grdAB(Index).Row = n
      grdAB(Index).Col = 2
      Set grdAB(Index).CellPicture = IIf(Report, imgSquareTick.Picture, imgSquareCross.Picture)
      Exit For
    End If
  Next
  tb.MoveNext
Loop

End Sub

Private Function LoadIdent() As Integer
'Returns number of Isolates Loaded

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim max As Integer

ClearIdent

max = 0

For n = 1 To 8
  sql = "Select * from UrineIdent where " & _
        "SampleID = '" & SampleIDWithOffset & "' " & _
        "and Isolate = " & n
  Set tb = New Recordset
  RecOpenClient 0, tb, sql
  
  If Not tb.EOF Then
    max = n + 1
    cmbGram(n) = tb!Gram & ""
    cmbWetPrep(n) = tb!WetPrep & ""
    txtCoagulase(n) = tb!Coagulase & ""
    txtCatalase(n) = tb!Catalase & ""
    txtOxidase(n) = tb!Oxidase & ""
    txtAPI1(n) = tb!API0 & ""
    txtAPI1(n) = tb!api1 & ""
    cmbAPI2(n) = tb!Ident0 & ""
    cmbAPI2(n) = tb!Ident1 & ""
    cmbRapidec(n) = tb!Rapidec & ""
    cmbChromogenic(n) = tb!Chromogenic & ""
    txtReincubation(n) = tb!Reincubation & ""
    txtUrineSensitivity(n) = tb!urinesensitivity & ""
    txtExtraSensitivity(n) = tb!extrasensitivity & ""
    txtNotes(n) = Trim$(tb!Notes & "")
    If txtNotes(n) <> "" Then
      cmdNotes(n).BackColor = vbYellow
    Else
      cmdNotes(n).BackColor = vbButtonFace
    End If
  End If
Next

LoadIdent = max

cmdSaveMicro.Enabled = False

End Function

Private Function LoadIdentification() As Integer
'Returns number of Isolates Loaded

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim intMax As Integer

intMax = 0

For n = 1 To 8
  cmbIdentification(n) = ""
  txtIdentification(n) = ""
Next

For n = 1 To 8
  sql = "Select * from UrineIdent where " & _
        "SampleID = '" & SampleIDWithOffset & "' " & _
        "and Isolate = " & n
  Set tb = New Recordset
  RecOpenClient 0, tb, sql
  
  If Not tb.EOF Then
    intMax = intMax + 1
    txtIdentification(n) = Trim$(tb!Notes & "")
  End If
Next

LoadIdentification = intMax

cmdSaveMicro.Enabled = False

End Function

Private Sub LoadIsolates()

Dim tb As Recordset
Dim sql As String
Dim intIsolate As Integer

For intIsolate = 1 To 8
  cmbOrgGroup(intIsolate) = ""
  cmbOrgName(intIsolate) = ""
  cmbQualifier(intIsolate) = ""
Next

sql = "Select * from Isolates where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenClient 0, tb, sql
Do While Not tb.EOF
  cmbOrgGroup(tb!IsolateNumber) = tb!OrganismGroup & ""
  cmbOrgName(tb!IsolateNumber) = tb!OrganismName & ""
  cmbQualifier(tb!IsolateNumber) = tb!Qualifier & ""
  tb.MoveNext
Loop

End Sub

Private Function LoadUrine() As Boolean
'Returns true if Urine Results Present

Dim tb As Recordset
Dim sql As String

ClearUrine

LoadUrine = False

sql = "Select * from Urine where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  LoadUrine = True
  Select Case tb!Pregnancy & ""
    Case "P": txtPregnancy = "Positive"
    Case "N": txtPregnancy = "Negative"
    Case "I": txtPregnancy = "Inconclusive"
  End Select
  txtBacteria = Trim$(tb!Bacteria & "")
  txtHCGLevel = Trim$(tb!HCGLevel & "")
  txtBenceJones = Trim$(tb!BenceJones & "")
  txtSG = Trim$(tb!SG & "")
  txtFatGlobules = Trim$(tb!FatGlobules & "")
  txtPh = Trim$(tb!pH & "")
  txtProtein = Trim$(tb!Protein & "")
  txtGlucose = Trim$(tb!Glucose & "")
  txtKetones = Trim$(tb!Ketones & "")
  txtUrobilinogen = Trim$(tb!Urobilinogen & "")
  txtBilirubin = Trim$(tb!Bilirubin & "")
  txtBloodHb = Trim$(tb!BloodHb & "")
  txtWCC = Trim$(tb!WCC & "")
  txtRCC = Trim$(tb!RCC & "")
  cmbCrystals = Trim$(tb!Crystals & "")
  cmbCasts = Trim$(tb!Casts & "")
  cmbMisc(0) = Trim$(tb!Misc0 & "")
  cmbMisc(1) = Trim$(tb!Misc1 & "")
  cmbMisc(2) = Trim$(tb!Misc2 & "")
End If

End Function

Private Sub LockCS(ByVal intIsolate As Integer, ByVal Lockit As Boolean)

cmdUnlock(intIsolate).Visible = Lockit
cmdRemoveSecondary(intIsolate).Visible = Not Lockit
cmdUseSecondary(intIsolate).Visible = Not Lockit
lblSetAllR(intIsolate).Visible = Not Lockit
lblSetAllS(intIsolate).Visible = Not Lockit

End Sub

Private Sub OrderFaeces()
    
Dim F As Form
Dim lngOrders As Long
Dim n As Integer

For n = 1 To 13
  ssTab1.TabVisible(n) = False
Next
ssTab1.TabVisible(4) = True 'c&S
ssTab1.TabVisible(12) = True 'Identification
ssTab1.TabVisible(13) = True 'Identification 5-8

Set F = New frmMicroOrderFaeces
With F
  .txtSampleID = txtSampleID
  .Show 1
  lngOrders = .FaecalOrders
End With
Unload F
Set F = Nothing

If lngOrders And 2 ^ 6 Then 'rota/adeno
  ssTab1.TabVisible(9) = True
End If
If lngOrders And 2 ^ 1 Then 'cdiff
  ssTab1.TabVisible(10) = True
End If
If (lngOrders And 2 ^ 3) Or (lngOrders And 2 ^ 4) Or (lngOrders And 2 ^ 5) Then 'fob
  ssTab1.TabVisible(8) = True
End If

If HospName(0) <> "Cavan" Then
  ssTab1.TabVisible(3) = True
Else
  ssTab1.TabVisible(11) = True
End If

cmdOrderTests.Enabled = True

End Sub

Private Sub PrintThis()

Dim tb As Recordset
Dim sql As String

PBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Len(cmbWard) = 0 Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Len(cmbGP) = 0 Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If

SaveDemographics

sql = "Select * from PrintPending where " & _
      "Department = 'M' " & _
      "and SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If
tb!SampleID = txtSampleID
tb!Ward = cmbWard
tb!Clinician = cmbClinician
tb!GP = cmbGP
tb!Department = "M"
tb!Initiator = UserName
tb!UsePrinter = pPrintToPrinter
tb.Update

End Sub

Private Sub SaveComments()

Dim Cx As New Comment
Dim Cxs As New Comments

If Trim$(txtSampleID) = "" Then Exit Sub

With Cx
  .SampleID = SampleIDWithOffset
  .Demographic = Trim$(txtDemographicComment)
  .MicroGeneral = Trim$(txtUrineComment)
  .MicroCS = Trim$(txtCSComment(0))
  .MicroConsultant = Trim$(txtConsultantComment(0))
  Cxs.Save Cx
  
End With

End Sub

Private Sub SaveFaeces()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Counter As Integer
Dim strB As String

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then tb.AddNew

tb!SampleID = SampleIDWithOffset

tb!pcDone = chkPCDone
tb!pc = Left(lblPC, 1)

tb!SeleniteDone = chkSeleniteDone
tb!selenite = Left$(lblSelenite, 1)

Counter = 0
For n = 0 To 4
  If chkScreen(n) Then Counter = Counter + 2 ^ n
Next
tb!Screen = Counter
    
Counter = 0
For n = 0 To 4
  If chkPurity(n) Then Counter = Counter + 2 ^ n
Next
tb!Purity = Counter
    
Counter = 0
For n = 0 To 4
  If chkAPI(n) Then Counter = Counter + 2 ^ n
Next
tb!api = Counter
    
For n = 0 To 4
  tb("APICode" & Format(n)) = txtAPICode(n)
  tb("APIName" & Format(n)) = txtAPIName(n)
Next

strB = ""
For n = 0 To 4
  strB = strB & Left$(lblLact(n) & " ", 1)
Next
tb!Lact = strB

strB = ""
For n = 0 To 4
  strB = strB & Left$(lblUrea(n) & " ", 1)
Next
tb!Urea = strB

tb!Camp = Left$(lblCamp, 1)
tb!CampLatex = Left$(lblCampLatex, 1)
tb!Gram = txtGram
tb!CampCulture = txtCampCulture


tb!PC0157 = Left$(lblPC0157, 1)
tb!PC0157Latex = Left$(lbl0157Latex, 1)
tb!PC0157Report = lbl0157

strB = ""
For n = 0 To 3
  strB = strB & Left(lblEPC(n) & " ", 1)
Next
tb!EPC = strB

Counter = 0
strB = ""
For n = 0 To 2
  If chkOccult(n) Then Counter = Counter + 2 ^ n
  strB = strB & Left$(lblOccult(n) & " ", 1)
Next
tb!chkOccult = Counter
tb!Occult = strB

tb!Rota = Left$(lblRota, 1)
tb!Adeno = Left$(lblAdeno, 1)

tb!ToxinAL = Left$(lblToxinAL, 1)
tb!ToxinATA = Left$(lblToxinATA, 1)

tb!Aus = Left$(lblAus, 1)
For n = 0 To 2
  tb("OP" & Format(n)) = cmbOP(n)
Next

tb.Update

End Sub

Private Sub SaveOP()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Counter As Integer
Dim strB As String

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then tb.AddNew

tb!SampleID = SampleIDWithOffset

tb!Aus = Left$(lblCrypto, 1)
For n = 0 To 2
  tb("OP" & Format(n)) = cmbOva(n)
Next

tb.Update

End Sub


Private Sub SaveCdiff()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Counter As Integer
Dim strB As String

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then tb.AddNew

tb!SampleID = SampleIDWithOffset

tb!ToxinAL = Left$(lblToxinA, 1)
tb!ToxinATA = Left$(lblToxinB, 1)

tb.Update

End Sub


Private Sub SaveRotaAdeno()

Dim tb As Recordset
Dim sql As String

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then tb.AddNew

tb!SampleID = SampleIDWithOffset

tb!Rota = Left$(txtRota, 1)
tb!Adeno = Left$(txtAdeno, 1)

tb.Update

End Sub


Private Sub SaveFOB()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Counter As Integer
Dim strB As String

sql = "Select * from Faeces where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then tb.AddNew

tb!SampleID = SampleIDWithOffset

Counter = 0
strB = ""
For n = 0 To 2
  If chkFOB(n) Then Counter = Counter + 2 ^ n
  strB = strB & Left$(lblFOB(n) & " ", 1)
Next
tb!chkOccult = Counter
tb!Occult = strB

tb.Update

End Sub

Private Sub SaveRSV()

Dim tb As Recordset
Dim sql As String

If Trim$(lblRSV.Caption) = "" Then
  sql = "Delete from GenericResults where " & _
        "SampleID = '" & SampleIDWithOffset & "' " & _
        "and TestName = 'RSV'"
  Cnxn(0).Execute sql
Else
  sql = "Select * from GenericResults where " & _
        "SampleID = '" & SampleIDWithOffset & "' " & _
        "and TestName = 'RSV'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql

  If tb.EOF Then tb.AddNew

  tb!SampleID = SampleIDWithOffset
  tb!TestName = "RSV"
  tb!Result = lblRSV.Caption

  tb.Update
End If

End Sub



Private Sub SaveIdent(ByVal Validate As Boolean)
  
Dim tb As Recordset
Dim sql As String
Dim n As Integer

For n = 1 To 8
  If IdentIsSaveable(n) Then
    sql = "Select * from UrineIdent where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and Isolate = " & n
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    
    If tb.EOF Then tb.AddNew
    
    tb!Isolate = n
    tb!SampleID = SampleIDWithOffset
    tb!Gram = cmbGram(n)
    tb!WetPrep = cmbWetPrep(n)
    tb!Coagulase = txtCoagulase(n)
    tb!Catalase = txtCatalase(n)
    tb!Oxidase = txtOxidase(n)
    tb!API0 = txtAPI1(n)
    tb!api1 = txtAPI2(n)
    tb!Ident0 = cmbAPI1(n)
    tb!Ident1 = cmbAPI2(n)
    tb!Rapidec = cmbRapidec(n)
    tb!Chromogenic = cmbChromogenic(n)
    tb!Reincubation = txtReincubation(n)
    tb!urinesensitivity = txtUrineSensitivity(n)
    tb!extrasensitivity = txtExtraSensitivity(n)
    tb!Notes = txtNotes(n)
    tb!Valid = Validate
    
    tb.Update
  Else
    sql = "Delete from UrineIdent where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and Isolate = " & n
    Cnxn(0).Execute sql
  End If

Next

End Sub
Private Sub SaveIdentification(ByVal Validate As Boolean)
  
Dim tb As Recordset
Dim sql As String
Dim n As Integer

For n = 1 To 8
  If Trim$(txtIdentification(n)) <> "" Then
    sql = "Select * from UrineIdent where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and Isolate = " & n
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    
    If tb.EOF Then tb.AddNew
    
    tb!Isolate = n
    tb!SampleID = SampleIDWithOffset
    tb!Notes = txtIdentification(n)
    tb!Valid = Validate
    
    tb.Update
  Else
    sql = "Delete from UrineIdent where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and Isolate = " & n
    Cnxn(0).Execute sql
  End If

Next

End Sub

Private Sub SaveIsolates()

Dim tb As Recordset
Dim sql As String
Dim intIsolate As Integer

For intIsolate = 1 To 8
  If cmbOrgGroup(intIsolate) <> "" Then
    sql = "Select * from Isolates where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and IsolateNumber = '" & intIsolate & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If tb.EOF Then
      tb.AddNew
      tb!SampleID = SampleIDWithOffset
      tb!IsolateNumber = intIsolate
    End If
    tb!OrganismGroup = cmbOrgGroup(intIsolate)
    tb!OrganismName = cmbOrgName(intIsolate)
    tb!Qualifier = cmbQualifier(intIsolate)
    tb.Update
  Else
    sql = "Delete from Isolates where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and IsolateNumber = '" & intIsolate & "'"
    Cnxn(0).Execute sql
  End If
Next

End Sub

Private Sub SaveUrine(ByVal Validate As Boolean)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Urine where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then tb.AddNew

tb!SampleID = SampleIDWithOffset

tb!Pregnancy = Left$(txtPregnancy, 1)
tb!Bacteria = txtBacteria
tb!HCGLevel = txtHCGLevel
tb!BenceJones = txtBenceJones
tb!SG = txtSG
tb!FatGlobules = txtFatGlobules
tb!pH = txtPh
tb!Protein = txtProtein
tb!Glucose = txtGlucose
tb!Ketones = txtKetones
tb!Urobilinogen = txtUrobilinogen
tb!Bilirubin = txtBilirubin
tb!BloodHb = txtBloodHb
tb!WCC = txtWCC
tb!RCC = txtRCC
tb!Crystals = cmbCrystals
tb!Casts = cmbCasts
tb!Misc0 = cmbMisc(0)
tb!Misc1 = cmbMisc(1)
tb!Misc2 = cmbMisc(2)
tb!Valid = Validate

tb.Update

End Sub

Private Sub SetAsForced(ByVal intIndex As Integer, _
                        ByVal strABName As String, _
                        ByVal blnReport As Boolean)

Dim tb As Recordset
Dim sql As String

sql = "Select * from ForcedABReport where " & _
      "ABName = '" & strABName & "' " & _
      "and [Index] = " & intIndex & " " & _
      "and SampleID = " & sysOptMicroOffset(0) + Val(txtSampleID)
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If
tb!SampleID = sysOptMicroOffset(0) + Val(txtSampleID)
tb!ABName = strABName
tb!Report = blnReport
tb!Index = intIndex
tb.Update

End Sub

Private Sub SetPositiveNegative(ByVal lbl As Label)

With lbl
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

End Sub

Private Sub SetViewHistory()
'
'Select Case SSTab1.Tab
'  Case 0: bHistory.Visible = False
'  Case 1: bHistory.Visible = PreviousHaem
'  Case 2: bHistory.Visible = PreviousBio
'  Case 3: bHistory.Visible = PreviousCoag
'End Select

End Sub


Private Sub chkAPI_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub

Private Sub chkFOB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub


Private Sub chkOccult_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub

Private Sub chkPCDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub

Private Sub chkPregnant_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveDemographics.Enabled = True

End Sub

Private Sub chkPurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub

Private Sub chkScreen_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub

Private Sub chkSeleniteDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSaveMicro.Enabled = True

End Sub

Private Sub cmbABSelect_Click(Index As Integer)

Dim sql As String
Dim tb As Recordset
Dim Y As Integer

grdAB(Index).AddItem cmbABSelect(Index).Text
grdAB(Index).Row = grdAB(Index).Rows - 1
grdAB(Index).Col = 0
grdAB(Index).CellBackColor = &HFFFFC0
grdAB(Index).Col = 2
Set grdAB(Index).CellPicture = Me.Picture
    
sql = "Select distinct * from Sensitivities as S, Antibiotics as A where " & _
      "SampleID = '" & SampleIDWithOffset & "' " & _
      "and IsolateNumber = '" & Index & "' " & _
      "and S.AntibioticCode = A.Code " & _
      "and AntibioticName = '" & cmbABSelect(Index).Text & "'"
Set tb = New Recordset
RecOpenClient 0, tb, sql
If Not tb.EOF Then
          
  With grdAB(Index)
    Y = .Rows - 1
    .Row = Y
    .TextMatrix(Y, 1) = tb!RSI & ""
    .TextMatrix(Y, 2) = tb!CPOFlag & ""
    .TextMatrix(Y, 3) = tb!Result & ""
    .TextMatrix(Y, 4) = Format(tb!RunDateTime, "dd/mm/yy hh:mm")
    .TextMatrix(Y, 5) = tb!UserCode & ""
    .Col = 2
    If IsNull(tb!Report) Then
      Set .CellPicture = Me.Picture
    Else
      Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
    End If
  End With

End If

cmbABSelect(Index) = ""

FillABSelect Index

cmdSaveMicro.Enabled = True

End Sub

Private Sub cmbABSelect_KeyPress(Index As Integer, KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub cmbABsInUse_Click()

Dim n As Integer

lstABsInUse.AddItem cmbABsInUse
cmbABsInUse.Visible = False
lstABsInUse.Visible = True

lblABsInUse = ""
For n = 0 To lstABsInUse.ListCount - 1
  lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
Next

End Sub


Private Sub cmbAPI1_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbAPI1_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'PA' " & _
      "and Code = '" & cmbAPI1(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbAPI1(Index) = tb!Text & ""
End If

End Sub


Private Sub cmbAPI2_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbAPI2_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'PA' " & _
      "and Code = '" & cmbAPI2(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbAPI2(Index) = tb!Text & ""
End If

End Sub


Private Sub cmbCasts_Click()

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbCasts_LostFocus()

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'CA' " & _
      "and Code = '" & UCase(cmbCasts) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbCasts = tb!Text & ""
End If

End Sub

Private Sub cmbChromogenic_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbChromogenic_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'PA' " & _
      "and Code = '" & cmbChromogenic(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbChromogenic(Index) = tb!Text & ""
End If

End Sub


Private Sub cmbConsultantVal_Click()

txtSampleID = cmbConsultantVal
txtSampleID = Format$(Val(txtSampleID))
If txtSampleID = 0 Then Exit Sub

GetSampleIDWithOffset

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveMicro.Enabled = False

End Sub

Private Sub cmbCrystals_Click()

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbCrystals_LostFocus()

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'CR' " & _
      "and Code = '" & UCase(cmbCrystals) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbCrystals = tb!Text & ""
End If

End Sub

Private Sub cmbDemogComment_Click()

txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
cmbDemogComment = ""

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbDemogComment_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbDemogComment_LostFocus()

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'DE' " & _
      "and Code = '" & cmbDemogComment & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtDemographicComment = Trim$(txtDemographicComment & " " & tb!Text & "")
Else
  txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
End If
cmbDemogComment = ""

End Sub


Private Sub cmbGram_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbGram_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'GS' " & _
      "and Code = '" & cmbGram(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbGram(Index) = tb!Text & ""
End If

End Sub


Private Sub cmbHospital_Click()

FillWards Me, cmbHospital
FillClinicians Me, cmbHospital
FillGPs Me, cmbHospital

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbIdentification_KeyPress(Index As Integer, KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub cmbMisc_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbMisc_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'MI' " & _
      "and Code = '" & cmbMisc(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbMisc(Index) = tb!Text & ""
End If

End Sub

Private Sub cmbOrgGroup_Click(Index As Integer)

FillAbGrid Index
FillABSelect Index
FillOrgNames Index

cmdSaveMicro.Enabled = True
grdAB(Index).Visible = True

End Sub

Private Sub cmbOrgGroup_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'OR' " & _
      "and Code = '" & cmbOrgGroup(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbOrgGroup(Index) = tb!Text & ""
End If

End Sub


Private Sub cmbOrgName_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub cmbOva_KeyPress(Index As Integer, KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbRapidec_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbRapidec_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'PA' " & _
      "and Code = '" & cmbRapidec(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbRapidec(Index) = tb!Text & ""
End If

End Sub


Private Sub cmbSite_Change()

lblSiteDetails = cmbSite & " " & txtSiteDetails

cmdOrderTests.Enabled = False
If cmbSite = "Faeces" Or cmbSite = "Urine" Then
  cmdOrderTests.Enabled = True
End If

End Sub

Private Sub cmbSite_KeyUp(KeyCode As Integer, Shift As Integer)

Dim Found As Boolean
Dim tb As Recordset
Dim sql As String

If Trim$(cmbSite) = "" Then
  Exit Sub
End If

Found = False
sql = "Select * from Lists where " & _
      "ListType = 'SI' " & _
      "and ( Text = '" & cmbSite & "' " & _
      "or Code = '" & cmbSite & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  cmbSite = ""
End If

cmbSiteEffects

End Sub


Private Sub cmbWetPrep_Click(Index As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbWetPrep_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select * from Lists where " & _
      "ListType = 'WP' " & _
      "and Code = '" & cmbWetPrep(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbWetPrep(Index) = tb!Text & ""
End If

End Sub


Private Sub cmdABsInUse_Click()

lstABsInUse.Visible = False
cmbABsInUse.Visible = True
cmbABsInUse.SetFocus

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmdCopyFromPrevious_Click()

Dim tb As Recordset
Dim sql As String
Dim PrevSID As Long
Dim n As Integer

PrevSID = sysOptMicroOffset(0) + Val(txtSampleID) - 1

sql = "Select * from Demographics where " & _
      "SampleID = " & PrevSID
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Trim$(tb!Hospital & "") <> "" Then
  cmbHospital = Trim$(tb!Hospital)
  lblChartNumber = Trim$(tb!Hospital) & " Chart #"
  If tb!Hospital = HospName(0) Then
    lblChartNumber.BackColor = &H8000000F
    lblChartNumber.ForeColor = vbBlack
  Else
    lblChartNumber.BackColor = vbRed
    lblChartNumber.ForeColor = vbYellow
  End If
Else
  cmbHospital = HospName(0)
  lblChartNumber.Caption = HospName(0) & " Chart #"
  lblChartNumber.BackColor = &H8000000F
  lblChartNumber.ForeColor = vbBlack
End If
If IsDate(tb!SampleDate) Then
  dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
Else
  dtSampleDate = Format$(Now, "dd/mm/yyyy")
End If
If IsDate(tb!Rundate) Then
  dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
Else
  dtRunDate = Format$(Now, "dd/mm/yyyy")
End If
mNewRecord = False
If Not IsNull(tb!RooH) Then
  cRooH(0) = IIf(tb!RooH, True, False)
  cRooH(1) = Not tb!RooH
Else
  cRooH(0) = True
End If
txtChart = tb!Chart & ""
txtAandE = tb!AandE & ""
txtNoPas = tb!NOPAS & ""
txtName = tb!PatName & ""
txtAddress(0) = tb!Addr0 & ""
txtAddress(1) = tb!Addr1 & ""
Select Case Left$(Trim$(UCase$(tb!Sex & "")), 1)
  Case "M": txtSex = "Male"
  Case "F": txtSex = "Female"
  Case Else: txtSex = ""
End Select
txtDoB = Format$(tb!DoB, "dd/mm/yyyy")
txtAge = tb!Age & ""
cmbWard = tb!Ward & ""
cmbClinician = tb!Clinician & ""
cmbGP = tb!GP & ""
txtClinDetails = tb!ClDetails & ""
If IsDate(tb!SampleDate) Then
  dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
  If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
    tSampleTime = Format$(tb!SampleDate, "hh:mm")
  Else
    tSampleTime.Mask = ""
    tSampleTime.Text = ""
    tSampleTime.Mask = "##:##"
  End If
Else
  dtSampleDate = Format$(Now, "dd/mm/yyyy")
  tSampleTime.Mask = ""
  tSampleTime.Text = ""
  tSampleTime.Mask = "##:##"
End If
'If sysOptDemoVal(0) = True Then
'  If tb!Valid = True Then
'    cmdDemoVal.Caption = "VALID"
'    Set_Demo False
'  Else
'    cmdDemoVal.Caption = "&Validate"
'    Set_Demo True
'  End If
'End If
If IsDate(tb!RecDate & "") Then
  dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
'  If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
'    tRecTime = Format$(tb!RecDate, "hh:mm")
'  Else
'    tRecTime.Mask = ""
'    tRecTime.Text = ""
'    tRecTime.Mask = "##:##"
'  End If
Else
  dtRecDate = Format$(Now, "dd/mm/yyyy")
'  tRecTime.Mask = ""
'  tRecTime.Text = ""
'  tRecTime.Mask = "##:##"
End If
'If sysOptUrgent(0) Then
'  If tb!Urgent = 1 Then
'    lblUrgent.Visible = True
'    chkUrgent.Value = 1
'    UrgentTest = True
'  Else
'    chkUrgent.Value = 0
'    UrgentTest = False
'  End If
'End If
  
cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

If sysOptBloodBank(0) Then
  If Trim$(txtChart) <> "" Then
    sql = "Select  * from PatientDetails where " & _
          "PatNum = '" & txtChart & "'"
    Set tb = New Recordset
    RecOpenClientBB 0, tb, sql
    bViewBB.Enabled = Not tb.EOF
  End If
End If

sql = "Select * from Comments where " & _
      "SampleID = " & PrevSID
Set tb = New Recordset
RecOpenClient 0, tb, sql

If Not tb.EOF Then
  txtDemographicComment = tb!Demographic & ""
  txtUrineComment = tb!MicroGeneral & ""
  txtCSComment(0) = tb!MicroCS & ""
  txtCSComment(1) = tb!MicroCS & ""
End If

sql = "Select * from MicroSiteDetails where " & _
      "SampleID = " & PrevSID
Set tb = New Recordset
RecOpenClient 0, tb, sql
If Not tb.EOF Then
  'cmbSite = tb!Site & ""
  'txtSiteDetails = tb!SiteDetails & ""
  If tb!PCA0 & "" <> "" Then lstABsInUse.AddItem tb!PCA0 & ""
  If tb!PCA1 & "" <> "" Then lstABsInUse.AddItem tb!PCA1 & ""
  If tb!PCA2 & "" <> "" Then lstABsInUse.AddItem tb!PCA2 & ""
  If tb!PCA3 & "" <> "" Then lstABsInUse.AddItem tb!PCA3 & ""
End If
lblABsInUse = ""
For n = 0 To lstABsInUse.ListCount - 1
  lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
Next

cmdCopyFromPrevious.Visible = False


End Sub

Private Sub cmdNADMicro_Click()

txtBacteria = "Nil"
txtWCC = "Nil"
txtRCC = "Nil"

End Sub

Private Sub cmdNotes_Click(Index As Integer)

If txtNotes(Index).Visible Then
  txtNotes(Index).Visible = False
Else
  txtNotes(Index).Visible = True
End If
    
If txtNotes(Index) <> "" Then
  cmdNotes(Index).BackColor = vbYellow
Else
  cmdNotes(Index).BackColor = vbButtonFace
End If

End Sub

Private Sub cmdPhone_Click()

With frmPhoneLog
  .SampleID = txtSampleID
  If cmbGP <> "" Then
    .GP = cmbGP
    .WardOrGP = "GP"
  Else
    .GP = cmbWard
    .WardOrGP = "Ward"
  End If
  .Show 1
End With

CheckIfPhoned

End Sub

Private Sub CheckIfPhoned()

If CheckPhoneLog(txtSampleID) Then
  cmdPhone.BackColor = vbYellow
  cmdPhone.Caption = "Results Phoned"
  cmdPhone.ToolTipText = "Results Phoned"
Else
  cmdPhone.BackColor = &H8000000F
  cmdPhone.Caption = "Phone Results"
  cmdPhone.ToolTipText = "Phone Results"
End If

End Sub

Private Sub cmdRemoveSecondary_Click(Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Found As Boolean
Dim ABName As String
Dim intABs As Integer

On Error GoTo ehla

sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
      "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
      "from ABDefinitions as D, Antibiotics as A where " & _
      "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
      "and D.Site = '" & cmbSite & "' " & _
      "and D.PriSec = 'S' " & _
      "and D.AntibioticName = A.AntibioticName " & _
      "order by D.ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
        "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
        "from ABDefinitions as D, Antibiotics as A where " & _
        "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
        "and D.PriSec = 'S' " & _
        "and D.AntibioticName = A.AntibioticName " & _
        "order by D.ListOrder"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If tb.EOF Then
    Exit Sub
  End If
End If
Do While Not tb.EOF
  
  Found = False
  ABName = Trim$(tb!AntibioticName & "")
  For n = 1 To grdAB(Index).Rows - 1
    If Trim$(grdAB(Index).TextMatrix(n, 0)) = ABName Then
      Found = True
      For intABs = 0 To lstABsInUse.ListCount - 1
        If lstABsInUse.List(intABs) = ABName Then
          Found = False
        End If
      Next
      Exit For
    End If
  Next
  
  If Found Then
    If grdAB(Index).Rows = 2 Then
      grdAB(Index).AddItem ""
    End If
    grdAB(Index).RemoveItem n
  End If
  
  tb.MoveNext
Loop

FillABSelect Index

cmdSaveMicro.Enabled = True

Exit Sub

ehla:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

LogError "fEditMicrobiology/cmdRemoveSecondary_Click:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub cmdSaveHold_Click()

PBar = 0

GetSampleIDWithOffset

If cmdUnlock(1).Visible Then
  iMsg "Validated Result - Cannot Save", vbCritical
  Exit Sub
End If
      
Select Case ssTab1.Tab
  Case 1:
    SaveUrine False
  Case 2, 5:
    SaveIdent False
  Case 3:
    SaveFaeces
  Case 4, 6:
    SaveIsolates
    SaveSensitivities gNO
  Case 8:    SaveFOB
  Case 9:    SaveRotaAdeno
  Case 10:   SaveCdiff
  Case 11:   SaveOP
  Case 12, 13: SaveIdentification False
  Case 14: SaveRSV
End Select

SaveComments
UpdateMRU Me

cmdSaveMicro.Enabled = False

End Sub

Private Sub cmdUnlock_Click(Index As Integer)

Dim tb As Recordset
Dim sql As String

sql = "Select Password from Users where " & _
      "Name = '" & AddTicks(UserName) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  If UCase$(iBOX("Password Required", , , True)) = UCase$(tb!PassWord & "") Then
    LockCS Index, False
  End If
End If
    
End Sub

Private Sub cmdUseSecondary_Click(Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Found As Boolean
Dim ABName As String
Dim ABCode As String
Dim tbC As Recordset
Dim Res As String
Dim RSI As String
Dim RunDateTime As String
Dim Operator As String

On Error GoTo ehla

sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
      "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
      "from ABDefinitions as D, Antibiotics as A where " & _
      "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
      "and D.Site = '" & cmbSite & "' " & _
      "and D.PriSec = 'S' " & _
      "and D.AntibioticName = A.AntibioticName " & _
      "order by D.ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
        "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
        "from ABDefinitions as D, Antibiotics as A where " & _
        "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
        "and (D.Site = 'Generic' or D.Site is Null ) and D.PriSec = 'S' " & _
        "and D.AntibioticName = A.AntibioticName " & _
        "order by D.ListOrder"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If tb.EOF Then
    Exit Sub
  End If
End If
Do While Not tb.EOF
  
  Found = False
  ABName = Trim$(tb!AntibioticName & "")
  ABCode = AntibioticCodeFor(ABName)
  sql = "Select * from Sensitivities where " & _
        "SampleID = '" & sysOptMicroOffset(0) + txtSampleID & "' " & _
        "and IsolateNumber = '" & Index & "' " & _
        "and AntibioticCode = '" & ABCode & "'"
  Set tbC = New Recordset
  RecOpenServer 0, tbC, sql
  If Not tbC.EOF Then
    RSI = tbC!RSI & ""
    Res = tbC!Result & ""
    RunDateTime = Format(tbC!RunDateTime, "dd/mm/yy hh:mm")
    Operator = tbC!UserCode & ""
  Else
    RSI = ""
    Res = ""
    RunDateTime = ""
    Operator = ""
  End If
  
  For n = 1 To grdAB(Index).Rows - 1
    If Trim$(grdAB(Index).TextMatrix(n, 0)) = ABName Then
      Found = True
      Exit For
    End If
  Next
  
  If Not Found Then
    grdAB(Index).AddItem ABName & vbTab & _
                         RSI & vbTab & _
                         vbTab & _
                         Res & vbTab & _
                         RunDateTime & vbTab & Operator
    grdAB(Index).Row = grdAB(Index).Rows - 1
    grdAB(Index).Col = 0
    grdAB(Index).CellFontBold = True
    grdAB(Index).Col = 2
    If IsChild() And Not tb!AllowIfChild Then
      Set grdAB(Index).CellPicture = imgSquareCross.Picture
      grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "C"
    ElseIf IsPregnant() And Not tb!AllowIfPregnant Then
      Set grdAB(Index).CellPicture = imgSquareCross.Picture
      grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "P"
    ElseIf IsOutPatient() And Not tb!AllowIfOutPatient Then
      Set grdAB(Index).CellPicture = imgSquareCross.Picture
      grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "O"
    Else
      Set grdAB(Index).CellPicture = imgSquareCross.Picture
    End If
  End If
  
  tb.MoveNext
Loop

FillABSelect Index

cmdSaveMicro.Enabled = True

Exit Sub

ehla:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

LogError "fEditMicrobiology/cmdUseSecondary_Click:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub cMRU_GotFocus()

If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
  If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
    GetSampleIDWithOffset
    SaveDemographics
    cmdSaveDemographics.Enabled = False
    cmdSaveInc.Enabled = False
  End If
End If

End Sub

Private Sub cmdAddToConsultantList_Click()

Dim sql As String
Dim tb As Recordset

Select Case Left$(cmdAddToConsultantList.Caption, 3)
  Case "Add":
    sql = "Select * from ConsultantList where " & _
          "SampleID = " & sysOptMicroOffset(0) + Val(txtSampleID)
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If tb.EOF Then tb.AddNew
    tb!SampleID = sysOptMicroOffset(0) + Val(txtSampleID)
    tb.Update
  
  Case "Rem":
    sql = "Delete from ConsultantList " & _
          "where SampleID = '" & sysOptMicroOffset(0) + Val(txtSampleID) & "'"
    Cnxn(0).Execute sql
End Select

FillForConsultantValidation

End Sub

Private Sub dtRecDate_CloseUp()

PBar = 0

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub grdAB_Click(Index As Integer)

Dim s As String

If cmdUnlock(Index).Visible Then Exit Sub

cmdSaveMicro.Enabled = True

With grdAB(Index)
  If .MouseRow = 0 Then Exit Sub

  If .CellBackColor = &HFFFFC0 Then
    .Enabled = False
    If iMsg("Remove " & Trim$(.Text) & " from List?", vbQuestion + vbYesNo) = vbYes Then
      .RemoveItem .Row
      FillABSelect Index
    End If
    .Enabled = True
  ElseIf .Col = 1 Then
    s = Trim$(.TextMatrix(.Row, 1))
    Select Case s
      Case "": s = "R"
      Case "R": s = "S"
      Case "S": s = "I"
      Case "I": s = ""
      Case Else: s = ""
    End Select
    .TextMatrix(.Row, 1) = s
  ElseIf .Col = 2 Then
    If .CellPicture = imgSquareTick.Picture Then
      Set .CellPicture = imgSquareCross.Picture
      SetAsForced Index, .TextMatrix(.Row, 0), False
    Else
      If .TextMatrix(.Row, 2) = "C" Then
        If MsgBox("Report " & .TextMatrix(.Row, 0) & " on a Child?", vbQuestion + vbYesNo) = vbNo Then
          Exit Sub
        End If
      ElseIf .TextMatrix(.Row, 2) = "P" Then
        If MsgBox("Report " & .TextMatrix(.Row, 0) & " for Pregnant Patient?", vbQuestion + vbYesNo) = vbNo Then
          Exit Sub
        End If
      ElseIf .TextMatrix(.Row, 2) = "O" Then
        If MsgBox("Report " & .TextMatrix(.Row, 0) & " for an Out-Patient?", vbQuestion + vbYesNo) = vbNo Then
          Exit Sub
        End If
      End If
      Set .CellPicture = imgSquareTick.Picture
      SetAsForced Index, .TextMatrix(.Row, 0), True
    End If
  End If
End With

End Sub

Private Sub imgMoreCS_Click()

ssTab1.TabVisible(5) = True
ssTab1.TabVisible(6) = True
imgMoreCS.Visible = False
lblMoreCS.Visible = False

End Sub

Private Sub imgMoreID_Click()

ssTab1.TabVisible(5) = True
ssTab1.TabVisible(6) = True
imgMoreID.Visible = False
lblMoreID.Visible = False

End Sub

Private Sub iRecDate_Click(Index As Integer)

If Index = 0 Then
  dtRecDate = DateAdd("d", -1, dtRecDate)
Else
  If DateDiff("d", dtRecDate, Now) > 0 Then
    dtRecDate = DateAdd("d", 1, dtRecDate)
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

End Sub

Private Sub lblAus_Click()

Dim n As Integer

With lblAus
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
    For n = 1 To 2
      If cmbOP(n) = "No Cryptosporidium Oocysts Seen" Then
        Exit Sub
      End If
    Next
    For n = 1 To 2
      If cmbOP(n) = "Cryptosporidium Oocysts Seen" Then
        cmbOP(n) = "No Cryptosporidium Oocysts Seen"
        Exit Sub
      End If
    Next
  For n = 1 To 2
    If cmbOP(n) = "" Then
      cmbOP(n) = "No Cryptosporidium Oocysts Seen"
      Exit For
    End If
  Next
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
    For n = 1 To 2
      If cmbOP(n) = "Cryptosporidium Oocysts Seen" Then
        Exit Sub
      End If
    Next
    For n = 1 To 2
      If cmbOP(n) = "No Cryptosporidium Oocysts Seen" Then
        cmbOP(n) = "Cryptosporidium Oocysts Seen"
        Exit Sub
      End If
    Next
    For n = 1 To 2
      If cmbOP(n) = "" Then
        cmbOP(n) = "Cryptosporidium Oocysts Seen"
        Exit For
      End If
    Next
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
    For n = 1 To 2
      If cmbOP(n) = "No Cryptosporidium Oocysts Seen" Then
        cmbOP(n) = ""
        Exit Sub
      End If
    Next
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblColindale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With frmSalmShigWorkSheet
  .lblChart = lblChart
  .lblDoB = lblDoB
  .lblAge = lblAge
  .lblSex = lblSex
  .lblName = lblName
  .lblSampleID = txtSampleID
  .Show 1
End With

End Sub


Private Sub lblCrypto_Click()

Dim n As Integer

With lblCrypto
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblFOB_Click(Index As Integer)

With lblFOB(Index)
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblMoreCS_Click()

ssTab1.TabVisible(5) = True
ssTab1.TabVisible(6) = True
imgMoreCS.Visible = False
lblMoreCS.Visible = False

End Sub

Private Sub lblMoreID_Click()

ssTab1.TabVisible(5) = True
ssTab1.TabVisible(6) = True
imgMoreID.Visible = False
lblMoreID.Visible = False

End Sub

Private Sub lblRSV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With lblRSV
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = "Inconclusive"
    .BackColor = vbYellow
  Case "Inconclusive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub


Private Sub lblSalmonella_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With frmSalmShigWorkSheet
  .lblChart = lblChart
  .lblDoB = lblDoB
  .lblAge = lblAge
  .lblSex = lblSex
  .lblName = lblName
  .lblSampleID = txtSampleID
  .Show 1
End With

End Sub


Private Sub lblSetAllR_Click(Index As Integer)

Dim Y As Integer

With grdAB(Index)
  For Y = 1 To .Rows - 1
    If .TextMatrix(Y, 0) <> "" Then
      .TextMatrix(Y, 1) = "R"
    End If
  Next
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblSetAllS_Click(Index As Integer)

Dim Y As Integer

With grdAB(Index)
  For Y = 1 To .Rows - 1
    If .TextMatrix(Y, 0) <> "" Then
      .TextMatrix(Y, 1) = "S"
    End If
  Next
End With

cmdSaveMicro.Enabled = True

End Sub


Private Sub lblShigella_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With frmSalmShigWorkSheet
  .lblChart = lblChart
  .lblDoB = lblDoB
  .lblAge = lblAge
  .lblSex = lblSex
  .lblName = lblName
  .lblSampleID = txtSampleID
  .Show 1
End With

End Sub


Private Sub lblToxinA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With lblToxinA
  Select Case .Caption
  Case ""
    .Caption = "Not Detected"
    .BackColor = vbGreen
  Case "Not Detected"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = "Inconclusive"
    .BackColor = vbYellow
  Case "Inconclusive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
  lblToxinB.Caption = .Caption
  lblToxinB.BackColor = .BackColor
End With

cmdSaveMicro.Enabled = True

End Sub


Private Sub lblToxinAL_Click()

SetPositiveNegative lblToxinAL

cmdSaveMicro.Enabled = True

End Sub

Private Sub bDoB_Click()

PBar = 0

With fpathistory
  If HospName(0) = "Monaghan" And ssTab1.Tab = 0 Then
    .oHD(0) = True
  Else
    .oHD(1) = True
  End If
  .oFor(2) = True
  .txtName = txtDoB
  .FromEdit = True
  .EditScreen = Me
  .bsearch = True
  If Not .NoPreviousDetails Then
    .Show 1
  Else
    FlashNoPrevious Me
  End If
End With

End Sub

Private Sub bFAX_Click(Index As Integer)

PBar = 0

End Sub

Private Sub bHistory_Click()

PBar = 0

With frmMicroReport
  .lblChart = txtChart
  .lblName = txtName
  .Show 1
End With

End Sub



Private Sub cmbSite_Click()

cmbSiteEffects

End Sub


Private Sub cmdNAD_Click()

If txtProtein = "" Then txtProtein = "Nil"
If txtGlucose = "" Then txtGlucose = "Nil"
If txtKetones = "" Then txtKetones = "Nil"
If txtWCC = "" Then txtWCC = "Nil"
If txtRCC = "" Then txtRCC = "Nil"
If cmbCasts = "" Then cmbCasts = "Nil"
If cmbCrystals = "" Then cmbCrystals = "Nil"
If txtBilirubin = "" Then txtBilirubin = "Nil"
If txtUrobilinogen = "" Then txtUrobilinogen = "Nil"
If txtBloodHb = "" Then txtBloodHb = "Nil"

cmdSaveMicro.Enabled = True

End Sub

Private Sub cmdOrderTests_Click()

PBar = 0

If cmbSite = "Urine" Then
  With frmMicroOrderUrine
    .txtSampleID = txtSampleID
    .Show 1
  End With
ElseIf cmbSite = "Faeces" Then
  OrderFaeces
Else
  With frmMicroOrders
    .txtSampleID = txtSampleID
    .Show 1
  End With
End If

End Sub


Private Sub bPrint_Click()

If cmdSaveMicro.Enabled Then
  cmdSaveHold_Click
End If

PrintThis

txtSampleID = Format$(Val(txtSampleID) + 1)
GetSampleIDWithOffset
LoadAllDetails

End Sub

Private Sub SaveDemographics()

Dim sql As String
Dim tb As Recordset
Dim Hosp As String
Dim n As Integer

On Error GoTo eh3

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

SaveComments

If Trim$(tSampleTime) <> "__:__" Then
  If Not IsDate(tSampleTime) Then
    iMsg "Invalid Time", vbExclamation
    Exit Sub
  End If
End If

If InStr(lblChartNumber, "Cavan") Then
  Hosp = "Cavan"
ElseIf InStr(lblChartNumber, "Monaghan") Then
  Hosp = "Monaghan"
Else
  Hosp = ""
End If
  
sql = "Select * from MicroSiteDetails where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenClient 0, tb, sql
If tb.EOF Then tb.AddNew
tb!SampleID = SampleIDWithOffset
tb!Site = cmbSite
tb!SiteDetails = txtSiteDetails
For n = 0 To 3
  tb("PCA" & Format(n)) = ""
Next
For n = 0 To lstABsInUse.ListCount - 1
  If n < 4 Then
    tb("PCA" & Format(n)) = lstABsInUse.List(n)
  End If
Next
tb.Update

sql = "Select * from Demographics where " & _
      "SampleID = '" & SampleIDWithOffset & "'"

Set tb = New Recordset
RecOpenClient 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Fasting = 0
  tb!ForESR = 0
  tb!ForBio = 0
  tb!Faxed = 0
  tb!ForHbA1c = 0
  tb!ForPSA = 0
  tb!ForCoag = AreResultsPresent("Coag", txtSampleID)
  tb!ForFerritin = 0
  tb!ForHaem = AreResultsPresent("Haem", txtSampleID)
End If

tb!RooH = cRooH(0)

'If IsDate(tRecTime) Then
'  tb!RecDate = Format$(dtRecDate & " " & tRecTime, "dd/mmm/yyyy hh:mm")
'Else
  tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy")
'End If
tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
If IsDate(tSampleTime) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
Else
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
End If
tb!SampleID = SampleIDWithOffset
tb!Chart = txtChart
tb!PatName = Trim$(txtName)
If IsDate(txtDoB) Then
  tb!DoB = Format$(txtDoB, "dd/mmm/yyyy")
Else
  tb!DoB = Null
End If
tb!Age = txtAge
tb!Sex = Left$(txtSex, 1)
tb!Addr0 = txtAddress(0)
tb!Addr1 = txtAddress(1)
tb!Ward = Left$(cmbWard, 50)
tb!Clinician = Left$(cmbClinician, 50)
tb!GP = Left$(cmbGP, 50)
tb!ClDetails = txtClinDetails
tb!Hospital = Hosp
tb!Pregnant = chkPregnant
tb!Operator = Left$(UserName, 20)
tb.Update

LogTimeOfPrinting SampleIDWithOffset, "D"

Screen.MousePointer = 0


Exit Sub

eh3:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

LogError "fEditMicrobiology/SaveDemographics:" & Format$(er) & ":" & ers
Exit Sub

End Sub


Private Sub bPrintAll_Click()

'Dim tb As Recordset
'Dim sql As String
'
'pBar = 0
'
'If Trim$(txtSex) = "" Then
'  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
'    Exit Sub
'  End If
'End If
'
'If Trim$(txtSampleID) = "" Then
'  iMsg "Must have Lab Number.", vbCritical
'  Exit Sub
'End If
'
'If Trim$(cmbWard) = "" Then
'  iMsg "Must have Ward entry.", vbCritical
'  Exit Sub
'End If
'
'If Trim$(cmbWard) = "GP" Then
'  If Trim$(cmbGP) = "" Then
'    iMsg "Must have Ward or GP entry.", vbCritical
'    Exit Sub
'  End If
'End If
'
'SaveDemographics
'
'If SSTab1.Tab <> 0 Then
'  sql = "Select * from PrintPending where " & _
'        "Department = 'D' " & _
'        "and SampleID = '" & txtSampleID & "'"
'  Set tb = New Recordset
'  recopenclient 0, tb, sql
'  If tb.EOF Then
'    tb.AddNew
'  End If
'  tb!SampleID = txtSampleID
'  tb!Department = "D"
'  tb!Initiator = UserName
'  tb!UsePrinter = pPrintToPrinter
'  tb.Update
'End If
'
'SaveCoag 1
'sql = "Update CoagResults " & _
'      "Set Valid = 1, Printed = 1 where " & _
'      "SampleID = '" & txtSampleID & "'"
'Cnxn(0).Execute sql
'
'txtSampleID = Format$(Val(txtSampleID) + 1)
'LoadAllDetails
'
End Sub

Private Sub bPrintHold_Click()

If cmdSaveMicro.Enabled Then
  cmdSaveHold_Click
End If

PrintThis

End Sub

Private Sub cmdSaveMicro_Click()

PBar = 0

GetSampleIDWithOffset

If cmdUnlock(1).Visible Then
  iMsg "Validated Result - Cannot Save", vbCritical
  Exit Sub
End If

Select Case ssTab1.Tab
  Case 1:    SaveUrine False
  Case 2, 5: SaveIdent False
  Case 3:    SaveFaeces
  Case 4, 6: SaveIsolates
             SaveSensitivities gNO
  Case 8:    SaveFOB
  Case 9:    SaveRotaAdeno
  Case 10:   SaveCdiff
  Case 11:   SaveOP
  Case 12, 13: SaveIdentification False
End Select

SaveComments
SaveRSV
UpdateMRU Me

txtSampleID = Format$(Val(txtSampleID) + 1)

GetSampleIDWithOffset
LoadAllDetails

cmdSaveMicro.Enabled = False

End Sub

Private Sub SaveSensitivities(ByVal Validate As Integer)

Dim tb As Recordset
Dim sql As String
Dim intOrg As Integer
Dim n As Integer
Dim ABCode As String
Dim ReportCounter As Integer

On Error GoTo ehss

ReportCounter = 0

For intOrg = 1 To 8
  
  With grdAB(intOrg)
    
    For n = 1 To .Rows - 1
      If .TextMatrix(n, 0) <> "" Then
        ABCode = AntibioticCodeFor(.TextMatrix(n, 0))
        sql = "Select * from Sensitivities where " & _
              "SampleID = '" & SampleIDWithOffset & "' " & _
              "and IsolateNumber = '" & intOrg & "' " & _
              "and AntibioticCode = '" & ABCode & "'"
        Set tb = New Recordset
        RecOpenServer 0, tb, sql
        If tb.EOF Then
          tb.AddNew
          tb!Rundate = Format(Now, "dd/mmm/yyyy")
          tb!RunDateTime = Format(Now, "dd/mmm/yyyy hh:mm")
          tb!UserCode = UserCode
        End If
        tb!SampleID = SampleIDWithOffset
        tb!IsolateNumber = intOrg
        tb!AntibioticCode = ABCode
        tb!RSI = .TextMatrix(n, 1)
        tb!CPOFlag = .TextMatrix(n, 2)
        tb!Result = .TextMatrix(n, 3)
        
        .Row = n
        .Col = 0
        If .CellFontBold = True Then
          tb!Secondary = 1
        Else
          tb!Secondary = 0
        End If
        If .CellBackColor = &HFFFFC0 Then
          tb!Forced = 1
        Else
          tb!Forced = 0
        End If
        .Col = 2
                 
        If .CellPicture = 0 Then
          If .TextMatrix(n, 1) = "R" Then
            tb!Report = 1
          ElseIf .TextMatrix(n, 1) = "S" Then
            ReportCounter = ReportCounter + 1
            If ReportCounter < 4 Then
              tb!Report = 1
            Else
              tb!Report = 0
            End If
          Else
            tb!Report = Null
          End If
        Else
          If .CellPicture = imgSquareTick.Picture Then
            tb!Report = 1
          ElseIf .CellPicture = imgSquareCross.Picture Then
            tb!Report = 0
          Else
            tb!Report = Null
          End If
        End If
        tb.Update
      End If
    Next
  End With
  
Next

If Validate = gYES Then
  sql = "Update Sensitivities " & _
        "Set Valid = 1, " & _
        "AuthoriserCode = '" & UserCode & "' " & _
        "where SampleID = '" & SampleIDWithOffset & "'"
  Cnxn(0).Execute sql
ElseIf Validate = gNO Then
  sql = "Update Sensitivities " & _
        "Set Valid = 0, " & _
        "AuthoriserCode = NULL " & _
        "where SampleID = '" & SampleIDWithOffset & "'"
  Cnxn(0).Execute sql
End If

Exit Sub

ehss:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "fEditMicrobiology/SaveSensitivities:" & Format(er) & ":" & ers
Exit Sub

End Sub


Private Function LoadSensitivities() As Integer
'Returns number of Isolates

Dim tb As Recordset
Dim sql As String
Dim intIsolate As Integer
Dim s As String
Dim max As Integer
Dim n As Integer
Dim t As Single
Dim strInclude As String
Dim Y As Integer
Dim ReportCounter As Integer

On Error GoTo ehls

t = Timer

ReportCounter = 0

For intIsolate = 1 To 8
  With grdAB(intIsolate)
    .Visible = False
    .Rows = 2
    .AddItem ""
    .RemoveItem 1
  
    If cmbOrgGroup(intIsolate) <> "" Then
      FillAbGrid (intIsolate)
    End If
  
    strInclude = ""
    For n = 1 To .Rows - 1
      If .TextMatrix(n, 0) <> "" Then
        strInclude = strInclude & " AntibioticName = '" & .TextMatrix(n, 0) & "' or"
      End If
    Next
    If strInclude <> "" Then
      strInclude = Left(strInclude, Len(strInclude) - 2)
    End If
    sql = "Select distinct * from Sensitivities as S, Antibiotics as A where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "and IsolateNumber = '" & intIsolate & "' " & _
          "and S.AntibioticCode = A.Code " & _
          "and ("
    If strInclude <> "" Then
      sql = sql & "(" & strInclude & ") or "
    End If
    sql = sql & " (Forced = 1 or Secondary = 1) )"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    If Not tb.EOF Then
      max = intIsolate + 1
      Do While Not tb.EOF
        If tb!Forced Or tb!Secondary Then
          s = tb!AntibioticName & vbTab & _
              tb!RSI & vbTab & _
              tb!CPOFlag & vbTab & _
              tb!Result & vbTab & _
              Format(tb!RunDateTime, "dd/mm/yy hh:mm") & vbTab & _
              tb!UserCode & ""
          .AddItem s
          .Row = .Rows - 1
          .Col = 0
          If tb!Forced Then
            .CellBackColor = &HFFFFC0
          ElseIf tb!Secondary Then
            .CellFontBold = True
          End If
          .Col = 2
          If IsNull(tb!Report) Then
            Set .CellPicture = Me.Picture
          Else
            Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
          End If
          
        Else
          For Y = 1 To .Rows - 1
            If .TextMatrix(Y, 0) = Trim$(tb!AntibioticName & "") Then
              .TextMatrix(Y, 1) = tb!RSI & ""
              .TextMatrix(Y, 2) = tb!CPOFlag & ""
              .TextMatrix(Y, 3) = tb!Result & ""
              .TextMatrix(Y, 4) = Format(tb!RunDateTime, "dd/mm/yy hh:mm")
              .TextMatrix(Y, 5) = tb!UserCode & ""
              .Row = Y
              .Col = 2
              If IsNull(tb!Report) Then
                If .TextMatrix(Y, 1) = "R" Then
                  Set .CellPicture = imgSquareTick.Picture
                ElseIf .TextMatrix(Y, 1) = "S" Then
                  ReportCounter = ReportCounter + 1
                  If ReportCounter < 4 Then
                    Set .CellPicture = imgSquareTick.Picture
                  Else
                    Set .CellPicture = imgSquareCross.Picture
                  End If
                Else
                  Set .CellPicture = Me.Picture
                End If
              Else
                Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
              End If
            End If
          Next
        End If
        tb.MoveNext
      Loop
    End If
        
    
    .Visible = True
    
    FillABSelect intIsolate
    
    cmdUnlock(intIsolate).Visible = False
    If CheckIfValid(intIsolate) Then
      LockCS intIsolate, True
    Else
      LockCS intIsolate, False
      AutoFillReport intIsolate
    End If
    
  End With
  
Next
    
LoadSensitivities = max - 1

Debug.Print "Load Sens"; Timer - t

LoadForcedSens

SaveSensitivities gNOCHANGE

Exit Function

ehls:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "fEditMicrobiology/LoadSensitivities:" & Format(er) & ":" & ers
Exit Function

End Function


Private Sub AutoFillReport(ByVal IsolateNumber As Integer)

Dim Y As Integer
Dim SCounter As Integer
Dim tb As Recordset
Dim sql As String

sql = "Select [Default] from Lists where " & _
      "listType = 'SI' " & _
      "and Text like '" & cmbSite & "%' " & _
      "and [Default] is not null"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  SCounter = Val(tb!Default & "")
End If
If SCounter = 0 Then
  SCounter = 3
End If

With grdAB(IsolateNumber)
  If .Rows > 2 Then
    For Y = 1 To .Rows - 1
      .Row = Y
      .Col = 2
      If Trim$(.TextMatrix(Y, 1)) = "R" Then
        Set .CellPicture = imgSquareTick.Picture
      ElseIf Trim$(.TextMatrix(Y, 1)) = "S" Then
        If SCounter > 0 Then
          SCounter = SCounter - 1
          Set .CellPicture = imgSquareTick.Picture
        Else
          Set .CellPicture = imgSquareCross.Picture
        End If
      Else
        Set .CellPicture = imgSquareCross.Picture
      End If
    Next
  End If
End With
            
End Sub

Private Sub cmdSaveDemographics_Click()

PBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(txtName) <> "" Then
  If Trim$(cmbWard) = "" Then
    iMsg "Must have Ward entry.", vbCritical
    Exit Sub
  End If
  
  If Trim$(cmbWard) = "GP" Then
    If Trim$(cmbGP) = "" Then
      iMsg "Must have GP entry.", vbCritical
      Exit Sub
    End If
  End If
End If

cmdSaveDemographics.Caption = "Saving"

GetSampleIDWithOffset

SaveDemographics
UpdateMRU Me

cmdSaveDemographics.Caption = "Save && &Hold"
cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False

End Sub


Private Sub cmdSaveInc_Click()

PBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(txtName) <> "" Then
  If Trim$(cmbWard) = "" Then
    iMsg "Must have Ward entry.", vbCritical
    Exit Sub
  End If
  
  If Trim$(cmbWard) = "GP" Then
    If Trim$(cmbGP) = "" Then
      iMsg "Must have GP entry.", vbCritical
      Exit Sub
    End If
  End If
End If

If lblChartNumber.BackColor = vbRed Then
  If iMsg("Confirm this Patient has" & vbCrLf & _
          lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
    Exit Sub
  End If
End If
          
cmdSaveDemographics.Caption = "Saving"

GetSampleIDWithOffset

SaveDemographics
UpdateMRU Me

cmdSaveDemographics.Caption = "Save && &Hold"
cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False

txtSampleID = Format$(Val(txtSampleID) + 1)

GetSampleIDWithOffset
LoadAllDetails

cmdSaveMicro.Enabled = False

End Sub

Private Sub bsearch_Click()

PBar = 0

With fpathistory
  If HospName(0) = "Monaghan" And ssTab1.Tab = 0 Then
    .oHD(0) = True
  Else
    .oHD(1) = True
  End If
  .oFor(0) = True
  .txtName = txtName
  .FromEdit = True
  .EditScreen = Me
  .bsearch = True
  If Not .NoPreviousDetails Then
    .Show 1
  Else
    FlashNoPrevious Me
  End If
End With

End Sub

Private Sub cmdValidateMicro_Click()

PBar = 0

GetSampleIDWithOffset

Select Case ssTab1.Tab
  Case 1:
    SaveUrine True
  Case 2, 5:
    SaveIdent True
  Case 3:
    SaveFaeces
  Case 4, 6:
    SaveIsolates
    SaveSensitivities gYES
End Select

SaveRSV
SaveComments
UpdateMRU Me

PrintThis

cmdSaveMicro.Enabled = False

End Sub

Private Sub bViewBB_Click()

PBar = 0

If Trim$(txtChart) <> "" Then
  fViewBB.lchart = txtChart
  fViewBB.Show 1
End If

End Sub


Private Sub LoadAllDetails()

ssTab1.TabCaption(1) = "Urine"
ssTab1.TabCaption(2) = "Identification"
ssTab1.TabCaption(3) = "Faeces"
ssTab1.TabCaption(4) = "C && S"
ssTab1.TabCaption(5) = "Identification 5/8"
ssTab1.TabCaption(6) = "C && S 5/8"

LoadDemographics

ssTab1.TabVisible(1) = False
If HospName(0) = "Cavan" Then
  ssTab1.TabVisible(12) = True
  ssTab1.TabVisible(2) = False
  ssTab1.TabVisible(3) = False
Else
  ssTab1.TabVisible(2) = True
  ssTab1.TabVisible(3) = False
End If
ssTab1.TabVisible(4) = True
ssTab1.TabVisible(5) = False
ssTab1.TabVisible(6) = False

FaecesLoaded = False
UrineLoaded = False
ClearIndividualFaeces

If cmbSite = "Urine" Then
  ssTab1.TabVisible(1) = True
  If LoadUrine() Then
    UrineLoaded = True
  End If
ElseIf cmbSite = "Faeces" Then
  If HospName(0) = "Cavan" Then
    ssTab1.TabVisible(8) = True
    If LoadFOB() Then
      FOBLoaded = True
    End If
    ssTab1.TabVisible(9) = True
    If LoadRotaAdeno() Then
      RotaAdenoLoaded = True
    End If
    ssTab1.TabVisible(10) = True
    If LoadCDiff() Then
      CdiffLoaded = True
    End If
    ssTab1.TabVisible(11) = True
    If LoadOP() Then
      OPLoaded = True
    End If
  Else
    ssTab1.TabVisible(3) = True
    If LoadFaeces() Then
      FaecesLoaded = True
    End If
  End If
End If

imgMoreID.Visible = True
lblMoreID.Visible = True
imgMoreCS.Visible = True
lblMoreCS.Visible = True
If HospName(0) <> "Cavan" Then
  Select Case LoadIdent()
    Case 0:
      IdentLoaded = False
    Case 1, 2, 3, 4:
      IdentLoaded = True
    Case 5, 6, 7, 8:
      IdentLoaded = True
      ssTab1.TabVisible(5) = True
      ssTab1.TabVisible(6) = True
      imgMoreID.Visible = False
      lblMoreID.Visible = False
      imgMoreCS.Visible = False
      lblMoreCS.Visible = False
  End Select
Else
  Select Case LoadIdentification()
    Case 0:
      IdentificationLoaded = False
    Case 1, 2, 3, 4:
      IdentificationLoaded = True
      ssTab1.TabVisible(13) = False
      imgMoreIdentity.Visible = True
      lblMoreIdentity.Visible = True
      imgMoreCS.Visible = True
      lblMoreCS.Visible = True
    Case 5, 6, 7, 8:
      IdentificationLoaded = True
      ssTab1.TabVisible(13) = True
      ssTab1.TabVisible(6) = True
      imgMoreIdentity.Visible = False
      lblMoreIdentity.Visible = False
      imgMoreCS.Visible = False
      lblMoreCS.Visible = False
  End Select
End If
LoadIsolates
'CSLoaded = False

Select Case LoadSensitivities()
  Case 0:
    CSLoaded = False
  Case 1, 2, 3, 4:
    CSLoaded = True
  Case 5, 6, 7, 8:
    CSLoaded = True
    ssTab1.TabVisible(5) = True
    ssTab1.TabVisible(6) = True
    imgMoreID.Visible = False
    lblMoreID.Visible = False
    imgMoreCS.Visible = False
    lblMoreCS.Visible = False
End Select

If HospName(0) = "Cavan" Then
  ssTab1.TabVisible(2) = False
Else
  ssTab1.TabVisible(2) = True
End If
ssTab1.TabVisible(4) = True
  
LoadComments
LoadRSV

SetViewHistory

FillForConsultantValidation

EnableCopyFrom

CheckIfPhoned

End Sub
Private Sub bcancel_Click()

PBar = 0

Unload Me

End Sub

Private Sub cmbClinDetails_Click()

txtClinDetails = txtClinDetails & cmbClinDetails & " "
cmbClinDetails.ListIndex = -1

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbClinDetails_LostFocus()

Dim tb As Recordset
Dim sql As String

PBar = 0

If Trim$(cmbClinDetails) = "" Then Exit Sub

sql = "Select * from Lists where " & _
      "ListType = 'CD' " & _
      "and Code = '" & cmbClinDetails & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbClinDetails = tb!Text & ""
End If

End Sub


Private Sub cmbClinician_Click()

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbClinician_LostFocus()

PBar = 0
cmbClinician = QueryKnown("Clin", cmbClinician, cmbHospital)

End Sub

Private Sub cmbGP_Change()

lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub cmbGP_Click()

PBar = 0

lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbGP_LostFocus()

cmbGP = QueryKnown("GP", cmbGP)

End Sub


Private Sub cmdSetPrinter_Click()

frmForcePrinter.From = Me
frmForcePrinter.Show 1

If pPrintToPrinter = "Automatic Selection" Then
  pPrintToPrinter = ""
End If

If pPrintToPrinter <> "" Then
  cmdSetPrinter.BackColor = vbRed
  cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
Else
  cmdSetPrinter.BackColor = vbButtonFace
  pPrintToPrinter = ""
  cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
End If
  
End Sub

Private Sub cMRU_Click()

txtSampleID = cMRU

GetSampleIDWithOffset

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveMicro.Enabled = False

End Sub


Private Sub cMRU_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cRooH_Click(Index As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbWard_Change()

lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub cmbWard_Click()

lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbWard_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub cmbWard_LostFocus()

Dim Hospital As String

Hospital = ListCodeFor("HO", HospName(0))

cmbWard = GetWard(cmbWard, Hospital)

If Trim$(cmbWard) = "" Then
  cmbWard = "GP"
  Exit Sub
End If

End Sub



Private Sub dtRunDate_CloseUp()

PBar = 0

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub dtSampleDate_CloseUp()

PBar = 0

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub Form_Activate()

TimerBar.Enabled = True
PBar = 0

End Sub

Private Sub FillOrganisms()

Dim n As Integer
Dim tb As Recordset
Dim sql As String
Dim Temp As String

sql = "Select * from Lists where " & _
      "ListType = 'OR' " & _
      "order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql

For n = 1 To 8
  cmbOrgGroup(n).Clear
  cmbOrgName(n).Clear
  cmbRapidec(n).Clear
  cmbRapidec(n).AddItem "Pending"
  cmbChromogenic(n).Clear
Next

Do While Not tb.EOF
  Temp = tb!Text & ""
  For n = 1 To 8
    cmbOrgGroup(n).AddItem Temp
    cmbAPI1(n).AddItem Temp
    cmbAPI2(n).AddItem Temp
    cmbChromogenic(n).AddItem Temp
    If UCase(Left(Temp, 5)) = "STAPH" Then
      cmbRapidec(n).AddItem Temp
    End If
  Next
  tb.MoveNext
Loop

End Sub

Private Sub ClearFaeces()
  
Dim n As Integer

chkPCDone = False
lblPC = ""
lblPC.BackColor = &H8000000F
chkSeleniteDone = False
lblSelenite = ""
lblSelenite.BackColor = &H8000000F

For n = 0 To 4
  chkScreen(n) = False
  chkPurity(n) = False
  chkAPI(n) = False
  txtAPICode(n) = ""
  txtAPIName(n) = ""
  lblLact(n) = ""
  lblLact(n).BackColor = &H8000000F
  lblUrea(n) = ""
  lblUrea(n).BackColor = &H8000000F
Next

lblCamp = ""
lblCamp.BackColor = &H8000000F
lblCampLatex = ""
lblCampLatex.BackColor = &H8000000F
txtGram = ""
txtCampCulture = ""

lblSalmonella = ""
lblShigella = ""
lblColindale = ""

lblPC0157 = ""
lblPC0157.BackColor = &H8000000F
lbl0157Latex = ""
lbl0157Latex.BackColor = &H8000000F
lbl0157 = ""

For n = 0 To 3
  lblEPC(n) = ""
  lblEPC(n).BackColor = &H8000000F
Next

For n = 0 To 2
  chkOccult(n) = False
  lblOccult(n) = ""
  lblOccult(n).BackColor = &H8000000F
Next

lblRota = ""
lblRota.BackColor = &H8000000F
lblAdeno = ""
lblAdeno.BackColor = &H8000000F

lblToxinAL = ""
lblToxinAL.BackColor = &H8000000F
lblToxinATA = ""
lblToxinATA.BackColor = &H8000000F

lblAus = ""
lblAus.BackColor = &H8000000F
For n = 0 To 2
  cmbOP(n) = ""
Next

End Sub
Private Sub FillCastsCrystalsMiscSite()

Dim n As Integer
Dim tb As Recordset
Dim sql As String

cmbCasts.Clear
cmbCrystals.Clear
cmbMisc(0).Clear
cmbMisc(1).Clear
cmbMisc(2).Clear
cmbSite.Clear
cmbClinDetails.Clear

For n = 1 To 8
  cmbQualifier(n).Clear
  cmbIdentification(n).Clear
Next

sql = "Select * from Lists where " & _
      "ListType = 'IN' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  For n = 1 To 8
    cmbIdentification(n).AddItem tb!Text & ""
  Next
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'CA' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbCasts.AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'CR' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbCrystals.AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'MI' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbMisc(0).AddItem tb!Text & ""
  cmbMisc(1).AddItem tb!Text & ""
  cmbMisc(2).AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'SI' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbSite.AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'OV' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbOP(0).AddItem tb!Text & ""
  cmbOP(1).AddItem tb!Text & ""
  cmbOP(2).AddItem tb!Text & ""
  cmbOva(0).AddItem tb!Text & ""
  cmbOva(1).AddItem tb!Text & ""
  cmbOva(2).AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'CD' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbClinDetails.AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'HO' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbHospital.AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'GS' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbGram(1).AddItem tb!Text & ""
  cmbGram(2).AddItem tb!Text & ""
  cmbGram(3).AddItem tb!Text & ""
  cmbGram(4).AddItem tb!Text & ""
  cmbGram(5).AddItem tb!Text & ""
  cmbGram(6).AddItem tb!Text & ""
  cmbGram(7).AddItem tb!Text & ""
  cmbGram(8).AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'WP' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbWetPrep(1).AddItem tb!Text & ""
  cmbWetPrep(2).AddItem tb!Text & ""
  cmbWetPrep(3).AddItem tb!Text & ""
  cmbWetPrep(4).AddItem tb!Text & ""
  cmbWetPrep(5).AddItem tb!Text & ""
  cmbWetPrep(6).AddItem tb!Text & ""
  cmbWetPrep(7).AddItem tb!Text & ""
  cmbWetPrep(8).AddItem tb!Text & ""
  tb.MoveNext
Loop

sql = "Select * from Lists where " & _
      "ListType = 'MQ' order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbQualifier(1).AddItem tb!Text & ""
  cmbQualifier(2).AddItem tb!Text & ""
  cmbQualifier(3).AddItem tb!Text & ""
  cmbQualifier(4).AddItem tb!Text & ""
  cmbQualifier(5).AddItem tb!Text & ""
  cmbQualifier(6).AddItem tb!Text & ""
  cmbQualifier(7).AddItem tb!Text & ""
  cmbQualifier(8).AddItem tb!Text & ""
  tb.MoveNext
Loop

End Sub
Private Sub FillAbGrid(ByVal Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim ReportCounter As Integer
Dim n As Integer
Dim Y As Integer
Dim Found As Boolean

On Error GoTo ehla

With grdAB(Index)
  .Visible = False
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

ReportCounter = 0

sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
      "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
      "from ABDefinitions as D, Antibiotics as A where " & _
      "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
      "and D.Site = '" & cmbSite & "' " & _
      "and D.PriSec = 'P' " & _
      "and D.AntibioticName = A.AntibioticName " & _
      "order by D.ListOrder"
Set tb = New Recordset
RecOpenClient 0, tb, sql
If tb.EOF Then
  sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
        "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
        "from ABDefinitions as D, Antibiotics as A where " & _
        "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
        "and Site = 'Generic' " & _
        "and D.PriSec = 'P' " & _
        "and D.AntibioticName = A.AntibioticName " & _
        "order by D.ListOrder"
  Set tb = New Recordset
  RecOpenClient 0, tb, sql
  If tb.EOF Then
   ' iMsg "Site/Organism not defined.", vbCritical
    Exit Sub
  End If
End If

Do While Not tb.EOF
  grdAB(Index).AddItem Trim$(tb!AntibioticName)
  grdAB(Index).Row = grdAB(Index).Rows - 1
  grdAB(Index).Col = 2
  If IsChild() And Not tb!AllowIfChild Then
    Set grdAB(Index).CellPicture = imgSquareCross.Picture
    grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "C"
  ElseIf IsPregnant() And Not tb!AllowIfPregnant Then
    Set grdAB(Index).CellPicture = imgSquareCross.Picture
    grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "P"
  ElseIf IsOutPatient() And Not tb!AllowIfOutPatient Then
    Set grdAB(Index).CellPicture = imgSquareCross.Picture
    grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "O"
  Else
    Set grdAB(Index).CellPicture = Me.Picture
  End If
  tb.MoveNext
Loop

For n = 0 To lstABsInUse.ListCount - 1
  If lstABsInUse.List(n) <> "Antibiotic Not Stated" And lstABsInUse.List(n) <> "None" Then
    Found = False
    For Y = 1 To grdAB(Index).Rows - 1
      If grdAB(Index).TextMatrix(Y, 0) = lstABsInUse.List(n) Then
        Found = True
        Exit For
      End If
    Next
    If Not Found Then
      grdAB(Index).AddItem lstABsInUse.List(n)
      grdAB(Index).Row = grdAB(Index).Rows - 1
    Else
      grdAB(Index).Row = Y
    End If
    grdAB(Index).Col = 2
    Set grdAB(Index).CellPicture = imgSquareTick.Picture
  End If
Next

If grdAB(Index).Rows > 2 Then
  grdAB(Index).RemoveItem 1
End If

Exit Sub

ehla:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

LogError "fEditMicrobiology/FillABGrig:" & Format(er) & ":" & ers
grdAB(Index).Visible = True
Exit Sub

End Sub

Private Sub CheckIfDoSensitivity()
'
'cSens = 0
'
'If Val(tProtein) > 0 Or tProtein = "Pos" Then
'  cSens = 1
'End If
'
'tWCC = Trim(tWCC)
'If tWCC <> "" And tWCC <> "1-10" And tWCC <> "Nil" Then
'  cSens = 1
'End If
'
'tRCC = Trim(tRCC)
'If tRCC <> "" And tRCC <> "1-10" And tRCC <> "10-50" And tRCC <> "Nil" Then
'  cSens = 1
'End If

End Sub

Function CheckConflict() As Boolean

Dim s As String
Dim Conflict As Boolean
Dim sn As Recordset
Dim tb As Recordset
Dim sql As String
Dim Organism(0 To 1) As String
Dim n As Integer
Dim OrgWas As String
Dim OrgIs As String
Dim ConflictList As String
Dim Grid As Integer
Dim Org As String
Dim ThisRunNumber As String
Dim SampleDate As String

On Error GoTo ehcc

If Trim(txtChart) = "" Then
  CheckConflict = False
  Exit Function
End If

sql = "select top 1 * from demographics where " & _
      "ForUrine = 1 and " & _
      "chart = '" & txtChart & "' and " & _
      "sampledatetime < '" & Format(dtSampleDate, "dd/mmm/yyyy") & "' and " & _
      "sampledatetime > '" & Format(DateAdd("d", -15, dtSampleDate), "dd/mmm/yyyy") & "' " & _
      "order by sampledatetime desc"

Set sn = New Recordset
RecOpenServer 0, sn, sql
If sn.EOF Then
  CheckConflict = False
  Exit Function
End If

ThisRunNumber = sn!RunNumber
SampleDate = Format(sn!GlobalSampleDateTime, "dd/mmm/yyyy")

sql = "Select * from urine where " & _
      "runnumber = '" & ThisRunNumber & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  CheckConflict = False
  Exit Function
End If
Organism(0) = tb!cult0 & ""
Organism(1) = tb!cult1 & ""
If Trim(Organism(0) & Organism(1) = "") Then
  CheckConflict = False
  Exit Function
End If
If cmbOrgGroup(0) <> Organism(0) And _
   cmbOrgGroup(0) <> Organism(1) And _
   cmbOrgGroup(1) <> Organism(0) And _
   cmbOrgGroup(1) <> Organism(1) Then
  CheckConflict = False
  Exit Function
End If

Conflict = False

If cmbOrgGroup(0) = Organism(0) Then
  Grid = 0: Org = cmbOrgGroup(0): GoSub SensCheck
End If
If cmbOrgGroup(1) = Organism(0) Then
  Grid = 1: Org = cmbOrgGroup(0): GoSub SensCheck
End If
If cmbOrgGroup(0) = Organism(1) Then
  Grid = 0: Org = cmbOrgGroup(1): GoSub SensCheck
End If
If cmbOrgGroup(1) = Organism(1) Then
  Grid = 1: Org = cmbOrgGroup(1): GoSub SensCheck
End If

If Conflict Then
  s = "Sensitivity Conflict" & vbCrLf & _
      "Sample Number " & ThisRunNumber & _
      " (" & Format(SampleDate, "dd/mm/yyyy") & ")" & vbCrLf & _
      ConflictList & _
      "Do you wish to procede?"
  If iMsg(s, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
    CheckConflict = False
  Else
    CheckConflict = True
  End If
End If

Exit Function

SensCheck:
With grdAB(Grid)
  For n = 0 To .Rows - 1
    .Col = 1
    .Row = n
    If .Text <> "" Then
      OrgIs = .Text
      .Col = 0
      sql = "Select * from sensitivities where " & _
            "Samplenumber = '" & ThisRunNumber & "' " & _
            "and Antibiotic = '" & .Text & "' " & _
            "and Organism = '" & Org & "'"
      Set tb = New Recordset
      RecOpenServer 0, tb, sql
      
      If Not tb.EOF Then
        OrgWas = tb!Result & ""
        If OrgWas <> "" And (OrgWas <> OrgIs) Then
          Conflict = True
          ConflictList = ConflictList & cmbOrgGroup(Grid) & " " & .Text & " was " & _
                    Switch(OrgWas = "S", "Sensitive", _
                    OrgWas = "R", "Resistant", _
                    OrgWas = "I", "Indeterminate") & vbCrLf
        End If
      End If
    End If
  Next
End With

Return

Exit Function

ehcc:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "fEditMicrobiology/CheckConflict:" & Format(er) & ":" & ers
Exit Function

End Function


Private Sub Form_Deactivate()

PBar = 0
TimerBar.Enabled = False

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

PBar = 0

End Sub

Private Sub Form_Load()

Dim n As Integer
Dim tb As Recordset
Dim sql As String

ssTab1.TabVisible(7) = False 'MRSA

FillLists
FillOrganisms
FillCurrentABs

FillMRU Me

With lblChartNumber
  .BackColor = &H8000000F
  .ForeColor = vbBlack
  Select Case HospName(0)
    Case "Cavan"
      .Caption = "Cavan Chart #"
    Case "Hogwarts"
      .Caption = "Hogwarts Chart #"
    Case "Monaghan"
      .Caption = "Monaghan Chart #"
  End Select
End With

dtRunDate = Format$(Now, "dd/mm/yyyy")
dtSampleDate = Format$(Now, "dd/mm/yyyy")

UpDown1.max = 9999999

txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")
GetSampleIDWithOffset
LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveMicro.Enabled = False

frDipStick.Visible = sysOptDipStick(0)
fraUrineSpecific.Visible = frmOptUrineSpecific

For n = 1 To 8
  With txtNotes(n)
    If sysOptUseFullID(0) Then
      .Visible = False
      .Top = 3000
      .Height = 2565
    Else
      .Visible = True
      .Top = 2580
      .Height = 2955
    End If
  End With
Next

'SQL = "Select MemberOf from Users where " & _
'      "Name = '" & AddTicks(UserName) & "' " & _
'      "and (MemberOf = 'Managers' " & _
'      "or MemberOf = 'Administrators')"
'Set tb = New Recordset
'RecOpenServer 0, tb, SQL
'cmdValidateMicro.Enabled = False
'If Not tb.EOF Then
  cmdValidateMicro.Enabled = True
'End If

Activated = False

End Sub
Private Sub LoadDemographics()

Dim sql As String
Dim tb As Recordset
Dim SampleDate As String
Dim RooH As Boolean
Dim n As Integer

On Error GoTo gld

RooH = IsRoutine()
cRooH(0) = RooH
cRooH(1) = Not RooH
bViewBB.Enabled = False
lstABsInUse.Clear
cmbSite = ""
txtSiteDetails = ""
lstABsInUse.Clear
lblABsInUse = ""

If Trim$(txtSampleID) = "" Then Exit Sub
  
sql = "Select * from MicroSiteDetails where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenClient 0, tb, sql
If Not tb.EOF Then
  cmbSite = tb!Site & ""
  txtSiteDetails = tb!SiteDetails & ""
  If tb!PCA0 & "" <> "" Then lstABsInUse.AddItem tb!PCA0 & ""
  If tb!PCA1 & "" <> "" Then lstABsInUse.AddItem tb!PCA1 & ""
  If tb!PCA2 & "" <> "" Then lstABsInUse.AddItem tb!PCA2 & ""
  If tb!PCA3 & "" <> "" Then lstABsInUse.AddItem tb!PCA3 & ""
End If
lblABsInUse = ""
For n = 0 To lstABsInUse.ListCount - 1
  lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
Next

sql = "Select * from Demographics where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
      
Set tb = New Recordset
RecOpenClient 0, tb, sql
If tb.EOF Then
  mNewRecord = True
  dtRecDate = Format$(Now, "dd/mm/yyyy")
  dtRunDate = Format$(Now, "dd/mm/yyyy")
  dtSampleDate = Format$(Now, "dd/mm/yyyy")
  txtChart = ""
  txtName = ""
  txtAddress(0) = ""
  txtAddress(1) = ""
  txtSex = ""
  txtDoB = ""
  txtAge = ""
  cmbWard = "GP"
  cmbClinician = ""
  cmbGP = ""
  cmbHospital = HospName(0)
  txtClinDetails = ""
  txtDemographicComment = ""
  tSampleTime.Mask = ""
  tSampleTime.Text = ""
  tSampleTime.Mask = "##:##"
  lblChartNumber.Caption = HospName(0) & " Chart #"
  lblChartNumber.BackColor = &H8000000F
  lblChartNumber.ForeColor = vbBlack
  chkPregnant = 0
Else
  If Trim$(tb!Hospital & "") <> "" Then
    cmbHospital = Trim$(tb!Hospital)
    lblChartNumber = Trim$(tb!Hospital) & " Chart #"
    If tb!Hospital = HospName(0) Then
      lblChartNumber.BackColor = &H8000000F
      lblChartNumber.ForeColor = vbBlack
    Else
      lblChartNumber.BackColor = vbRed
      lblChartNumber.ForeColor = vbYellow
    End If
  Else
    cmbHospital = HospName(0)
    lblChartNumber.Caption = HospName(0) & " Chart #"
    lblChartNumber.BackColor = &H8000000F
    lblChartNumber.ForeColor = vbBlack
  End If
  If IsDate(tb!SampleDate) Then
    dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
  Else
    dtSampleDate = Format$(Now, "dd/mm/yyyy")
  End If
  If IsDate(tb!Rundate) Then
    dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
  Else
    dtRunDate = Format$(Now, "dd/mm/yyyy")
  End If
  If IsDate(tb!RecDate) Then
    dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
  Else
    dtRecDate = dtRunDate
  End If
  mNewRecord = False
  cRooH(0) = tb!RooH
  cRooH(1) = Not tb!RooH
  txtChart = tb!Chart & ""
  txtName = tb!PatName & ""
  txtAddress(0) = tb!Addr0 & ""
  txtAddress(1) = tb!Addr1 & ""
  Select Case Left$(Trim$(UCase$(tb!Sex & "")), 1)
    Case "M": txtSex = "Male"
    Case "F": txtSex = "Female"
    Case Else: txtSex = ""
  End Select
  txtDoB = Format$(tb!DoB, "dd/mm/yyyy")
  txtAge = tb!Age & ""
  cmbWard = tb!Ward & ""
  cmbClinician = tb!Clinician & ""
  cmbGP = tb!GP & ""
  txtClinDetails = tb!ClDetails & ""
  If IsDate(tb!SampleDate) Then
    dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
    If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
      tSampleTime = Format$(tb!SampleDate, "hh:mm")
    Else
      tSampleTime.Mask = ""
      tSampleTime.Text = ""
      tSampleTime.Mask = "##:##"
    End If
  Else
    dtSampleDate = Format$(Now, "dd/mm/yyyy")
    tSampleTime.Mask = ""
    tSampleTime.Text = ""
    tSampleTime.Mask = "##:##"
  End If
  If IsNull(tb!Pregnant) Then
    chkPregnant = 0
  Else
    chkPregnant = IIf(tb!Pregnant, 1, 0)
  End If
End If
cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False

If sysOptBloodBank(0) Then
  If Trim$(txtChart) <> "" Then
    sql = "Select  * from PatientDetails where " & _
          "PatNum = '" & txtChart & "'"
    Set tb = New Recordset
    RecOpenClientBB 0, tb, sql
    bViewBB.Enabled = Not tb.EOF
  End If
End If

Screen.MousePointer = 0

Exit Sub

gld:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "fEditMicrobiology/LoadDemographics:" & Str(er) & ":" & ers
Exit Sub

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Val(txtSampleID) > Val(GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")) Then
  SaveSetting "NetAcquire", "StartUp", "LastUsedMicro", txtSampleID
End If

pPrintToPrinter = ""

Activated = False

End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub


Private Sub Frame2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub Frame7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub fraUrineSpecific_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub

Private Sub irelevant_Click(Index As Integer)

Dim sql As String
Dim tb As Recordset
Dim strDept As String
Dim strDirection As String

On Error GoTo ehic

strDept = "Urine"

If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
  If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
    GetSampleIDWithOffset
    SaveDemographics
    cmdSaveDemographics.Enabled = False
    cmdSaveInc.Enabled = False
  End If
End If

Select Case ssTab1.Tab
  Case 0:
    If cmbSite = "Urine" Then
      strDept = "Urine"
    ElseIf cmbSite = "Faeces" Then
      strDept = "Faeces"
    Else
      If Index = 0 Then
        txtSampleID = Format$(Val(txtSampleID) - 1)
      Else
        txtSampleID = Format$(Val(txtSampleID) + 1)
      End If
      GetSampleIDWithOffset
      LoadAllDetails
  
      cmdSaveDemographics.Enabled = False
      cmdSaveInc.Enabled = False
      cmdSaveMicro.Enabled = False
      Exit Sub
    End If
  Case 1: strDept = "Urine"
  Case 2: strDept = "UrineIdent"
  Case 3: strDept = "Faeces"
  Case 4:
    If ssTab1.TabVisible(1) Then
      strDept = "Urine"
    ElseIf ssTab1.TabVisible(3) Then
      strDept = "Faeces"
    Else
      If Index = 0 Then
        txtSampleID = Format$(Val(txtSampleID) - 1)
      Else
        txtSampleID = Format$(Val(txtSampleID) + 1)
      End If
      GetSampleIDWithOffset
      LoadAllDetails
  
      cmdSaveDemographics.Enabled = False
      cmdSaveInc.Enabled = False
      cmdSaveMicro.Enabled = False
      Exit Sub
    End If
End Select

strDirection = IIf(Index = 0, "<", ">")
GetSampleIDWithOffset

sql = "Select top 1 SampleID from " & strDept & " where " & _
      "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
      "Order by SampleID " & IIf(strDirection = "<", "Desc", "Asc")

Set tb = New Recordset
RecOpenClient 0, tb, sql
If Not tb.EOF Then
  txtSampleID = Val(tb!SampleID & "") - sysOptMicroOffset(0)
End If

GetSampleIDWithOffset
LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveMicro.Enabled = False

Exit Sub

ehic:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "fEditMicrobiology/iRelevant()_Click:" & Str(er) & ":" & ers
Exit Sub

End Sub

Private Sub iRunDate_Click(Index As Integer)

If Index = 0 Then
  dtRunDate = DateAdd("d", -1, dtRunDate)
Else
  If DateDiff("d", dtRunDate, Now) > 0 Then
    dtRunDate = DateAdd("d", 1, dtRunDate)
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

End Sub

Private Sub iSampleDate_Click(Index As Integer)

If Index = 0 Then
  dtSampleDate = DateAdd("d", -1, dtSampleDate)
Else
  If DateDiff("d", dtSampleDate, Now) > 0 Then
    dtSampleDate = DateAdd("d", 1, dtSampleDate)
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

End Sub


Private Sub iToday_Click(Index As Integer)

If Index = 0 Then
  dtRunDate = Format$(Now, "dd/mm/yyyy")
ElseIf Index = 1 Then
  If DateDiff("d", dtRunDate, Now) > 0 Then
    dtSampleDate = dtRunDate
  Else
    dtSampleDate = Format$(Now, "dd/mm/yyyy")
  End If
Else
  dtRecDate = Format$(Now, "dd/mm/yyyy")
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

End Sub


Private Sub lbl0157_Click()

If lbl0157 = "" Then
  lbl0157 = "E Coli 0157 not isolated"
ElseIf InStr(lbl0157, "not") Then
  lbl0157 = "E Coli 0157 isolated"
Else
  lbl0157 = ""
End If

cmdSaveMicro.Enabled = True

End Sub

Private Sub lbl0157Latex_Click()

With lbl0157Latex
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
    lbl0157 = "E Coli 0157 Not Isolated"
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
    lbl0157 = ""
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
    lbl0157 = ""
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblAdeno_Click()

SetPositiveNegative lblAdeno

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblCamp_Click()

With lblCamp
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
     txtCampCulture = "Not Isolated"
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
    txtCampCulture = ""
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
    txtCampCulture = ""
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblEPC_Click(Index As Integer)

With lblEPC(Index)
  Select Case .Caption
  Case ""
    .Caption = "Neg"
    .BackColor = vbGreen
  Case "Neg"
    .Caption = "Pos"
    .BackColor = vbRed
  Case "Pos"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblLact_Click(Index As Integer)

With lblLact(Index)
  Select Case .Caption
  Case ""
    .Caption = "Neg"
    .BackColor = vbGreen
  Case "Neg"
    .Caption = "Pos"
    .BackColor = vbRed
  Case "Pos"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblChartNumber_Click()

With lblChartNumber
  .BackColor = &H8000000F
  .ForeColor = vbBlack
  If .Caption = "Monaghan Chart #" Then
    .Caption = "Cavan Chart #"
    cmbHospital = "Cavan"
    cmbHospital_Click
    If HospName(0) = "Monaghan" Then
      .BackColor = vbRed
      .ForeColor = vbYellow
    End If
  ElseIf .Caption = "Cavan Chart #" Then
    .Caption = "Monaghan Chart #"
    cmbHospital = "Monaghan"
    cmbHospital_Click
    If HospName(0) = "Cavan" Then
      .BackColor = vbRed
      .ForeColor = vbYellow
    End If
  End If
End With

If Trim$(txtChart) <> "" Then
  LoadPatientFromChart Me, mNewRecord
  cmdSaveDemographics.Enabled = True
  cmdSaveInc.Enabled = True
End If

End Sub

Private Sub lblOccult_Click(Index As Integer)

With lblOccult(Index)
  Select Case .Caption
  Case ""
    .Caption = "Neg"
    .BackColor = vbGreen
  Case "Neg"
    .Caption = "Pos"
    .BackColor = vbRed
  Case "Pos"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblPC_Click()

With lblPC
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblPC0157_Click()

With lblPC0157
  Select Case Trim$(.Caption)
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
    lbl0157 = "E Coli 0157 Not Isolated"
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
    lbl0157 = ""
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
    lbl0157 = ""
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblRota_Click()

SetPositiveNegative lblRota

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblSelenite_Click()

With lblSelenite
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

End Sub


Private Sub lblToxinATA_Click()

SetPositiveNegative lblToxinATA

cmdSaveMicro.Enabled = True

End Sub


Private Sub lblToxinB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With lblToxinB
  Select Case .Caption
  Case ""
    .Caption = "Not Detected"
    .BackColor = vbGreen
  Case "Not Detected"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = "Inconclusive"
    .BackColor = vbYellow
  Case "Inconclusive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
  lblToxinA.Caption = .Caption
  lblToxinA.BackColor = .BackColor
End With

cmdSaveMicro.Enabled = True

End Sub


Private Sub lblUrea_Click(Index As Integer)

With lblUrea(Index)
  Select Case .Caption
  Case ""
    .Caption = "Neg"
    .BackColor = vbGreen
  Case "Neg"
    .Caption = "Pos"
    .BackColor = vbRed
  Case "Pos"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lblCampLatex_Click()

With lblCampLatex
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
     txtCampCulture = "Not Isolated"
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
    txtCampCulture = ""
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
    txtCampCulture = ""
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub lstABsInUse_Click()

Dim n As Integer

lstABsInUse.RemoveItem lstABsInUse.ListIndex

lblABsInUse = ""
For n = 0 To lstABsInUse.ListCount - 1
  lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
Next

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case PreviousTab
  Case 0
    If cmdSaveDemographics.Enabled Then
      If iMsg("Demographic Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveDemographics_Click
      End If
    End If
  Case 1
    If cmdSaveMicro.Enabled Then
      If iMsg("Urine Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveUrine False
      End If
    End If
  Case 2, 5
    If cmdSaveMicro.Enabled Then
      If iMsg("Identification Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveIdent False
      End If
    End If
  Case 3
    If cmdSaveMicro.Enabled Then
      If iMsg("Faeces Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveFaeces
      End If
    End If
  Case 4, 6
    If cmdSaveMicro.Enabled Then
      If iMsg("Culture/Sensitivity Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveIsolates
        SaveSensitivities gNO
      End If
    End If
  Case 7: 'MRSA
  Case 8: 'FOB
    If cmdSaveMicro.Enabled Then
      If iMsg("Faecal Occult Blood Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveFOB
      End If
    End If
  Case 9: 'Rota/Adeno
    If cmdSaveMicro.Enabled Then
      If iMsg("Rota/Adeno Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveRotaAdeno
      End If
    End If
  Case 10: 'C.diff
    If cmdSaveMicro.Enabled Then
      If iMsg("C.diff Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveCdiff
      End If
    End If
  Case 11: 'O/P
    If cmdSaveMicro.Enabled Then
      If iMsg("Ova/Parasites Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveOP
      End If
    End If
  Case 12, 13
    If cmdSaveMicro.Enabled Then
      If iMsg("Identification Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        SaveIdentification False
      End If
    End If
End Select

GetSampleIDWithOffset

Select Case ssTab1.Tab
  Case 0: 'Demographics

  Case 1: 'Urine
    If Not UrineLoaded Then
      LoadUrine
      UrineLoaded = True
    End If

  Case 2, 5: 'Identification
    If Not IdentLoaded Then
      LoadIdent
      IdentLoaded = True
    End If

  Case 3: 'Faeces
    If Not FaecesLoaded Then
      LoadFaeces
      FaecesLoaded = True
    End If

  Case 4, 6: 'Sensitivities
    If Not CSLoaded Then
      LoadSensitivities
      CSLoaded = True
    End If
  Case 8: 'FOB
    If Not FOBLoaded Then
      LoadFOB
      FOBLoaded = True
    End If
  Case 9: 'Rota/Adeno
    If Not RotaAdenoLoaded Then
      LoadRotaAdeno
      RotaAdenoLoaded = True
    End If
  Case 10: 'cdiff
    If Not CdiffLoaded Then
      LoadCDiff
      CdiffLoaded = True
    End If
  Case 11: 'OP
    If Not OPLoaded Then
      LoadOP
      OPLoaded = True
    End If
  Case 12, 13: 'Identification
    If Not IdentificationLoaded Then
      LoadIdentification
      IdentificationLoaded = True
    End If
    
End Select

cmdSaveMicro.Enabled = False

SetViewHistory

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

End Sub


Private Sub txtaddress_Change(Index As Integer)

lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub txtaddress_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtaddress_LostFocus(Index As Integer)

txtAddress(Index) = Initial2Upper(txtAddress(Index))

End Sub


Private Sub txtAdeno_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub txtAdeno_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With txtAdeno
  Select Case .Text
  Case ""
    .Text = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Text = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Text = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtage_Change()

lblAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtAPI1_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtAPI2_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtBacteria_KeyUp(KeyCode As Integer, Shift As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtBacteria_LostFocus()

Select Case UCase(txtBacteria)
  Case "0": txtBacteria = "Nil"
  Case "F": txtBacteria = "Occasional"
  Case "1": txtBacteria = "+"
  Case "2": txtBacteria = "++"
  Case "3": txtBacteria = "+++"
  Case "4": txtBacteria = "++++"
  Case Else: txtBacteria = ""
End Select

End Sub


Private Sub txtBenceJones_KeyUp(KeyCode As Integer, Shift As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtchart_Change()

lblChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtchart_LostFocus()

If Trim$(txtChart) = "" Then Exit Sub
If Trim$(txtName) <> "" Then Exit Sub

LoadPatientFromChart Me, mNewRecord

End Sub


Private Sub txtClinDetails_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtConsultantComment_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtConsultantComment_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim s As Variant
Dim n As Integer

If Trim$(txtConsultantComment(Index)) = "" Then Exit Sub

s = Split(txtConsultantComment(Index), " ")

For n = 0 To UBound(s)
  sql = "Select * from Lists where " & _
        "ListType = 'BA' " & _
        "and Code = '" & s(n) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
    s(n) = tb!Text & ""
  End If
Next

txtConsultantComment(0) = Join(s, " ")
txtConsultantComment(1) = txtCSComment(0)

End Sub


Private Sub txtCrystal_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtDoB_Change()

lblDoB = txtDoB

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)
 
cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtDoB_LostFocus()

txtDoB = Convert62Date(txtDoB, BACKWARD)
txtAge = CalcAge(txtDoB)

End Sub


Private Sub TimerBar_Timer()

PBar = PBar + 1
  
If PBar = PBar.max Then
  Unload Me
  Exit Sub
End If

End Sub


Private Sub txtFatGlobules_KeyPress(KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtGlucose_KeyPress(KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtHCGLevel_KeyPress(KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtIdentification_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtName_Change()

lblName = txtName

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtname_LostFocus()

Dim strName As String
Dim strSex As String

strName = txtName
strSex = txtSex

NameLostFocus strName, strSex

txtName = strName
txtSex = strSex

End Sub


Private Sub txtCatalase_Click(Index As Integer)

ClickMe txtCatalase(Index)

End Sub


Private Sub ClickMe(c As Control)

With c
  Select Case Trim(UCase(.Text))
    Case "": .Text = "Pending"
    Case "PENDING":
      Select Case .Tag
        Case "Coa": .Text = "Negative"
        Case "Cat": .Text = "Negative"
        Case "Oxi": .Text = "Negative"
        Case "Ure": .Text = "Negative"
        Case "Rei": .Text = "Done"
        Case "Uri": .Text = "Done"
        Case "Ext": .Text = "Done"
      End Select
    Case "DONE": .Text = ""
    Case "NEGATIVE": .Text = "Positive"
    Case "POSITIVE": .Text = ""
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtCoagulase_Click(Index As Integer)

ClickMe txtCoagulase(Index)

End Sub


Private Sub txtExtraSensitivity_Click(Index As Integer)

ClickMe txtExtraSensitivity(Index)

End Sub


Private Sub txtNotes_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtOxidase_Click(Index As Integer)

ClickMe txtOxidase(Index)

End Sub


Private Sub txtPregnancy_Click()

cmdSaveMicro.Enabled = True

Select Case Left$(txtPregnancy & " ", 1)
  Case " ":
    txtPregnancy = "Negative"
    txtHCGLevel = "<20"
    txtUrineComment = ""
  Case "N":
    txtPregnancy = "Positive"
    txtHCGLevel = ">=20"
    txtUrineComment = ""
  Case "P":
    If HospName(0) = "SIVH" Then
      txtPregnancy = ""
    Else
      txtPregnancy = "Equivocal"
      txtHCGLevel = ""
      txtUrineComment = "Please repeat specimen in 24-48 hours."
    End If
  Case "E":
    txtPregnancy = "Specimen Unsuitable"
    txtHCGLevel = ""
    txtUrineComment = "Specimen Unsuitable - Please repeat."
  Case "S":
    txtPregnancy = ""
    txtHCGLevel = ""
    txtUrineComment = ""
End Select

End Sub


Private Sub txtPregnancy_KeyPress(KeyAscii As Integer)
'
'Dim DoLevel As Boolean
'
'DoLevel = False
'If HospName(0) = "Cavan" Or _
'   HospName(0) = "Monaghan" Or _
'   HospName(0) = "Hogwarts" Then
'  DoLevel = True
'End If
'
'Select Case UCase(Chr(KeyAscii))
'  Case "N":
'    txtPregnancy = "Negative"
'    If DoLevel Then
'      txtHCGLevel = QueryTwo("<200", "<25", "HCG Level", False)
'    End If
'  Case "P":
'    txtPregnancy = "Positive"
'    If DoLevel Then
'      txtHCGLevel = QueryTwo(">=200", ">=25", "HCG Level", False)
'    End If
'  Case "E":
'    txtPregnancy = "Equivocal"
'    txtHCGLevel = ""
'  Case "U":
'    txtPregnancy = "Specimen Unsuitable"
'    txtHCGLevel = ""
'  Case Else:
'    txtPregnancy = ""
'    txtHCGLevel = ""
'End Select
'
'KeyAscii = 0
'
'cmdSaveMicro.Enabled = True

End Sub


Private Sub txtBilirubin_Click()

Select Case txtBilirubin
  Case "": txtBilirubin = "+"
  Case "+": txtBilirubin = "++"
  Case "++": txtBilirubin = "+++"
  Case "+++": txtBilirubin = "++++"
  Case "++++": txtBilirubin = "Nil"
  Case "Nil": txtBilirubin = ""
End Select

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtBilirubin_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case vbKey0, vbKeyNumpad0, vbKeyN: txtBilirubin = "Nil"
  Case vbKey1, vbKeyNumpad1: txtBilirubin = "+"
  Case vbKey2, vbKeyNumpad2: txtBilirubin = "++"
  Case vbKey3, vbKeyNumpad3: txtBilirubin = "+++"
  Case vbKey4, vbKeyNumpad4: txtBilirubin = "++++"
  Case Else: txtBilirubin = ""
End Select
KeyAscii = 0

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtBloodHb_Click()

Select Case txtBloodHb
  Case "": txtBloodHb = "+"
  Case "+": txtBloodHb = "++"
  Case "++": txtBloodHb = "+++"
  Case "+++": txtBloodHb = "++++"
  Case "++++": txtBloodHb = "Nil"
  Case "Nil": txtBloodHb = ""
End Select

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtBloodHb_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case vbKey0, vbKeyNumpad0, vbKeyN: txtBloodHb = "Nil"
  Case vbKey1, vbKeyNumpad1: txtBloodHb = "+"
  Case vbKey2, vbKeyNumpad2: txtBloodHb = "++"
  Case vbKey3, vbKeyNumpad3: txtBloodHb = "+++"
  Case vbKey4, vbKeyNumpad4: txtBloodHb = "++++"
  Case Else: txtBloodHb = ""
End Select
KeyAscii = 0

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtGlucose_Click()

If txtGlucose = "" Then
  txtGlucose = "Pos"
Else
  txtGlucose.SelStart = 0
  txtGlucose.SelLength = Len(txtGlucose)
End If

End Sub


Private Sub txtKetones_Click()

Select Case txtKetones
  Case "": txtKetones = "+"
  Case "+": txtKetones = "++"
  Case "++": txtKetones = "+++"
  Case "+++": txtKetones = "++++"
  Case "++++": txtKetones = "Nil"
  Case "Nil": txtKetones = ""
End Select

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtKetones_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case vbKey0, vbKeyNumpad0, vbKeyN: txtKetones = "Nil"
  Case vbKey1, vbKeyNumpad1: txtKetones = "+"
  Case vbKey2, vbKeyNumpad2: txtKetones = "++"
  Case vbKey3, vbKeyNumpad3: txtKetones = "+++"
  Case vbKey4, vbKeyNumpad4: txtKetones = "++++"
  Case Else: txtKetones = ""
End Select
KeyAscii = 0

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtpH_Click()

Select Case txtPh
  Case "": txtPh = "Acid"
  Case "Acid": txtPh = "Alkaline"
  Case "Alkaline": txtPh = "Neutral"
  
  Case "Neutral":
    If iMsg("Is Sample Unsuitable?", vbQuestion + vbYesNo) = vbYes Then
      txtPh = "Unsuitable"
      txtProtein = ""
      txtGlucose = ""
      txtKetones = ""
      txtUrobilinogen = ""
      txtBilirubin = ""
    Else
      txtPh = ""
    End If
    
  Case Else: txtPh = ""
End Select

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtPh_KeyPress(KeyAscii As Integer)

KeyAscii = 0

Select Case txtPh
  Case "": txtPh = "Acid"
  Case "Acid": txtPh = "Alkaline"
  Case "Alkaline": txtPh = "Neutral"
  Case "Neutral": txtPh = "Unsuitable"
  Case Else: txtPh = ""
End Select

End Sub


Private Sub txtProtein_Click()

If txtProtein = "" Then
  txtProtein = "Pos"
Else
  txtProtein.SelStart = 0
  txtProtein.SelLength = Len(txtProtein)
End If

CheckIfDoSensitivity

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtProtein_KeyPress(KeyAscii As Integer)

CheckIfDoSensitivity

End Sub


Private Sub txtRCC_KeyUp(KeyCode As Integer, Shift As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtRCC_LostFocus()

Select Case UCase(txtRCC)
  Case "0": txtRCC = "Nil"
  Case "F": txtRCC = "Occasional"
  Case "1": txtRCC = "+"
  Case "2": txtRCC = "++"
  Case "3": txtRCC = "+++"
  Case "4": txtRCC = "++++"
  Case Else: txtRCC = ""
End Select

End Sub


Private Sub txtReincubation_Click(Index As Integer)

ClickMe txtReincubation(Index)

End Sub


Private Sub txtRota_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub txtRota_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With txtRota
  Select Case .Text
  Case ""
    .Text = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Text = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Text = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtSampleID_GotFocus()

If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
  If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
    GetSampleIDWithOffset
    SaveDemographics
    cmdSaveDemographics.Enabled = False
    cmdSaveInc.Enabled = False
  End If
End If

End Sub

Private Sub txtSampleID_LostFocus()

txtSampleID = Format$(Val(txtSampleID))
If txtSampleID = 0 Then Exit Sub

GetSampleIDWithOffset

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveMicro.Enabled = False

End Sub

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtSex_Change()

lblSex = txtSex

End Sub

Private Sub txtsex_Click()

Select Case Trim$(txtSex)
  Case "": txtSex = "Male"
  Case "Male": txtSex = "Female"
  Case "Female": txtSex = ""
  Case Else: txtSex = ""
End Select

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtsex_KeyPress(KeyAscii As Integer)

KeyAscii = 0
txtsex_Click

End Sub


Private Sub txtSex_LostFocus()

SexLostFocus txtSex, txtName

End Sub


Private Sub txtCSComment_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtCSComment_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim s As Variant
Dim n As Integer

If Trim$(txtCSComment(Index)) = "" Then Exit Sub

s = Split(txtCSComment(Index), " ")

For n = 0 To UBound(s)
  sql = "Select * from Lists where " & _
        "ListType = 'BA' " & _
        "and Code = '" & s(n) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
    s(n) = tb!Text & ""
  End If
Next

txtCSComment(0) = Join(s, " ")
txtCSComment(1) = txtCSComment(0)
txtUrineComment = txtCSComment(0)

End Sub


Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtUrineComment_KeyPress(KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub

Private Sub txtUrineComment_LostFocus()

Dim tb As Recordset
Dim sql As String
Dim s As Variant
Dim n As Integer

If Trim$(txtUrineComment) = "" Then Exit Sub

s = Split(txtUrineComment, " ")

For n = 0 To UBound(s)
  sql = "Select * from Lists where " & _
        "ListType = 'HA' " & _
        "and Code = '" & s(n) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
    s(n) = tb!Text & ""
  End If
Next

txtUrineComment = Join(s, " ")
txtCSComment(0) = txtUrineComment
txtCSComment(1) = txtUrineComment

End Sub


Private Sub txtSG_KeyPress(KeyAscii As Integer)

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtSiteDetails_Change()

lblSiteDetails = cmbSite & " " & txtSiteDetails

End Sub

Private Sub txtSiteDetails_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub


Private Sub txtUrineSensitivity_Click(Index As Integer)

ClickMe txtUrineSensitivity(Index)

End Sub


Private Sub txtUrobilinogen_Click()

Select Case txtUrobilinogen
  Case "": txtUrobilinogen = "+"
  Case "+": txtUrobilinogen = "++"
  Case "++": txtUrobilinogen = "+++"
  Case "+++": txtUrobilinogen = "++++"
  Case "++++": txtUrobilinogen = "Nil"
  Case "Nil": txtUrobilinogen = ""
End Select

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtUrobilinogen_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case vbKey0, vbKeyNumpad0, vbKeyN: txtUrobilinogen = "Nil"
  Case vbKey1, vbKeyNumpad1: txtUrobilinogen = "+"
  Case vbKey2, vbKeyNumpad2: txtUrobilinogen = "++"
  Case vbKey3, vbKeyNumpad3: txtUrobilinogen = "+++"
  Case vbKey4, vbKeyNumpad4: txtUrobilinogen = "++++"
  Case Else: txtUrobilinogen = ""
End Select
KeyAscii = 0

cmdSaveMicro.Enabled = True

End Sub


Private Sub txtWCC_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
  Case "P", "p":
    txtWCC = "Packed"
    KeyAscii = 0
  Case ">", "G", "g":
    txtWCC = ">100"
    KeyAscii = 0
  Case "N", "n":
    txtWCC = "Nil"
    KeyAscii = 0
End Select

cmdSaveMicro.Enabled = True

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

Private Sub UpDown1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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


Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

PBar = 0

GetSampleIDWithOffset

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveMicro.Enabled = False

End Sub



Public Property Let PrintToPrinter(ByVal strNewValue As String)

pPrintToPrinter = strNewValue

End Property
Public Property Get PrintToPrinter() As String

PrintToPrinter = pPrintToPrinter

End Property

