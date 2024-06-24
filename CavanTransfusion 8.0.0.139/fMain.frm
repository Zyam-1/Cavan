VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Transfusion2000"
   ClientHeight    =   7365
   ClientLeft      =   2100
   ClientTop       =   2490
   ClientWidth     =   9765
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7365
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fMain"
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   7110
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar pBar 
      Height          =   30
      Left            =   4080
      TabIndex        =   38
      Top             =   5145
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   53
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame3 
      Height          =   525
      Left            =   0
      TabIndex        =   35
      Top             =   7440
      Visible         =   0   'False
      Width           =   9645
      Begin VB.CheckBox Check2 
         Caption         =   "Pending Orders"
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
         Left            =   5370
         TabIndex        =   37
         Top             =   150
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Process Orders"
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
         Left            =   7530
         TabIndex        =   36
         Top             =   150
         Width           =   1995
      End
      Begin VB.Label Label5 
         Caption         =   "Process Orders: 0"
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
         Left            =   3600
         TabIndex        =   0
         Top             =   120
         Width           =   2805
      End
      Begin VB.Label Label6 
         Caption         =   "Pending Orders: 0"
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
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   30
      TabIndex        =   30
      Top             =   7890
      Visible         =   0   'False
      Width           =   9645
      Begin VB.CheckBox chkProcess 
         Caption         =   "Process Orders"
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
         Left            =   7530
         TabIndex        =   34
         Top             =   150
         Width           =   1995
      End
      Begin VB.CheckBox chkPending 
         Caption         =   "Pending Orders"
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
         Left            =   5370
         TabIndex        =   33
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label lblProcess 
         Caption         =   "Process Orders: 0"
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
         Left            =   2670
         TabIndex        =   32
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblPending 
         Caption         =   "Pending Orders: 0"
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
         Left            =   60
         TabIndex        =   31
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.TextBox txtBarcode 
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CheckBox chkRefresh 
      Caption         =   "Auto Refresh Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3570
      TabIndex        =   27
      Top             =   8520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Timer timRefreshRecords 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3540
      Top             =   9420
   End
   Begin VB.CommandButton btnRefresh 
      Caption         =   "Refresh Records"
      Height          =   525
      Left            =   2010
      TabIndex        =   25
      Top             =   9420
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdPlaceOrder 
      Caption         =   "Place Order"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   24
      Top             =   8490
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   22
      Top             =   9300
      Width           =   1425
      Begin VB.Image imgUnSelect 
         Height          =   225
         Left            =   60
         Picture         =   "fMain.frx":08CA
         Top             =   150
         Width           =   210
      End
      Begin VB.Image imgSelect 
         Height          =   225
         Left            =   75
         Picture         =   "fMain.frx":0BA0
         Top             =   150
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label3 
         Caption         =   "Select All"
         Height          =   255
         Left            =   330
         TabIndex        =   23
         Top             =   165
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   1335
         Picture         =   "fMain.frx":0E76
         Top             =   450
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetail 
      Height          =   1875
      Left            =   9720
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   3307
      _Version        =   393216
      Cols            =   9
   End
   Begin VB.Timer timAnalyserHeartBeat 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5190
      Top             =   9540
   End
   Begin VB.Frame fraHeartBeat 
      Height          =   465
      Left            =   7920
      TabIndex        =   19
      Top             =   6600
      Width           =   1755
      Begin VB.Image imgNOTOK 
         Height          =   225
         Left            =   1335
         Picture         =   "fMain.frx":114C
         Top             =   450
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblBTCHeartBeat 
         Caption         =   "BTC Heartbeat"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1170
      End
      Begin VB.Image imgOK 
         Height          =   225
         Left            =   1335
         Picture         =   "fMain.frx":1422
         Top             =   150
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Timer tmrVersion 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4020
      Top             =   9420
   End
   Begin VB.Frame ShortCut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   90
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   9585
      Begin VB.CommandButton cmdAntibodies 
         Height          =   735
         Left            =   870
         Picture         =   "fMain.frx":16F8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Positive Antibody List"
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint5 
         Caption         =   "5"
         Height          =   735
         Left            =   8790
         Picture         =   "fMain.frx":1A02
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton cmdMatch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8070
         Picture         =   "fMain.frx":1E44
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bMove 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7350
         Picture         =   "fMain.frx":214E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bXmatch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   150
         Picture         =   "fMain.frx":2590
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bXMReport 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2310
         Picture         =   "fMain.frx":289A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Daily CrossMatch Report "
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bQC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4470
         Picture         =   "fMain.frx":2BA4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bPatHistory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5190
         Picture         =   "fMain.frx":2EAE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bOrder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6630
         Picture         =   "fMain.frx":32F0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bPatSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3030
         Picture         =   "fMain.frx":35FA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bAntiD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3750
         Picture         =   "fMain.frx":3904
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bXMQuery 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1590
         Picture         =   "fMain.frx":3C0E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Search For CrossMatch"
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton bGoodsIn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5910
         Picture         =   "fMain.frx":3F18
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   165
         Width           =   705
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4650
      Top             =   9090
   End
   Begin MSFlexGridLib.MSFlexGrid flxProduct 
      Height          =   2655
      Left            =   8400
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   12
   End
   Begin VB.Label Label4 
      Caption         =   "Barcode:"
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
      Left            =   210
      TabIndex        =   29
      Top             =   8550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "fMain.frx":4222
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "fMain.frx":44F8
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgLockl 
      Height          =   480
      Left            =   9210
      Picture         =   "fMain.frx":47CE
      Top             =   1230
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTest 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "   Caution Test System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   9615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "tmrVersion"
      Height          =   195
      Left            =   3810
      TabIndex        =   16
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "imgMisc"
      Height          =   195
      Left            =   1380
      TabIndex        =   15
      Top             =   9120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgUnlock 
      Height          =   480
      Left            =   8700
      Picture         =   "fMain.frx":5098
      Top             =   1230
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnulogon 
         Caption         =   "&Log On"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log &Off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuIncidentLog 
         Caption         =   "&Incident Log"
         Enabled         =   0   'False
         Begin VB.Menu mnuAddIncident 
            Caption         =   "&Add"
         End
         Begin VB.Menu mnuViewIncident 
            Caption         =   "&View"
         End
      End
      Begin VB.Menu mnuResetLastUsed 
         Caption         =   "&Reset 'Last Used'"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAmendStatus 
         Caption         =   "&Amend Status"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBarCodes 
         Caption         =   "Co&des"
         Enabled         =   0   'False
         Begin VB.Menu mnuCodeCancel 
            Caption         =   "&Cancel"
         End
         Begin VB.Menu mnuCodeValidate 
            Caption         =   "&Validate"
         End
      End
      Begin VB.Menu mnuUnlockReasons 
         Caption         =   "&Unlock and OverRide Events"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMergeClinsWards 
         Caption         =   "Manage Clinicians/&Wards"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVersionControl 
         Caption         =   "&Version Control"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPrinterSelection 
         Caption         =   "&Printer Selection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuErrorLog 
         Caption         =   "&Error Log"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password"
         Enabled         =   0   'False
      End
      Begin VB.Menu mNull0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuBatches 
      Caption         =   "&Batches"
      Enabled         =   0   'False
      Begin VB.Menu mnuBatchEntry 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuBatchStock 
         Caption         =   "&Stock"
      End
      Begin VB.Menu mnuBatchHistory 
         Caption         =   "&History"
      End
      Begin VB.Menu mnuBatchAmendment 
         Caption         =   "Amendment"
      End
      Begin VB.Menu mnuBatchMovement 
         Caption         =   "&Movement"
      End
      Begin VB.Menu mnuNull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBatchHistoricNull 
         Caption         =   "&Historic"
         Begin VB.Menu mnuBatchIssueOld 
            Caption         =   "&Issue"
         End
         Begin VB.Menu mnuBatchStockOld 
            Caption         =   "&Stock"
         End
         Begin VB.Menu mnuBatchHistoryOld 
            Caption         =   "&History"
         End
         Begin VB.Menu mnuBatchAmendmentOld 
            Caption         =   "&Amendment"
         End
         Begin VB.Menu mnuBatchMovementOld 
            Caption         =   "&Movement"
         End
      End
   End
   Begin VB.Menu mnuUnit 
      Caption         =   "&Unit"
      Enabled         =   0   'False
      Begin VB.Menu mtransfuse 
         Caption         =   "&Transfuse/Destroy"
      End
      Begin VB.Menu mxm 
         Caption         =   "&X-Match"
         Begin VB.Menu mxmworklist 
            Caption         =   "&Worklist"
         End
         Begin VB.Menu mnuXMatch 
            Caption         =   "&New X-Match"
         End
         Begin VB.Menu mxmdaily 
            Caption         =   "&Daily Report"
         End
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuMoveCodabar 
         Caption         =   "&Move Codabar"
      End
      Begin VB.Menu mnuPrintReclaimed 
         Caption         =   "&Print Reclaimed Forms"
      End
      Begin VB.Menu mnuUnitFating 
         Caption         =   "&Fating"
      End
      Begin VB.Menu mnuConfirmGroup 
         Caption         =   "&Confirm Group"
      End
   End
   Begin VB.Menu mnuPatient 
      Caption         =   "&Patient"
      Enabled         =   0   'False
      Begin VB.Menu manatal 
         Caption         =   "&Ante-Natal"
      End
      Begin VB.Menu mdat 
         Caption         =   "&D.A.T."
      End
      Begin VB.Menu mghold 
         Caption         =   "&Group && Hold"
      End
      Begin VB.Menu mnuXMatchMain 
         Caption         =   "&X-Match"
         Begin VB.Menu mnupatnew 
            Caption         =   "&X-Match"
         End
         Begin VB.Menu mnuKleihauer 
            Caption         =   "&Kleihauer"
         End
      End
      Begin VB.Menu mdr 
         Caption         =   "&Daily Report"
         Begin VB.Menu mdaily 
            Caption         =   "&Group && Hold"
            Index           =   1
         End
         Begin VB.Menu mdaily 
            Caption         =   "&Antenatal"
            Index           =   2
         End
         Begin VB.Menu mdaily 
            Caption         =   "&D.A.T."
            Index           =   3
         End
         Begin VB.Menu mxmd2 
            Caption         =   "&Cross Match"
         End
         Begin VB.Menu malldaily 
            Caption         =   "&All"
         End
         Begin VB.Menu mpaf 
            Caption         =   "&Print All Forms"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mworklists 
         Caption         =   "&Worklists"
      End
      Begin VB.Menu mnupatsearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mreaction 
         Caption         =   "&Reaction"
      End
      Begin VB.Menu mchangehistory 
         Caption         =   "&Change History"
      End
   End
   Begin VB.Menu mnuEnquiry 
      Caption         =   "&Enquiry"
      Enabled         =   0   'False
      Begin VB.Menu mnuAntibodyList 
         Caption         =   "Po&sitive Antibody List"
      End
      Begin VB.Menu mpatenq 
         Caption         =   "&Patient Enquiry"
      End
      Begin VB.Menu mFreeStock 
         Caption         =   "&Free Stock"
      End
      Begin VB.Menu mnutransfused 
         Caption         =   "&Transfused"
      End
      Begin VB.Menu mnuTransferred 
         Caption         =   "Trans&ferred"
      End
      Begin VB.Menu mnuxmatched 
         Caption         =   "&X-Matched"
      End
      Begin VB.Menu mnureceived 
         Caption         =   "&Received"
      End
      Begin VB.Menu mnureturned 
         Caption         =   "Re&turned"
      End
      Begin VB.Menu mnurestocked 
         Caption         =   "Restoc&ked"
      End
      Begin VB.Menu mnudestroyed 
         Caption         =   "&Destroyed"
      End
      Begin VB.Menu mpathistory 
         Caption         =   "&History"
      End
      Begin VB.Menu mhistory 
         Caption         =   "&Unit History"
      End
      Begin VB.Menu mnuCardValidation 
         Caption         =   "Card Validation"
      End
      Begin VB.Menu mnuGroupChecked 
         Caption         =   "Group Checked"
      End
      Begin VB.Menu mnuUnitsByOrderNumber 
         Caption         =   "Units By Order Number"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCourier 
         Caption         =   "Courier"
         Begin VB.Menu mnuCourierViewLog 
            Caption         =   "&View Message Log"
         End
         Begin VB.Menu mnuCourierUnitHistory 
            Caption         =   "Unit &History"
         End
      End
   End
   Begin VB.Menu mnulists 
      Caption         =   "&Lists"
      Enabled         =   0   'False
      Begin VB.Menu mnucliniccondx 
         Caption         =   "&Clinical Conditions"
      End
      Begin VB.Menu mnusurgical 
         Caption         =   "&Surgical Procedures"
      End
      Begin VB.Menu mnuphysicians 
         Caption         =   "&Physicians"
      End
      Begin VB.Menu mListGPs 
         Caption         =   "&GP's"
      End
      Begin VB.Menu mnuwards 
         Caption         =   "&Wards"
      End
      Begin VB.Menu mListHospitals 
         Caption         =   "&Hospitals"
      End
      Begin VB.Menu msuppliers 
         Caption         =   "Supp&liers"
      End
      Begin VB.Menu msprod 
         Caption         =   "Special P&roducts"
      End
      Begin VB.Menu mnuantibodies 
         Caption         =   "&Cell Panel"
      End
      Begin VB.Menu mantigens 
         Caption         =   "A&ntigens"
      End
      Begin VB.Menu mgenefreq 
         Caption         =   "&Gene Frequencies"
      End
      Begin VB.Menu mListBatches 
         Caption         =   "&Batched Products"
      End
      Begin VB.Menu mPrinters 
         Caption         =   "P&rinters"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "So&und"
      End
      Begin VB.Menu mnuXMComments 
         Caption         =   "Co&mments"
      End
   End
   Begin VB.Menu mnuQC 
      Caption         =   "&Q.C."
      Enabled         =   0   'False
      Begin VB.Menu mnutemperatures 
         Caption         =   "&Temperature Control"
      End
      Begin VB.Menu mnudaily 
         Caption         =   "&ABO-Rh Daily"
      End
      Begin VB.Menu mnureagents 
         Caption         =   "&Reagents"
      End
      Begin VB.Menu mservice 
         Caption         =   "&Service"
      End
      Begin VB.Menu mnuViewQC 
         Caption         =   "&View"
         Begin VB.Menu mnuStLukesViewQC 
            Caption         =   "&Centrifuge && Temperatures"
            Index           =   0
         End
         Begin VB.Menu mnuStLukesViewQC 
            Caption         =   "&AHG"
            Index           =   1
         End
         Begin VB.Menu mnuStLukesViewQC 
            Caption         =   "&Grouping Cards"
            Index           =   2
         End
         Begin VB.Menu mnuStLukesViewQC 
            Caption         =   "&Patient Phenotype"
            Index           =   3
         End
         Begin VB.Menu mnuABO_AntiD_Control_History 
            Caption         =   "ABO-Anti-D Control History"
         End
         Begin VB.Menu mnuKleihauerQC 
            Caption         =   "&Kleihauer"
         End
      End
   End
   Begin VB.Menu mnustatistics 
      Caption         =   "&Statistics"
      Enabled         =   0   'False
      Begin VB.Menu mnuPatListByProduct 
         Caption         =   "&Products Issued or Transfused"
      End
      Begin VB.Menu mnuti 
         Caption         =   "&TI"
      End
      Begin VB.Menu mnuantid 
         Caption         =   "&Anti-D usage"
      End
      Begin VB.Menu musage 
         Caption         =   "&C/T Ratio"
         Begin VB.Menu mby 
            Caption         =   "By &Procedure"
            Index           =   0
         End
         Begin VB.Menu mby 
            Caption         =   "By &Condition"
            Index           =   1
         End
         Begin VB.Menu mby 
            Caption         =   "By C&linician"
            Index           =   2
         End
         Begin VB.Menu mby 
            Caption         =   "By &Ward"
            Index           =   3
         End
      End
      Begin VB.Menu mtotals 
         Caption         =   "&Totals"
      End
      Begin VB.Menu mnuMonthly 
         Caption         =   "&Monthly"
         Begin VB.Menu mmonthly 
            Caption         =   "&Monthly Totals"
         End
         Begin VB.Menu mnuMonthlyWastage 
            Caption         =   "&Wastage"
         End
      End
      Begin VB.Menu mnuYearlyStats 
         Caption         =   "&Yearly"
      End
   End
   Begin VB.Menu mnuStock 
      Caption         =   "Stoc&k"
      Enabled         =   0   'False
      Begin VB.Menu morder 
         Caption         =   "&Order"
      End
      Begin VB.Menu menternewstock 
         Caption         =   "&Enter New"
      End
      Begin VB.Menu mListEither 
         Caption         =   "&List"
         Begin VB.Menu mStockList 
            Caption         =   "&Single Products"
         End
         Begin VB.Menu mStockListBatches 
            Caption         =   "&Batched Products"
         End
      End
      Begin VB.Menu mProductList 
         Caption         =   "&Product List"
      End
      Begin VB.Menu mCurrentStock 
         Caption         =   "&Current Stock"
      End
      Begin VB.Menu mnuBatchRestock 
         Caption         =   "&Batch RCC Restock"
      End
      Begin VB.Menu mnuStockReconciliation 
         Caption         =   "Stock &Reconciliation"
      End
   End
   Begin VB.Menu mnuAudit 
      Caption         =   "&Audit"
      Enabled         =   0   'False
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help"
      Begin VB.Menu mhelpcontents 
         Caption         =   "&Contents"
         Visible         =   0   'False
      End
      Begin VB.Menu mhelponhelp 
         Caption         =   "&Help on Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mhelpobtainingmore 
         Caption         =   "&Obtaining Technical Support"
      End
      Begin VB.Menu mNull 
         Caption         =   "-"
      End
      Begin VB.Menu mhelpabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ClosingFlag As Boolean

Private Const fcsLine_NO = 0
Private Const fcsRID = 1
Private Const fcsSID = 2
Private Const fcsTCode = 3
Private Const fcsPID = 4
Private Const fcsSDate = 5
Private Const fcsBit = 6
Private Const fcsR = 7
Private Const fcsSta = 8

Private Const fcLine_NO = 0
Private Const fcRID = 1
Private Const fcSID = 2
Private Const fcPName = 3
Private Const fcTCode = 4
Private Const fcUnits = 5
Private Const fcSDate = 6
Private Const fcRDate = 7
Private Const fcBit = 8
Private Const fcSta = 9
Private Const fcDel = 10
Private Const fcUID = 11

Private Sub FormatGrid()
    On Error GoTo ERROR_FormatGrid
    
    flxDetail.Rows = 1
    flxDetail.row = 0
    
    flxDetail.ColWidth(fcsLine_NO) = 250
    
    flxDetail.TextMatrix(0, fcsRID) = "MRN"
    flxDetail.ColWidth(fcsRID) = 1500
    flxDetail.ColAlignment(fcsRID) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsSID) = "Sample ID"
    flxDetail.ColWidth(fcsSID) = 1500
    flxDetail.ColAlignment(fcsSID) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsTCode) = "Test Code"
    flxDetail.ColWidth(fcsTCode) = 1500
    flxDetail.ColAlignment(fcsTCode) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsPID) = "Profile ID"
    flxDetail.ColWidth(fcsPID) = 1500
    flxDetail.ColAlignment(fcsPID) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsSDate) = "Sample Date"
    flxDetail.ColWidth(fcsSDate) = 1550
    flxDetail.ColAlignment(fcsSDate) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsBit) = ""
    flxDetail.ColWidth(fcsBit) = 0
    flxDetail.ColAlignment(fcsBit) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsR) = ""
    flxDetail.ColWidth(fcsR) = 0
    flxDetail.ColAlignment(fcsR) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsSta) = "Status"
    flxDetail.ColWidth(fcsSta) = 1500
    flxDetail.ColAlignment(fcsSta) = flexAlignLeftCenter
    
    
    
    
    
    flxProduct.Rows = 1
    flxProduct.row = 0
    
    flxProduct.ColWidth(fcsLine_NO) = 200
    
    flxProduct.TextMatrix(0, fcRID) = ""
    flxProduct.ColWidth(fcRID) = 0
    flxProduct.ColAlignment(fcRID) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcSID) = "Sample ID"
    flxProduct.ColWidth(fcSID) = 1100
    flxProduct.ColAlignment(fcSID) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcPName) = "Patient Name"
    flxProduct.ColWidth(fcPName) = 2500
    flxProduct.ColAlignment(fcPName) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcTCode) = "Test Code"
    flxProduct.ColWidth(fcTCode) = 1000
    flxProduct.ColAlignment(fcTCode) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcUnits) = "Units"
    flxProduct.ColWidth(fcUnits) = 1000
    flxProduct.ColAlignment(fcUnits) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcSDate) = "Sample Date"
    flxProduct.ColWidth(fcSDate) = 1100
    flxProduct.ColAlignment(fcSDate) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcRDate) = "Date Required"
    flxProduct.ColWidth(fcRDate) = 1200
    flxProduct.ColAlignment(fcRDate) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcBit) = ""
    flxProduct.ColWidth(fcBit) = 0
    flxProduct.ColAlignment(fcBit) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcSta) = "Status"
    flxProduct.ColWidth(fcSta) = 1000
    flxProduct.ColAlignment(fcSta) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcDel) = ""
    flxProduct.ColWidth(fcDel) = 250
    flxProduct.ColAlignment(fcDel) = flexAlignLeftCenter
    
    flxProduct.TextMatrix(0, fcUID) = ""
    flxProduct.ColWidth(fcUID) = 0
    flxProduct.ColAlignment(fcUID) = flexAlignLeftCenter
    
        
    Exit Sub
ERROR_FormatGrid:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "FormatGrid", intEL, strES
End Sub

Private Sub ShowDetail()
    On Error GoTo ERROR_ShowDetail
    
    Dim sql As String
    Dim tb As ADODB.Recordset
    Dim l_str As String
    
    flxDetail.Rows = 1
    flxDetail.row = 0
    sql = "Select D.*, R.Chart from ocmRequestDetails D Inner Join ocmRequest R On R.RequestID = D.RequestID Where D.SampleType = 'BL' And D.DepartmentID = 'BT' And IsNULL(D.Programmed,'0') = '0' And IsNULL(D.trans,'0') = '0' And D.transA is NULL Order By D.RequestID"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb Is Nothing Then
        If Not tb.EOF Then
            While Not tb.EOF
                l_str = "" & vbTab & tb!Chart & vbTab & tb!SampleID & vbTab & tb!TestCode & vbTab & tb!ProfileID & vbTab & tb!SampleDate & vbTab & "N" & vbTab & tb!RequestID
                flxDetail.AddItem (l_str)
                tb.MoveNext
            Wend
        End If
    End If
    Call DrawFlx
        
    Exit Sub
ERROR_ShowDetail:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "ShowDetail", intEL, strES
End Sub

Private Sub ShowProductDetail()
    On Error GoTo ERROR_ShowDetail
    
    Dim sql As String
    Dim tb As ADODB.Recordset
    Dim l_str As String
    Dim l_Pending As Integer
    Dim l_Process As Integer
    
    flxProduct.Rows = 1
    flxProduct.row = 0
    sql = "Select D.*, R.PatName from ocmRequestDetails D "
    sql = sql & " Inner Join ocmRequest R On D.RequestID = R.RequestID"
    If chkPending.Value = vbChecked And chkProcess.Value = vbChecked Then
        sql = sql & " Where (D.Status = 'Pending' OR D.Status = 'Process') And D.transA = '1' Order By D.RequestID Desc"
    ElseIf chkPending.Value = vbChecked And chkProcess.Value = vbUnchecked Then
        sql = sql & " Where (D.Status = 'Pending') And D.transA = '1' Order By D.RequestID Desc"
    ElseIf chkPending.Value = vbUnchecked And chkProcess.Value = vbChecked Then
        sql = sql & " Where (D.Status = 'Process') And D.transA = '1' Order By D.RequestID Desc"
    Else
        sql = sql & " Where (D.Status = 'Pending' OR D.Status = 'Process') And D.transA = '1' Order By D.RequestID Desc"
    End If
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb Is Nothing Then
        If Not tb.EOF Then
            While Not tb.EOF
                l_str = "" & vbTab & tb!RequestID & vbTab & tb!SampleID & vbTab & tb!PatName & vbTab & tb!TestCode & vbTab & tb!Units & vbTab & tb!SampleDate & vbTab & tb!daterequired & vbTab & "N" & vbTab & tb!Status & vbTab & "X" & vbTab & tb!UID
                flxProduct.AddItem (l_str)
                tb.MoveNext
            Wend
        End If
    End If
'    l_str = "" & vbTab & "123123" & vbTab & "321321" & vbTab & "ssdfsdf" & vbTab & "23423" & vbTab & "er" & vbTab & "23423423" & vbTab & "32423" & vbTab & "N" & vbTab & "Pending"
'    flxProduct.AddItem (l_str)
    Call DrawFlxProduct
        
    Exit Sub
ERROR_ShowDetail:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "ShowProductDetail", intEL, strES
End Sub

Private Sub DrawFlx()
On Error GoTo ErrorHandler

    Dim l_Count As Integer
    
    For l_Count = 1 To flxDetail.Rows - 1
        If flxDetail.TextMatrix(l_Count, fcsBit) = "N" Then
            flxDetail.row = l_Count
            flxDetail.col = fcsLine_NO
            Set flxDetail.CellPicture = imgRedCross.Picture
        ElseIf flxDetail.TextMatrix(l_Count, fcsBit) = "Y" Then
            flxDetail.row = l_Count
            flxDetail.col = fcsLine_NO
            Set flxDetail.CellPicture = imgGreenTick.Picture
        End If
    Next l_Count
    flxDetail.Redraw = True
        
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "DrawFlx", intEL, strES
End Sub

Private Sub DrawFlxProduct()
On Error GoTo ErrorHandler

    Dim l_Count As Integer
    Dim l_CountCol As Integer
    Dim l_Pending As Integer
    Dim l_Process As Integer
    
    l_Pending = 0
    l_Process = 0
    For l_Count = 1 To flxProduct.Rows - 1
        If flxProduct.TextMatrix(l_Count, fcBit) = "N" Then
            flxProduct.row = l_Count
            flxProduct.col = fcLine_NO
            Set flxProduct.CellPicture = imgRedCross.Picture
        ElseIf flxProduct.TextMatrix(l_Count, fcBit) = "Y" Then
            flxProduct.row = l_Count
            flxProduct.col = fcLine_NO
            Set flxProduct.CellPicture = imgGreenTick.Picture
        End If
        
        If Trim(flxProduct.TextMatrix(l_Count, fcSta)) = "Process" Then
            flxProduct.col = fcSta
            flxProduct.row = l_Count
            flxProduct.CellBackColor = &H80FF80
        End If

        If Trim(flxProduct.TextMatrix(l_Count, fcSta)) = "Pending" Then
            l_Pending = l_Pending + 1
        End If
        If Trim(flxProduct.TextMatrix(l_Count, fcSta)) = "Process" Then
            l_Process = l_Process + 1
        End If
    Next l_Count
    lblPending.Caption = "Pending Orders: " & l_Pending
    lblProcess.Caption = "Process Orders: " & l_Process
'    flxProduct.Redraw = True
        
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "DrawFlxProduct", intEL, strES
End Sub

Private Sub CheckAboutToExpire()

          Dim tb As Recordset
          Dim sql As String
          Dim Expiry As String
          Dim Interval As Long
          Static LastViewed As String

10        On Error GoTo CheckAboutToExpire_Error

20        If LastViewed = "" Then
30            LastViewed = "01/Jan/2000"
40        End If

50        If DateDiff("h", LastViewed, Now) > 3 Then

60            Interval = Val(sysOptTransfusionExpiry(0))
70            If Interval = 0 Then
80                sysOptTransfusionExpiry(0) = 7
90                Interval = 7
100           End If

110           Expiry = Format$(DateAdd("d", Interval, Now), "dd/mmm/yyyy")

120           sql = "Select * from Latest where " & _
                    "DateExpiry < '" & Expiry & "' " & _
                    "and DateExpiry >= '" & Format$(Now, "dd/mmm/yyyy") & "'"
130           Set tb = New Recordset
140           RecOpenServerBB 0, tb, sql
150           If Not tb.EOF Then
160               frmAboutToExpire.Show 1
170           End If
180           LastViewed = Format$(Now, "dd/mmm/yyyy hh:mm:ss")

190       End If

200       Exit Sub

CheckAboutToExpire_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmMain", "CheckAboutToExpire", intEL, strES, sql

End Sub

Private Sub EnableMenus(ByVal Enable As Boolean)

          Dim m As Control

10        On Error GoTo EnableMenus_Error

20        If Not IsIDE Then
30            CheckIfNewVersionAvailable
40        End If

50        If Not Enable Then
60            UserName = ""
70            UserCode = ""
80            UserInitials = ""
90            UserMemberOf = ""
100       End If

110       For Each m In Me.Controls
120           If TypeOf m Is Menu Then
130               If m.Caption <> "-" Then
140                   m.Enabled = Enable
150               End If
160           End If
170       Next
180       mnufile.Enabled = True
190       mnulogon.Enabled = True
200       mhelp.Enabled = True


210       mnuCourierUnitHistory.Enabled = False
220       mnuCourierViewLog.Enabled = False

230       mnuChangePassword.Enabled = Enable
240       mnuBatches.Enabled = Enable
250       mnuUnit.Enabled = Enable
260       mnulists.Enabled = Enable
270       mnulogon.Enabled = Not Enable
280       mnulogoff.Enabled = Enable
290       mnuPatient.Enabled = Enable
300       mnuStock.Enabled = Enable
310       mnuQC.Enabled = Enable
320       ShortCut.Visible = Enable
330       mnuAmendStatus.Enabled = False
340       mnuBarCodes.Enabled = False
350       mnuResetLastUsed.Enabled = Enable
360       mnuIncidentLog.Enabled = Enable
370       mnuEnquiry = Enable
380       mnuMergeClinsWards.Enabled = False
390       mnuVersionControl.Enabled = False
400       mnustatistics.Enabled = False
410       mnuErrorLog.Enabled = False
420       mnuUnlockReasons.Enabled = False
430       mnuPrinterSelection.Enabled = Enable
440       mnuAudit.Enabled = False
450       mnuCourierUnitHistory.Enabled = True
460       mnuCourierViewLog.Enabled = True
          cmdPlaceOrder.Enabled = Enable
          If Enable Then
            txtBarcode.Locked = False
          Else
            txtBarcode.Locked = True
          End If

470       Training = False
480       If UserName = "Training" Then
490           Training = True
500       End If

510       If UserMemberOf = "Managers" Then
520           mnustatistics.Enabled = True
530           mnulists.Enabled = True
540           mnuAmendStatus.Enabled = True
550           mnuBarCodes.Enabled = True
560           mnuMergeClinsWards.Enabled = True
570           mnuVersionControl.Enabled = True
580           mnuErrorLog.Enabled = True
590           mnuUnlockReasons.Enabled = True
600           mnuAudit.Enabled = True
610       End If

620       StatusBar.Panels(3).Text = UserName

630       'pBar.Visible = Enable

640       Exit Sub

EnableMenus_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frmMain", "EnableMenus", intEL, strES

End Sub

Private Sub GetLogOn()

          Dim X() As String
          Dim n As Integer

10        frmLogOn.Show 1

20        If Trim$(UserName) <> "" Then
30            X = Split(UserName, " ")
40            UserInitials = ""
50            For n = 0 To UBound(X)
60                UserInitials = UserInitials & UCase$(Left$(X(n), 1))
70            Next
80        End If

End Sub


Private Function IsQCDone() As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim AlarmHours As Long

10        On Error GoTo IsQCDone_Error

20        IsQCDone = True

30        AlarmHours = 24& * 365 * 10

40        sql = "Select top 1 * from StLukesAHG order by DateTime desc"
50        Set tb = New Recordset
60        RecOpenServerBB 0, tb, sql
70        If Not tb.EOF Then
80            If Abs(DateDiff("h", tb!DateTime, Now)) > AlarmHours Then
90                IsQCDone = False
100           End If
110       End If
120       sql = "Select top 1 * from StLukesGroupingCards order by DateTime desc"
130       Set tb = New Recordset
140       RecOpenServerBB 0, tb, sql
150       If Not tb.EOF Then
160           If Abs(DateDiff("h", tb!DateTime, Now)) > AlarmHours Then
170               IsQCDone = False
180           End If
190       End If

200       Exit Function

IsQCDone_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmMain", "IsQCDone", intEL, strES, sql

End Function

Private Sub btnRefresh_Click()
On Error GoTo ErrorHandler
    
    Call ShowDetail
    Call ShowProductDetail
    
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "btnRefresh_Click", intEL, strES
End Sub

Private Sub btnViewProducts_Click()
    frmViewProducts.Show 1
End Sub

Private Sub chkPending_Click()
    Call ShowProductDetail
End Sub

Private Sub chkProcess_Click()
    Call ShowProductDetail
End Sub

Private Sub chkRefresh_Click()
    If chkRefresh.Value = vbChecked Then
        Call ShowDetail
        Call ShowProductDetail
        timRefreshRecords.Enabled = True
    ElseIf chkRefresh.Value = vbUnchecked Then
        timRefreshRecords.Enabled = False
    End If
End Sub

Private Sub cmdAntibodies_Click()

10        frmAntibodyList.Show 1

End Sub

Private Sub cmdMatch_Click()

10        frmCheckMatch.Show 1

End Sub


Private Sub cmdPlaceOrder_Click()
On Error GoTo ErrorHandler

    Dim l_Count As Integer
    Dim sql As String
    Dim tb As ADODB.Recordset
    Dim tbR As ADODB.Recordset
    Dim l_RecCount As Integer
    Dim l_RequestID As String
    Dim l_FirstSampleID As String
    
    l_Count = 0
    For l_Count = 1 To flxProduct.Rows - 1
        If flxProduct.TextMatrix(l_Count, fcBit) = "Y" Then
            sql = "Update ocmRequestDetails Set Status = 'Process' "
            sql = sql & " Where UID = '" & flxProduct.TextMatrix(l_Count, fcUID) & "'"
            Cnxn(0).Execute sql
            DoEvents
            DoEvents
        End If
    Next
    Call ShowProductDetail
    DoEvents
    DoEvents
    
    l_Count = 0
    l_RecCount = 0
    l_RequestID = ""
    For l_Count = 1 To flxDetail.Rows - 1
        If flxDetail.TextMatrix(l_Count, fcsBit) = "Y" Then
            sql = "Insert into BBOrderComms(TestRequired,SampleID,Programmed,DateTimeOfRecord) "
            sql = sql & " Values('" & flxDetail.TextMatrix(l_Count, fcsTCode) & "','" & flxDetail.TextMatrix(l_Count, fcsSID) & "','1',GetDate())"
            CnxnBB(0).Execute sql
            If flxDetail.TextMatrix(l_Count, fcsR) <> l_RequestID Then
                sql = "Select Distinct IsNULL(RequestID,'') RequestID, convert(smalldatetime,SampleDate) SampleDate from ocmRequestDetails Where SampleID = '" & flxDetail.TextMatrix(l_Count, fcsSID) & "'"
                Set tb = New Recordset
                RecOpenServer 0, tb, sql
                If Not tb Is Nothing Then
                    If Not tb.EOF Then
                        sql = "Select * from ocmRequest Where RequestID = '" & tb!RequestID & "'"
                        Set tbR = New Recordset
                        RecOpenServer 0, tbR, sql
                        If Not tbR Is Nothing Then
                            If Not tbR.EOF Then
                                sql = "Insert into PatientDetails(patnum,name,ward,clinician,addr1,addr2,sex,DoB,SampleDate,GP,Hospital,RooH,AandE,SampleID,labnumber,DateReceived,Urgent) "
                                sql = sql & " Values('" & tbR!Chart & "','" & tbR!PatName & "','" & tbR!Ward & "','" & tbR!Clinician & "','" & tbR!Addr0 & "','" & tbR!Addr1 & "', '" & tbR!Sex & "','" & Format(tbR!DoB, "dd/mmm/yyyy") & "','" & Format(tb!SampleDate, "dd/mmm/yyyy hh:nn:ss") & "' ,'" & tbR!GP & "','C','" & tbR!RooH & "','" & tbR!AandE & "','" & flxDetail.TextMatrix(l_Count, fcsSID) & "','" & flxDetail.TextMatrix(l_Count, fcsSID) & "',getdate(),'" & tbR!Urgent & "')"
                                CnxnBB(0).Execute sql
                                
                            End If
                        End If
                    End If
                End If
                l_RequestID = flxDetail.TextMatrix(l_Count, fcsR)
            End If
            
            
            sql = "Update ocmRequestDetails Set trans = '1' "
            sql = sql & " Where SampleID = '" & flxDetail.TextMatrix(l_Count, fcsSID) & "'"
            Cnxn(0).Execute sql
            DoEvents
            DoEvents
        End If
    Next
    DoEvents
    DoEvents
    
    l_FirstSampleID = GetFirstSampleID()
'    MsgBox l_FirstSampleID
'    l_FirstSampleID = "1"
    If l_FirstSampleID <> "" Then
        frmxmatch.g_SampleID = l_FirstSampleID
        frmxmatch.Show 1
        frmxmatch.g_SampleID = l_FirstSampleID
'        Call frmxmatch.timFetchSampleID_Timer
        frmxmatch.timFetchSampleID.Enabled = True
    End If
    Call ShowDetail
    DoEvents
    DoEvents

    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
'    MsgBox Err.Description
    LogError "frmMain", "cmdPlaceOrder_Click", intEL, strES
End Sub

Private Sub flxDetail_Click()
'On Error GoTo ErrorHandler
'
'    If flxDetail.col = 1 Then
'        If flxDetail.row = 0 Then
'            Exit Sub
'        End If
'
'        If flxDetail.TextMatrix(flxDetail.row, fcsBit) = "Y" Then
'            flxDetail.TextMatrix(flxDetail.row, fcsBit) = "N"
'            '+++ SET EMPTY PICTUYRE
'            flxDetail.col = fcsLine_NO
'            Set flxDetail.CellPicture = imgRedCross.Picture
'        ElseIf flxDetail.TextMatrix(flxDetail.row, fcsBit) = "N" Then
'            flxDetail.TextMatrix(flxDetail.row, fcsBit) = "Y"
'            '+++ set picture
'            flxDetail.col = fcsLine_NO
'            Set flxDetail.CellPicture = imgGreenTick.Picture
'        End If
'        flxDetail.Redraw = True
'    End If
'
'    Call GetSelectStatus
'
'    Exit Sub
'ErrorHandler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "frmMain", "flxDetail_Click", intEL, strES
End Sub

Private Sub flxProduct_Click()
On Error GoTo ErrorHandler
    
    Dim sql As String
    
    If flxProduct.col = 1 Then
        If flxProduct.row = 0 Then
            Exit Sub
        End If
        
        If flxProduct.TextMatrix(flxProduct.row, fcBit) = "Y" Then
            flxProduct.TextMatrix(flxProduct.row, fcBit) = "N"
            '+++ SET EMPTY PICTUYRE
            flxProduct.col = fcLine_NO
            Set flxProduct.CellPicture = imgRedCross.Picture
        ElseIf flxProduct.TextMatrix(flxProduct.row, fcBit) = "N" Then
            flxProduct.TextMatrix(flxProduct.row, fcBit) = "Y"
            '+++ set picture
            flxProduct.col = fcLine_NO
            Set flxProduct.CellPicture = imgGreenTick.Picture
        End If
        flxProduct.Redraw = True
    End If
    
    If flxProduct.col = 10 Then
        If flxProduct.Rows > 1 Then
            If MsgBox("Are you sure to delete this row ?", vbInformation + vbYesNo) = vbYes Then
                sql = "Delete from ocmRequestDetails Where UID = '" & flxProduct.TextMatrix(flxProduct.row, fcUID) & "'"
                Cnxn(0).Execute sql
                Call ShowProductDetail
                DoEvents
                DoEvents
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "flxProduct_Click", intEL, strES
End Sub

Private Sub flxProduct_DblClick()
On Error GoTo ErrorHandler

    If flxProduct.row = 0 Then
        Exit Sub
    End If
    frmxmatch.g_SampleID = flxProduct.TextMatrix(flxProduct.row, fcSID)
    frmxmatch.Show 1
    frmxmatch.g_SampleID = flxProduct.TextMatrix(flxProduct.row, fcSID)
    
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "flxProduct_DblClick", intEL, strES
End Sub

Private Sub Form_Activate()

          Dim strVersion As String
          Dim strFileName As String

10        On Error GoTo frmActivateError

20        If ClosingFlag Then
30            Unload Me
40            Exit Sub
50        End If

60        strVersion = App.Major & "." & App.Minor & "." & App.Revision

70        If InStr(UCase$(App.Path), "TEST") Then
80            Me.BackColor = &H80FF&
90            lblTest.Visible = True
100       Else
110           Me.BackColor = &H8000000F
120           lblTest.Visible = False
130       End If

140       If Not IsIDE Then

150           If blnEndApp = True Then

160               strFileName = Dir(App.Path & "\" & strLatestVersion, vbNormal)
170               If Len(strFileName) = 0 Then
180                   iMsg "NetAcquire version has not changed!" & vbCr & vbCr & strLatestVersion & " does not exist!"
190                   If TimedOut Then Unload Me: Exit Sub
200                   blnEndApp = False
210                   Exit Sub
220               Else
230                   Shell App.Path & "\" & strLatestVersion, vbNormalFocus
240                   blnEndApp = False
250                   Unload Me
260                   Exit Sub
270               End If

280           End If

290       End If

300       bQC.Height = 735
          'Zyam added full appversion 26-1-24
310       strVersion = App.Major & "." & App.Minor & "." & App.Revision
          'Zyam
320       Me.Caption = "NetAcquire - Transfusion. Version " & strVersion

330       StatusBar.Panels(5) = strVersion

340       pBar = 0
350       If LogOffDelaySecs > 0 Then
360           pBar.max = LogOffDelaySecs
370       Else
380           pBar.max = 1
390       End If
400       Timer1.Enabled = True

410       CheckAboutToExpire
          'Zyam commented subs that were related to OCM 26-1-24
'          Call FormatGrid
'          Call ShowDetail
'          Call ShowProductDetail
          'Zyam
420       Exit Sub

frmActivateError:

          Dim er As Long
          Dim es As String
          Dim el As Long

430       er = Err.Number
440       es = Err.Description
450       el = Erl
460       iMsg "frmMain.Form.Activate Line " & Format(el) & " " & Format$(er) & " " & es
470       If TimedOut Then Unload Me: Exit Sub

End Sub
Private Sub CheckIfNewVersionAvailable()

          Dim strActiveVersion As String
          Dim fso As New FileSystemObject
          Dim fil As File
          Dim strFileName As String
          Dim dateCurrentApp As Date
          Dim dateNewApp As Date
          Dim strAppPath As String

          'Get Current active path\filename
10        On Error GoTo CheckIfNewVersionAvailable_Error

20        strAppPath = App.Path & "\"
30        strActiveVersion = strAppPath & App.EXEName & ".exe"

          'Get Date modified of Active version
          '  Set fso = CreateObject("Scripting.filesystemobject")
40        Set fil = fso.GetFile(strActiveVersion)
50        dateCurrentApp = fil.DateLastModified

          'Get 1st Exe name to check
60        strFileName = UCase$(Dir(strAppPath & "*.exe", vbNormal))

70        Do While strFileName <> ""    'Check all EXE files
80            If Right$(strFileName, 4) = ".EXE" Then    'Has to be an EXE
90                Set fil = fso.GetFile(strAppPath & strFileName)
100               dateNewApp = fil.DateLastModified  'Get files DateModified propertry

110               If dateNewApp > dateCurrentApp Then    ' IF New App ModifiedDate greater than current App ModifiedDate THEN display warning
120                   If AllowedToActivateVersion(strFileName) Then
130                       If Not tmrVersion.Enabled Then
140                           tmrVersion.Enabled = True
150                       End If
160                   End If
170               End If

180           End If
190           strFileName = UCase$(Dir)    'Get next file
200       Loop

210       Exit Sub

CheckIfNewVersionAvailable_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmMain", "CheckIfNewVersionAvailable", intEL, strES

End Sub
Public Function AllowedToActivateVersion(strFileName As String) As Boolean

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo AllowedToActivateVersion_Error

          'CheckVersionControlInDb CnxnBB(0)

20        sql = "Select * from VersionControl where FileName = '" & strFileName & "' and DoNotUse = 1 "

30        Set tb = New Recordset
40        RecOpenServerBB 0, tb, sql

50        If Not tb.EOF Then
60            AllowedToActivateVersion = False
70        Else
80            AllowedToActivateVersion = True
90        End If

100       Exit Function

AllowedToActivateVersion_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMain", "AllowedToActivateVersion", intEL, strES, sql

End Function

Private Sub Form_Click()

          Dim s As String
          '=R00011291915562
          '=R00011295791860
          '=R00011297078369
          '=R00011297081487

10        If IsIDE Then
20            s = ISOmod37_2("R00011291915562")    'E
30            s = ISOmod37_2("R00011295791860")    'A
40            s = ISOmod37_2("R00011297078369")    'S
50            s = ISOmod37_2("R00011297081487")    'M
60            s = ISOmod37_2("G123498654321")
70            s = ISOmod37_2("A999908123456")
80            s = ISOmod37_2("W000007123456")
90            s = ISOmod37_2("W000008123456")
100       End If

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        EnsureEventCodesInDatabase

30        If gEVENTCODES.Count = 0 Then
40            iMsg "Event Codes not Loaded"
50        End If

60        If App.PrevInstance Then End

70        ValidateCode = GetSetting("NetAcquire", "Transfusion6", "ValidateCode", "")
80        CancelCode = GetSetting("NetAcquire", "Transfusion6", "CancelCode", "")

90        mnuStLukesViewQC(0).Enabled = False
100       mnuStLukesViewQC(1).Enabled = False
110       mnuStLukesViewQC(2).Enabled = False

120       blnPrintingWithPreview = False
130       ShortCut.Width = 9585
140       cmdMatch.Visible = True
150       cmdPrint5.Visible = True
160       Entity = "01"
170       mnuMergeClinsWards.Visible = False
180       mnuStLukesViewQC(1).Enabled = True
190       mnuStLukesViewQC(2).Enabled = True
200       mnuUnitsByOrderNumber.Visible = True

210       If IsIDE Then
220           blnPrintingWithPreview = True
230       End If

240       LoadOptions

250       timAnalyserHeartBeat.Interval = GetOptionSetting("optHeartBeatTimeCheck_BTC", "10000")

260       strBTCourier_StorageLocation_StockFridge = GetOptionSetting("BTCourier_StorageLocation_StockEridge", "Stock Fridge")
270       strBTCourier_StorageLocation_RoomTempIssueFridge = GetOptionSetting("BTCourier_StorageLocation_RoomTempIssue", "Room Temp Issue")
280       strBTCourier_StorageLocation_HemoSafeFridge = GetOptionSetting("BTCourier_StorageLocation_HemoSafeFridge", "HemoSafe Fridge")

290       TransfusionForm = GetUserOptionSetting("TransfusionForm", "", UCase(vbGetComputerName))
300       TransfusionLabel = GetUserOptionSetting("TransfusionLabel", "", UCase(vbGetComputerName))
310       TransfusionPDF = GetUserOptionSetting("TransfusionPDF", "", UCase(vbGetComputerName))

320       If TransfusionForm = "" Or TransfusionLabel = "" Or TransfusionPDF = "" Then
330           frmPrinterSelect.Show 1
340       End If

350       mnuXMatchMain.Visible = True

360       pBar.Visible = False

370       EnsureColumnExistsBB "Latest", "ISBT128", "nvarchar(50)"
380       EnsureColumnExistsBB "Product", "ISBT128", "nvarchar(50)"
390       EnsureColumnExistsBB "PrintedLabels", "BarCode", "nvarchar(50)"
400       EnsureColumnExistsBB "PatientDetails", "PatSurName", "nvarchar (50) null"
410       EnsureColumnExistsBB "PatientDetails", "PatForeName", "nvarchar (50) null"

420       If IsIDE Then
430           UserName = "CRutter"
440           UserInitials = "CR"
450           UserCode = "C"
460           UserMemberOf = "Managers"
470           LogOffDelayMin = 5
480           LogOffDelaySecs = 300
490           EnableMenus True
500           UpdateLoggedOnUser
510       End If

520       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       MsgBox "Error in frmMain, Form_Load Line " & intEL

End Sub

Private Sub bantid_Click()

10        frmAntiD.Show 1

End Sub

Private Sub bgoodsin_Click()

10        frmNewStockBar.Show 1

End Sub

Private Sub bmove_Click()

10        With frmMovement
20            .Label1.Visible = False
30            .txtUnitNumber.Visible = False
40            .Show 1
50        End With

End Sub

Private Sub border_Click()

10        frmstock.Show 1

End Sub

Private Sub bpathistory_Click()

10        fpathistory.Show 1

End Sub

Private Sub bpatsearch_Click()

10        frmPatSearch.From = Me
20        frmPatSearch.btncopy.Enabled = False
30        frmPatSearch.Show 1

End Sub

Private Sub bqc_Click()

10        frmCavanAHG.Show 1
20        frmLukesQCOrthoWeekly.Show 1

End Sub

Private Sub bxmatch_Click()

10        frmxmatch.LoadingFlag = True

20        frmxmatch.Show 1

End Sub

Private Sub bxmquery_Click()

10        xmatchedsearch.Show 1

End Sub

Private Sub bxmreport_Click()

10        frmXMDaily.Show 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

          Dim f As Form

10        For Each f In Forms
20            Debug.Print "*"; f.Caption
30            Unload f
40        Next

End Sub

Private Sub imgSelect_Click()
On Error GoTo ErrorHandler

    Dim l_Count As Integer
    
    imgSelect.Visible = False
    imgUnSelect.Visible = True
    For l_Count = 1 To flxDetail.Rows - 1
        flxDetail.TextMatrix(l_Count, fcsBit) = "N"
    Next
    
    Call DrawFlx
    
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "imgSelect_Click", intEL, strES
End Sub

Private Sub imgUnSelect_Click()
On Error GoTo ErrorHandler

    Dim l_Count As Integer
    
    imgSelect.Visible = True
    imgUnSelect.Visible = False
    For l_Count = 1 To flxDetail.Rows - 1
        flxDetail.TextMatrix(l_Count, fcsBit) = "Y"
    Next
    
    Call DrawFlx
    
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "imgUnSelect_Click", intEL, strES
End Sub

Private Sub GetSelectStatus()
On Error GoTo ErrorHandler

    Dim l_Count As Integer
    
    For l_Count = 1 To flxDetail.Rows - 1
        If flxDetail.TextMatrix(l_Count, fcsBit) = "N" Then
            imgSelect.Visible = False
            imgUnSelect.Visible = True
        End If
    Next
    
    Exit Sub
ErrorHandler:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "GetSelectStatus", intEL, strES
End Sub

Private Sub malldaily_Click()

10        fdaily.optAll = True
20        fdaily.Show 1

End Sub

Private Sub manatal_Click()

          Dim Reason As String

10        If Not IsQCDone() Then
20            Answer = iMsg("QC has expired." & vbCrLf & vbCrLf & _
                            "Do you wish to continue?", vbQuestion + vbYesNo)
30            If TimedOut Then Exit Sub
40            If Answer = vbNo Then Exit Sub
50            Reason = iBOX("Why do you wish to continue" & vbCrLf & _
                            "with an expired QC")
60            If TimedOut Then Exit Sub
70            If Trim$(Reason) = "" Then Exit Sub
80            LogReasonWhy Reason, "XM"
90        End If

100       With frmxmatch
110           .LoadingFlag = True
120           .Show 1
130       End With

End Sub

Private Sub mantigens_Click()

10        frmAntigens.Show 1

End Sub

Private Sub mnuBatchEntry_Click()

10        frmBatchProductEntry.Show 1

End Sub

Private Sub mnuBatchHistoryOld_Click()

10        frmBatchHistory.Show 1

End Sub


Private Sub mnuBatchIssueOld_Click()

          Dim SID As String

10        SID = iBOX("Lab Number?")
20        If Trim$(SID) <> "" Then
30            With frmBatchIssue
40                .SampleID = SID
50                .Show 1
60            End With
70        End If

End Sub

Private Sub mnuBatchMovementOld_Click()

10        frmBatchMovement.Show 1

End Sub

Private Sub mnuBatchStock_Click()

10        frmBatchProductStock.Show 1

End Sub

Private Sub mnuBatchStockOld_Click()

10        frmBatchStock.Show 1

End Sub


Private Sub mnuBatchAmendment_Click()

10        frmBatchProductAmendment.Show 1

End Sub


Private Sub mby_Click(Index As Integer)

10        frmUsage.oSearchBy(Index).Value = True
20        frmUsage.Show 1

End Sub

Private Sub mchangehistory_Click()

10        frmChangeHistory.Show 1

End Sub

Private Sub mnuAntibodyList_Click()

10        frmAntibodyList.Show 1

End Sub

Private Sub mnuAudit_Click()

10        With frmArchive
20            .TableName = "PatientDetails"
30            .Show 1
40        End With

End Sub

Private Sub mnuBatchHistory_Click()

10        frmBatchProductHistory.Show 1

End Sub


Private Sub mnuCardValidation_Click()

10        frmViewCardValidation.Show 1

End Sub

Private Sub mnuChangePassword_Click()

          Dim NewPass As String
          Dim Confirm As String
          Dim tb As Recordset
          Dim sql As String
          Dim MinLength As Integer
          Dim Current As String
          Dim PasswordExpiry As Long
          Dim AllowReUse As String

10        On Error GoTo mnuChangePassword_Click_Error

20        Current = iBOX("Enter your current Password", , , True)
30        If TimedOut Then Unload Me: Exit Sub
40        sql = "SELECT * FROM Users WHERE " & _
                "Name = '" & AddTicks(UserName) & "' " & _
                "AND Password = '" & AddTicks(Current) & "' "
50        If GetOptionSetting("LogOnUpperLower", False) Then
60            sql = sql & "COLLATE SQL_Latin1_General_CP1_CS_AS"
70        End If
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then

110           NewPass = iBOX("Enter new password", , , True)
120           If TimedOut Then Unload Me: Exit Sub

130           MinLength = Val(GetOptionSetting("LogOnMinPassLen", "1"))
140           If Len(NewPass) < MinLength Then
150               iMsg "Passwords must have a minimum of " & Format(MinLength) & " characters!", vbExclamation
160               If TimedOut Then Unload Me: Exit Sub
170               Exit Sub
180           End If

190           If GetOptionSetting("LogOnUpperLower", False) Then
200               If AllLowerCase(NewPass) Or AllUpperCase(NewPass) Then
210                   iMsg "Passwords must have a mixture of UPPER CASE and lower case letters!", vbExclamation
220                   If TimedOut Then Unload Me: Exit Sub
230                   Exit Sub
240               End If
250           End If

260           If GetOptionSetting("LogOnNumeric", False) Then
270               If Not ContainsNumeric(NewPass) Then
280                   iMsg "Passwords must contain a numeric character!", vbExclamation
290                   If TimedOut Then Unload Me: Exit Sub
300                   Exit Sub
310               End If
320           End If

330           If GetOptionSetting("LogOnAlpha", False) Then
340               If Not ContainsAlpha(NewPass) Then
350                   iMsg "Passwords must contain an alphabetic character!", vbExclamation
360                   If TimedOut Then Unload Me: Exit Sub
370                   Exit Sub
380               End If
390           End If

400           AllowReUse = GetOptionSetting("PasswordReUse", "No")
410           If AllowReUse = "No" Then
420               If PasswordHasBeenUsed(NewPass) Then
430                   iMsg "Password has been used!", vbExclamation
440                   If TimedOut Then Unload Me: Exit Sub
450                   Exit Sub
460               End If
470           End If

480           Confirm = iBOX("Confirm password", , , True)
490           If TimedOut Then Unload Me: Exit Sub

500           If NewPass <> Confirm Then
510               iMsg "Passwords don't match!", vbExclamation
520               If TimedOut Then Unload Me: Exit Sub
530               Exit Sub
540           End If

550           ArchiveTable "Users", "", sql
560           Cnxn(0).Execute sql

570           PasswordExpiry = Val(GetOptionSetting("PasswordExpiry", "90"))

580           sql = "UPDATE Users SET " & _
                    "PassWord = '" & NewPass & "', " & _
                    "PassDate = '" & Format$(Now + PasswordExpiry, "dd/MMM/yyyy") & "' WHERE " & _
                    "Name = '" & AddTicks(UserName) & "'"
590           Cnxn(0).Execute sql

600           iMsg "Your Password has been changed.", vbInformation
610           If TimedOut Then Unload Me: Exit Sub

620       End If

630       Exit Sub

mnuChangePassword_Click_Error:

          Dim strES As String
          Dim intEL As Integer

640       intEL = Erl
650       strES = Err.Description
660       LogError "frmMain", "mnuChangePassword_Click", intEL, strES, sql

End Sub

Private Sub mnuCodeCancel_Click()

10        CancelCode = iBOX("Scan or Enter Code for Cancel", , CancelCode)
20        If TimedOut Then Unload Me: Exit Sub

30        SaveSetting "NetAcquire", "Transfusion6", "CancelCode", CancelCode

End Sub

Private Sub mnuCodeValidate_Click()

10        ValidateCode = iBOX("Scan or Enter Code for Validate", , ValidateCode)
20        If TimedOut Then Unload Me: Exit Sub

30        SaveSetting "NetAcquire", "Transfusion6", "ValidateCode", ValidateCode

End Sub


Private Sub mCurrentStock_Click()

10        fCurrentStock.Show 1

End Sub

Private Sub mdaily_Click(Index As Integer)

10        With fdaily
20            Select Case Index
              Case 1: .optGH = True
30            Case 2: .optAN = True
40            Case 3: .optDAT = True
50            Case Else: .optAll = True
60            End Select
70            .Show 1
80        End With

End Sub

Private Sub mdat_Click()

          Dim Reason As String

10        If Not IsQCDone() Then
20            Answer = iMsg("QC has expired." & vbCrLf & vbCrLf & _
                            "Do you wish to continue?", vbQuestion + vbYesNo)
30            If TimedOut Then Exit Sub
40            If Answer = vbNo Then Exit Sub
50            Reason = iBOX("Why do you wish to continue" & vbCrLf & _
                            "with an expired QC")
60            If TimedOut Then Exit Sub
70            If Trim$(Reason) = "" Then Exit Sub
80            LogReasonWhy Reason, "XM"
90        End If

100       With frmxmatch
110           .LoadingFlag = True
120           .Show 1
130       End With

End Sub

Private Sub menternewstock_Click()

10        frmNewStockBar.Show 1

End Sub

Private Sub mFreeStock_Click()

10        frmFreeStock.Show 1

End Sub


Private Sub mgenefreq_Click()

10        frmGeneFreq.Show 1

End Sub

Private Sub mghold_Click()

          Dim Reason As String

10        If Not IsQCDone() Then
20            Answer = iMsg("QC has expired." & vbCrLf & vbCrLf & _
                            "Do you wish to continue?", vbQuestion + vbYesNo)
30            If TimedOut Then Exit Sub
40            If Answer = vbNo Then Exit Sub
50            Reason = iBOX("Why do you wish to continue" & vbCrLf & _
                            "with an expired QC")
60            If TimedOut Then Exit Sub
70            If Trim$(Reason) = "" Then Exit Sub
80            LogReasonWhy Reason, "XM"
90        End If

100       With frmxmatch
110           .LoadingFlag = True
120           .Show 1
130       End With

End Sub

Private Sub mhelpabout_Click()

10        frmAbout.Show 1

End Sub

Private Sub mhelpcontents_Click()

10        iMsg "Not Yet Implimented"
20        If TimedOut Then Unload Me: Exit Sub

End Sub

Private Sub mhelpobtainingmore_Click()

          Dim s As String

10        s = "Technical assistance can be obtained" & vbCrLf & _
              "by phoning 057-8601230" & vbCrLf & _
              "Or e-mail info@customsoftware.ie"
20        iMsg s, vbInformation
30        If TimedOut Then Unload Me: Exit Sub

End Sub

Private Sub mhelponhelp_Click()

10        iMsg "Not Yet Implimented"
20        If TimedOut Then Unload Me: Exit Sub

End Sub

Private Sub mhistory_Click()

10        With frmUnitHistory
20            .UnitNumber = ""
30            .ProductName = ""
40            .Show 1
50        End With

End Sub

Private Sub mListBatches_Click()

10        flists.ListName = "B"
20        flists.oList(4) = True
30        flists.Show 1

End Sub

Private Sub mListGPs_Click()

10        fgps.Show 1

End Sub

Private Sub mListHospitals_Click()

10        fHospital.Show 1

End Sub

Private Sub mmonthly_Click()

10        fmonthly.Show 1

End Sub

Private Sub mnuAddIncident_Click()

          Dim sql As String
          Dim MSG As String

10        On Error GoTo mnuAddIncident_Click_Error

20        MSG = iBOX("New Incident")
30        If TimedOut Then Unload Me: Exit Sub
40        If Trim$(MSG) = "" Then Exit Sub

50        sql = "Insert into IncidentLog " & _
                "(Incident, DateTime, Technician ) " & _
                "VALUES " & _
                "('" & AddTicks(MSG) & "', " & _
                "'" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
                "'" & UserName & "')"
60        CnxnBB(0).Execute sql

70        Exit Sub

mnuAddIncident_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mnuAddIncident_Click", intEL, strES, sql


End Sub

Private Sub mnuAmendStatus_Click()

10        frmAmend.Show 1

End Sub

Private Sub mnuantibodies_Click()

10        frmDefineABPanel.Show 1

End Sub

Private Sub mnuantid_Click()

10        frmAntiDUsage.Show 1

End Sub

Private Sub mnuBatchAmendmentOld_Click()

10        frmBatchStatus.Show 1

End Sub

Private Sub mnuBatchMovement_Click()

10        frmBatchProductMovement.Show 1

End Sub


Private Sub mnuBatchRestock_Click()

10        frmBatchRestock.Show 1

End Sub

Private Sub mnucliniccondx_Click()

10        flists.ListName = "X"
20        flists.oList(0) = True
30        flists.Show 1

End Sub

Private Sub mnuclose_Click()

10        ClosingFlag = True
20        Unload Me

End Sub

Private Sub mnuConfirmGroup_Click()
10        frmGrpConf.Show 1
End Sub

Private Sub mnuCourierUnitHistory_Click()

10        frmCourierUnitHistory.Show 1

End Sub

Private Sub mnuCourierViewLog_Click()

10        frmViewCourier.Show 1

End Sub

Private Sub mnudaily_Click()

10        bqc_Click

End Sub

Private Sub mnudestroyed_Click()

10        With frmSearch
20            .Caption = "Destroyed Product"
30            .lblsearchfor.Caption = "D"
40            .Show 1
50        End With

End Sub

Private Sub mnuErrorLog_Click()

10        frmErrorLog.Show 1

End Sub

Private Sub mnuGroupChecked_Click()

10        frmGroupChecked.Show 1

End Sub

Private Sub mnuKleihauer_Click()

10        frmKleihauer.Show 1

End Sub

Private Sub mnuKleihauerQC_Click()

10        frmKleihauerQC.Show 1

End Sub

Private Sub mnulogoff_click()

10        EnableMenus False

20        UpdateLoggedOnUser

30        timAnalyserHeartBeat.Enabled = False

End Sub

Private Sub UpdateLoggedOnUser()

          Dim sql As String
          Dim MachineName As String

10        On Error GoTo UpdateLoggedOnUser_Error

20        MachineName = UCase$(vbGetComputerName())

30        sql = "IF EXISTS (SELECT * FROM LoggedOnUsers WHERE " & _
              "           MachineName = '" & MachineName & "' " & _
              "           AND AppName = 'Transfusion') " & _
              "  UPDATE LoggedOnUsers " & _
              "  SET UserName = '" & AddTicks(UserName) & "' " & _
              "  WHERE MachineName = '" & MachineName & "' " & _
              "  AND AppName = 'Transfusion' " & _
                "ELSE " & _
              "  INSERT INTO LoggedOnUsers " & _
              "  (MachineName, AppName, UserName) VALUES " & _
              "  ('" & MachineName & "', " & _
              "  'Transfusion', " & _
              "  '" & AddTicks(UserName) & "')"
40        CnxnBB(0).Execute sql
          '
          '30    sql = "SELECT * FROM LoggedOnUsers WHERE " & _
           '            "MachineName = '" & MachineName & "' " & _
           '            "AND AppName = 'Transfusion'"
          '40    Set tb = New Recordset
          '50    RecOpenServerBB 0, tb, sql
          '60    If tb.EOF Then
          '70      tb.AddNew
          '80    End If
          '90    tb!MachineName = MachineName
          '100   tb!UserName = UserName
          '110   tb!AppName = "Transfusion"
          '
          '120   tb.Update

50        Exit Sub

UpdateLoggedOnUser_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMain", "UpdateLoggedOnUser", intEL, strES, sql

End Sub

Private Sub mnulogon_Click()

10        GetLogOn

20        EnableMenus UserCode <> ""

30        UpdateLoggedOnUser

End Sub

Private Sub mnuMergeClinsWards_Click()

10        frmMergeClinsWards.Show 1

End Sub

Private Sub mnuMonthlyWastage_Click()

10        frmMonthlyWastage.Show 1

End Sub

Private Sub mnumove_Click()

10        With frmMovement
20            .Label1.Visible = False
30            .txtUnitNumber.Visible = False
40            .Show 1
50        End With

End Sub

Private Sub mnuMoveCodabar_Click()

10        With frmMovement
20            .txtUnitNumber.Left = 735
30            .Label1.Left = 105
40            .txtUnitNumber.Visible = True
50            .Label1.Visible = True

60            .txtISBT128.Visible = False
70            .txtISBT128.Enabled = False
80            .Label6.Visible = False
90            .Show 1
100       End With

End Sub

Private Sub mnuPatListByProduct_Click()

10        frmPatListByProduct.Show 1

End Sub

Private Sub mnupatnew_Click()

          Dim Reason As String

10        If Not IsQCDone() Then
20            Answer = iMsg("QC has expired." & vbCrLf & vbCrLf & _
                            "Do you wish to continue?", vbQuestion + vbYesNo)
30            If TimedOut Then Exit Sub
40            If Answer = vbNo Then Exit Sub
50            Reason = iBOX("Why do you wish to continue" & vbCrLf & _
                            "with an expired QC")
60            If TimedOut Then Exit Sub
70            If Trim$(Reason) = "" Then Exit Sub
80            LogReasonWhy Reason, "XM"
90        End If

100       frmxmatch.Show 1

End Sub

Private Sub mnupatsearch_Click()

10        frmPatSearch.From = Me
20        frmPatSearch.btncopy.Enabled = False
30        frmPatSearch.Show 1

End Sub

Private Sub mnuphysicians_Click()

10        fclinicians.Show 1

End Sub

Private Sub mnuPrinterSelection_Click()

10        frmPrinterSelect.Show 1

End Sub

Private Sub mnuPrintReclaimed_Click()

10        fPrintReclaimed.Show 1

End Sub

Private Sub mnureagents_Click()

10        fReagents.Show 1

End Sub

Private Sub mnureceived_Click()

10        With frmSearch
20            .Caption = "Received Product"
30            .lblsearchfor.Caption = "C"
40            .Show 1
50        End With

End Sub

Private Sub mnuResetLastUsed_Click()

          Dim LastUsed As String

10        LastUsed = GetSetting("NetAcquire", "Transfusion6", "LastUsed", "1")

20        LastUsed = iBOX("Enter 'Last Used' Number", , LastUsed)
30        If TimedOut Then Unload Me: Exit Sub

40        If Val(LastUsed) <> 0 Then
50            SaveSetting "NetAcquire", "Transfusion6", "LastUsed", LastUsed
60        End If

End Sub

Private Sub mnurestocked_Click()

10        With frmSearch
20            .Caption = "Restocked Product"
30            .lblsearchfor.Caption = "R"
40            .Show 1
50        End With

End Sub

Private Sub mnureturned_Click()

10        With frmSearch
20            .Caption = "Returned Product"
30            .lblsearchfor.Caption = "T"
40            .Show 1
50        End With

End Sub

Private Sub mnuSound_Click()

10        frmSound.Show 1

End Sub

Private Sub mnuStLukesViewQC_Click(Index As Integer)

10        Select Case Index
          Case 0: frmLukesCentrifugeView.Show 1
20        Case 1: frmLukesAHGView.Show 1
30        Case 2: frmLukesQCOrthoWeeklyView.Show 1
40        Case 3: frmLukesPhenotypeView.Show 1
50        End Select

End Sub



Private Sub mnuStockReconciliation_Click()
10        frmStockReconciliation.Show 1
End Sub

Private Sub mnusurgical_Click()

10        flists.ListName = "P"
20        flists.oList(1) = True
30        flists.Show 1

End Sub

Private Sub mnutemperatures_Click()

10        fTempsQC.Show 1

End Sub

Private Sub mnuti_Click()

10        ftransindex.Show 1

End Sub

Private Sub mnuTransferred_Click()

10        With frmSearch
20            .Caption = "Transferred Product"
30            .lblsearchfor.Caption = "F"
40            .Show 1
50        End With

End Sub

Private Sub mnutransfused_Click()

10        frmTransfused.Show 1

End Sub

Private Sub mnuUnitFating_Click()

10        frmUnitFating.Show 1

End Sub

Private Sub mnuUnitsByOrderNumber_Click()
10        frmUnitsByOrderNumber.Show 1
End Sub

Private Sub mnuUnlockReasons_Click()

10        frmUnlockReasons.Show 1

End Sub

Private Sub mnuVersionControl_Click()

10        frmVersionControl.Show 1

End Sub

Private Sub mnuViewIncident_Click()

10        frmIncidentLog.Show 1

End Sub

Private Sub mnuwards_Click()

10        fWardList.Show 1

End Sub

Private Sub mnuxmatch_Click()

10        With frmxmatch
20            .LoadingFlag = True
30            .Show 1
40        End With

End Sub

Private Sub mnuxmatched_Click()

10        xmatchedsearch.Show 1

End Sub



Private Sub mnuXMComments_Click()

10        flists.ListName = "XC"
20        flists.oList(5) = True
30        flists.Show 1

End Sub

Private Sub mnuYearlyStats_Click()

10        frmYearlyStats.Show 1

End Sub

Private Sub morder_Click()

10        frmstock.Show 1

End Sub

Private Sub mpatenq_Click()

10        frmPatSearch.From = Me
20        frmPatSearch.btncopy.Enabled = False
30        frmPatSearch.Show 1

End Sub

Private Sub mpathistory_Click()

10        fpathistory.Show 1

End Sub

Private Sub mPrinters_Click()

10        fPrinters.Show 1

End Sub

Private Sub mProductList_Click()

10        frmProducts.Show 1

End Sub

Private Sub mreaction_Click()

10        freaction.Show 1

End Sub

Private Sub mservice_Click()

10        fservice.Show 1

End Sub

Private Sub msprod_Click()

10        flists.ListName = "S"
20        flists.oList(2) = True
30        flists.Show 1

End Sub

Private Sub mstocklist_Click()

10        frmStockList.Show 1

End Sub

Private Sub mStockListBatches_Click()

10        frmBatchStock.Show 1

End Sub


Private Sub msuppliers_Click()

10        fSuppliers.Show 1

End Sub

Private Sub mtotals_Click()

10        frmTotals.Show 1

End Sub

Private Sub mtransfuse_Click()

10        fprodmove.Show 1

End Sub

Private Sub mworklists_Click()

10        frmPatientWorkList.Show 1

End Sub



Private Sub mxmd2_Click()

10        frmXMDaily.Show 1

End Sub

Private Sub mxmdaily_Click()

10        frmXMDaily.Show 1

End Sub

Private Sub mxmworklist_Click()

10        fWorkListXM.Show 1

End Sub



Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)

10        If Panel.Index = 4 And tmrVersion.Enabled Then
20            If UserMemberOf = "Managers" Then
30                frmVersionControl.Show 1
40            End If
50        End If

End Sub

Private Sub timAnalyserHeartBeat_Timer()
          Dim datHeartBeatTimeNow As Date
          Dim intHeartBeatInterval As Integer
          Dim douIntervalNow As Double
          Dim s As String

10        On Error GoTo timAnalyserHeartBeat_Timer_Error

20        intHeartBeatInterval = GetOptionSetting("optHeartBeatEquipmentInterval_BTC", "5")
30        datHeartBeatTimeNow = Format(GetOptionSetting("optHeartBeatTimeStamp_BTC", "06/Aug/2013 10:33:00"), "dd/mmm/yyyy hh:nn:ss")

40        douIntervalNow = DateDiff("s", datHeartBeatTimeNow, Format(Now, "dd/mmm/yyyy hh:nn:ss"))

50        If douIntervalNow > intHeartBeatInterval Then
60            imgOK.Visible = False
70            imgNOTOK.Visible = True
80            imgNOTOK.Top = 150
90            If Not blnBTCdownWarningDisplayed Then
100               s = iBOX(vbCrLf & "Blood Track Courier communications down!" & vbCrLf & vbCrLf & vbCrLf & "Enter your Password", , , True)
110               If TimedOut Then Unload Me: Exit Sub
                  'Do While s <> TechnicianPasswordForName(UserName)
                  '    iMsg "Invalid Password.", vbInformation
                  '    s = iBOX(vbCrLf & "Blood Track Courier communications down!" & vbCrLf & vbCrLf & vbCrLf & "Enter your Password", , , True)
                  'Loop
120               blnBTCdownWarningDisplayed = True
130           End If
140       Else
150           imgOK.Visible = True
160           imgNOTOK.Visible = False
170           imgNOTOK.Top = 150
180           blnBTCdownWarningDisplayed = False
190       End If

200       Exit Sub

timAnalyserHeartBeat_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmMain", "timAnalyserHeartBeat_Timer", intEL, strES

End Sub

Private Sub Timer1_Timer()

          Dim WarningTime As Boolean
          Dim DayOfWeek As Integer
          Dim HourOfDay As Integer
          Static X As Long
          Static Y As Long
          Dim TempX As Long
          Dim TempY As Long
          Dim H As Long

10        On Error GoTo Timer1_Timer_Error

20        If mnulogoff.Enabled Then
30            StatusBar.Panels(3).Text = UserName
40        Else
50            StatusBar.Panels(3).Text = ""
60        End If
          'Zyam Date
70        If Format$("04/05/2001", "dd/mmm/yyyy") <> "04/May/2001" Then
80            iMsg "Date/Time Format in" & vbCrLf & _
                   "International Settings" & vbCrLf & _
                   "are not set correctly." & vbCrLf & vbCrLf & _
                   "Cannot proceed!", vbCritical
90            If TimedOut Then Unload Me: Exit Sub
100           End
110       End If

120       If TimedOut Then
130           Debug.Print Forms.Count
140           If Forms.Count > 1 Then
150               Do While Forms.Count > 1
160                   Unload Forms(Forms.Count - 1)
170               Loop
180           Else
190               TimedOut = False
200           End If
210       End If

220       If TopMostWindow() = Screen.ActiveForm.Caption Then

230           H = Screen.ActiveForm.hwnd

240           TempX = MouseX(H)
250           TempY = MouseY(H)
260           If X <> TempX Or Y <> TempY Then
270               If TempX > 0 And _
                     TempY > -30 And _
                     TempX * Screen.TwipsPerPixelX < Screen.ActiveForm.Width And _
                     TempY * Screen.TwipsPerPixelY < Screen.ActiveForm.Height - 320 Then
280                   X = TempX
290                   Y = TempY
300                   Screen.ActiveForm.Controls("pBar").Value = 0
310               End If
320           End If

330           If KB() Then
340               Screen.ActiveForm.Controls("pBar").Value = 0
350           End If

360       End If

370       With Screen.ActiveForm.Controls("pBar")
380           If LogOffDelaySecs <> 0 Then
390               .max = LogOffDelaySecs
400           Else
410               .max = 30
420           End If
430           .Value = .Value + 1
440           If .Value = .max Then
450               .Value = 0
460               TimedOut = True
470               Do While Screen.ActiveForm.Name <> Me.Name
480                   Unload Screen.ActiveForm
490               Loop
500               EnableMenus False
510           End If
520       End With

530       If IsQCDone() Then
540           bQC.Visible = True
550           bXmatch.Enabled = True
560           bMove.Enabled = True
570           Exit Sub
580       End If

590       WarningTime = False

600       DayOfWeek = Val(Format(Now, "w"))
610       If DayOfWeek = vbSaturday Or DayOfWeek = vbSunday Then
620           WarningTime = True
630       Else
640           HourOfDay = Val(Format(Now, "hh"))
650           If HourOfDay < 12 Then
660               WarningTime = True
670           End If
680       End If


690       If Not WarningTime Then
700           bQC.Visible = Not bQC.Visible
710           bXmatch.Enabled = False
720           bMove.Enabled = False
730       Else
740           bQC.Visible = True
750           bXmatch.Enabled = True
760           bMove.Enabled = True
770           If bQC.Height = 735 Then
780               bQC.Height = 765
790           Else
800               bQC.Height = 735
810           End If
820       End If

830       Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

840       intEL = Erl
850       strES = Err.Description
860       LogError "frmMain", "Timer1_Timer", intEL, strES

End Sub

Private Sub timRefreshRecords_Timer()
    Call btnRefresh_Click
End Sub

Private Sub tmrVersion_Timer()

          Static blnCounter As Boolean

10        On Error GoTo tmrVersionError

20        If blnCounter Then
30            'StatusBar.Panels("Indicator").Picture = imgMisc.ListImages("Indicator").Picture
40            DoEvents
50            blnCounter = False
60        Else
70            StatusBar.Panels("Indicator").Picture = Nothing
80            DoEvents
90            blnCounter = True
100       End If

110       Exit Sub

tmrVersionError:
          Dim er As Long
          Dim es As String

120       er = Err.Number
130       es = Err.Description
140       iMsg "tmrVersion.Timer " & Format$(er) & " " & es
150       If TimedOut Then Unload Me: Exit Sub

End Sub


Private Sub mnuABO_AntiD_Control_History_Click()
10        frmABOAntiDBloodGroupControlHistory.Show 1
End Sub

Private Sub SelectSampleID(p_SampleID As String)
    On Error GoTo ERROR_ShowDetail
    
    Dim l_Count As Integer
    Dim l_Found As Boolean
    
    l_Found = False
    For l_Count = 1 To flxDetail.Rows - 1
        If flxDetail.TextMatrix(l_Count, fcsSID) = p_SampleID Then
            flxDetail.TextMatrix(l_Count, fcsBit) = "Y"
            l_Found = True
        End If
    Next
    If l_Found Then
        Call DrawFlx
    Else
        MsgBox "Sample ID not found.", vbInformation
    End If
    
    Exit Sub
ERROR_ShowDetail:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "SelectSampleID", intEL, strES
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
    On Error GoTo ERROR_ShowDetail
    
    If KeyAscii = 13 Then
        Call SelectSampleID(Trim(txtBarcode.Text))
        txtBarcode.Text = ""
    End If
    
    Exit Sub
ERROR_ShowDetail:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "txtBarcode_KeyPress", intEL, strES
End Sub

Private Function GetFirstSampleID()
    On Error GoTo ERROR_ShowDetail
    
    Dim l_Count As Integer
    
    GetFirstSampleID = ""
    For l_Count = 1 To flxDetail.Rows - 1
        If flxDetail.TextMatrix(l_Count, fcsBit) = "Y" Then
            GetFirstSampleID = flxDetail.TextMatrix(l_Count, fcsSID)
            Exit Function
        End If
    Next
    
    Exit Function
ERROR_ShowDetail:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "GetFirstSampleID", intEL, strES
End Function
