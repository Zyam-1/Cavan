VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Custom Software"
   ClientHeight    =   9255
   ClientLeft      =   765
   ClientTop       =   1440
   ClientWidth     =   16860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9255
   ScaleWidth      =   16860
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fMain"
   Begin VB.Timer microRepTim 
      Left            =   7830
      Top             =   4380
   End
   Begin VB.CommandButton cmdResetLabNo 
      Caption         =   "Reset Lab No"
      Height          =   375
      Left            =   12570
      TabIndex        =   54
      Top             =   540
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer Live -> Test"
      Height          =   375
      Left            =   13980
      TabIndex        =   49
      Top             =   540
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.ListBox lstAutoValFail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      IntegralHeight  =   0   'False
      Left            =   17550
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   3960
      Width           =   1875
   End
   Begin VB.CommandButton cmdSemen 
      Height          =   825
      Left            =   2190
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   90
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdMicro 
      Height          =   825
      Left            =   1140
      Picture         =   "frmMain.frx":20FC
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   90
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdBio 
      Height          =   825
      Left            =   120
      Picture         =   "frmMain.frx":2FC6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ComboBox cmbResultDays 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":3E90
      Left            =   8610
      List            =   "frmMain.frx":3EA6
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   360
      Width           =   795
   End
   Begin VB.CheckBox chkAutoRefresh 
      Alignment       =   1  'Right Justify
      Caption         =   "Auto-Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7530
      TabIndex        =   12
      Top             =   120
      Width           =   1275
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   30000
      Left            =   16590
      Top             =   360
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   17190
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EBF
            Key             =   "Ring0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41D9
            Key             =   "Ring1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44F3
            Key             =   "Ring2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":480D
            Key             =   "Ring3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B27
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51A1
            Key             =   "Fax"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55F3
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A45
            Key             =   "Locked"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7872
      Left            =   90
      TabIndex        =   2
      Top             =   1020
      Width           =   16755
      Begin VB.CommandButton btnComparison 
         Caption         =   "&Comparison"
         Height          =   435
         Left            =   6660
         TabIndex        =   76
         Top             =   7230
         Width           =   1905
      End
      Begin VB.Frame fmeComparison 
         Caption         =   "Comparison"
         Height          =   6405
         Left            =   3690
         TabIndex        =   69
         Top             =   420
         Visible         =   0   'False
         Width           =   9585
         Begin VB.CommandButton btnProcess 
            Caption         =   "Process"
            Height          =   435
            Left            =   7860
            TabIndex        =   77
            Top             =   210
            Width           =   1605
         End
         Begin VB.CommandButton btnHide 
            Caption         =   "&Hide"
            Height          =   435
            Left            =   7860
            TabIndex        =   75
            Top             =   5880
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker dtoFrom 
            Height          =   315
            Left            =   1260
            TabIndex        =   72
            Top             =   300
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            Format          =   228392961
            CurrentDate     =   45348
         End
         Begin MSFlexGridLib.MSFlexGrid flxComp 
            Height          =   5145
            Left            =   90
            TabIndex        =   70
            Top             =   720
            Width           =   9435
            _ExtentX        =   16642
            _ExtentY        =   9075
            _Version        =   393216
            Cols            =   7
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   4920
            TabIndex        =   74
            Top             =   300
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            Format          =   195493889
            CurrentDate     =   45348
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "To Date:"
            Height          =   195
            Left            =   4110
            TabIndex        =   73
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "From Date:"
            Height          =   195
            Left            =   270
            TabIndex        =   71
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdOrderAddOns 
         Caption         =   "&Order AddOns"
         Enabled         =   0   'False
         Height          =   435
         Left            =   12015
         TabIndex        =   63
         Top             =   6750
         Width           =   1905
      End
      Begin VB.Frame fraSelectOrder 
         Height          =   435
         Index           =   0
         Left            =   10755
         TabIndex        =   59
         Top             =   5580
         Width           =   1995
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   0
            Left            =   1320
            Picture         =   "frmMain.frx":5E97
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   0
            Left            =   1650
            Picture         =   "frmMain.frx":616D
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<-Select for Order"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   44
            Left            =   45
            TabIndex        =   62
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmdScaneSample 
         Caption         =   "&Scan Samples"
         Enabled         =   0   'False
         Height          =   435
         Left            =   12000
         TabIndex        =   56
         Top             =   7230
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.ComboBox cmbBioNoResult 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdValidationList 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Demographic Validation"
         Enabled         =   0   'False
         Height          =   435
         Left            =   13980
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   7230
         Width           =   2688
      End
      Begin VB.ListBox lstCoagNotValid 
         Height          =   1230
         Left            =   3540
         TabIndex        =   37
         Top             =   4236
         Width           =   2652
      End
      Begin VB.ListBox lstHaemNotValid 
         Height          =   1932
         IntegralHeight  =   0   'False
         Left            =   3555
         TabIndex        =   34
         Top             =   2280
         Width           =   2652
      End
      Begin VB.CommandButton cmdUnvalidated 
         Caption         =   "&View Unvalidate/Not Printed"
         Enabled         =   0   'False
         Height          =   435
         Left            =   13980
         TabIndex        =   28
         Top             =   6240
         Width           =   2688
      End
      Begin VB.CommandButton cmdUnvalidatedSamples 
         Caption         =   "Microbiology Un&validated Samples"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13980
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6735
         Width           =   2688
      End
      Begin MSFlexGridLib.MSFlexGrid grdAutoValFail 
         Height          =   1635
         Left            =   300
         TabIndex        =   24
         Top             =   6180
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   2884
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Sample ID  |^D|^V|^O|^R|^D|^F|^A|^24"
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
      Begin MSFlexGridLib.MSFlexGrid grdPhone 
         Height          =   1935
         Index           =   0
         Left            =   13140
         TabIndex        =   15
         Top             =   300
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   255
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "<Sample ID |<Parameter   |<Ward                    "
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
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   6780
         Picture         =   "frmMain.frx":6443
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print Lists to Default Printer"
         Top             =   5580
         Width           =   585
      End
      Begin MSFlexGridLib.MSFlexGrid gBioNoResults 
         Height          =   1575
         Left            =   300
         TabIndex        =   9
         ToolTipText     =   "Click on Heading to Sort"
         Top             =   660
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^Analyser    |<Sample ID            "
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
      Begin VB.Timer tmrNotPrinted 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10020
         Top             =   5460
      End
      Begin VB.ListBox lstCoagNotPrinted 
         Height          =   1230
         Left            =   6720
         TabIndex        =   8
         Top             =   4236
         Width           =   2652
      End
      Begin VB.ListBox lstCoagNoResults 
         Height          =   1230
         Left            =   300
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   4236
         Width           =   2652
      End
      Begin VB.ListBox lstHaemNotPrinted 
         Height          =   1932
         IntegralHeight  =   0   'False
         Left            =   6720
         TabIndex        =   6
         Top             =   2280
         Width           =   2652
      End
      Begin MSFlexGridLib.MSFlexGrid gBioNotPrinted 
         Height          =   1935
         Left            =   6720
         TabIndex        =   10
         ToolTipText     =   "Click on Heading to Sort, Left Click to View/Edit, Right Click to Remove"
         Top             =   315
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^Analyser      |<Sample ID         "
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
      Begin MSFlexGridLib.MSFlexGrid gHaemNoResults 
         Height          =   1932
         Left            =   300
         TabIndex        =   13
         ToolTipText     =   "Click on Heading to Sort"
         Top             =   2280
         Width           =   2652
         _ExtentX        =   4683
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Sample ID|<Date                      "
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
      Begin MSFlexGridLib.MSFlexGrid grdPhone 
         Height          =   1770
         Index           =   2
         Left            =   13140
         TabIndex        =   16
         Top             =   4230
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   3122
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   255
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "<Sample ID |<Parameter   |<Ward                    "
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
      Begin MSFlexGridLib.MSFlexGrid grdPhone 
         Height          =   1935
         Index           =   1
         Left            =   13140
         TabIndex        =   27
         Top             =   2280
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   255
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "<Sample ID |<Parameter   |<Ward                    "
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
      Begin MSFlexGridLib.MSFlexGrid gBioNotValid 
         Height          =   1932
         Left            =   3540
         TabIndex        =   33
         ToolTipText     =   "Click on Heading to Sort, Left Click to View/Edit, Right Click to Remove"
         Top             =   300
         Width           =   2652
         _ExtentX        =   4683
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^Analyser     |<Sample ID           "
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
      Begin MSFlexGridLib.MSFlexGrid grdUrg 
         Height          =   1575
         Left            =   3525
         TabIndex        =   41
         Top             =   6180
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Sample ID     |^H |^B |^C |^E |^G |^I  "
      End
      Begin MSFlexGridLib.MSFlexGrid grdAddOns 
         Height          =   1905
         Left            =   10110
         TabIndex        =   58
         ToolTipText     =   "Click on Heading to Sort, Left Click to View/Edit, Right Click to Remove"
         Top             =   315
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3360
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^SampleID    |<Sample Date |<         "
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
      Begin MSFlexGridLib.MSFlexGrid grdAddOnsHAEM 
         Height          =   1905
         Left            =   10110
         TabIndex        =   64
         ToolTipText     =   "Click on Heading to Sort, Left Click to View/Edit, Right Click to Remove"
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3360
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^SampleID    |<Sample Date |<         "
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
      Begin MSFlexGridLib.MSFlexGrid grdAddOnsCOAG 
         Height          =   1275
         Left            =   10110
         TabIndex        =   65
         ToolTipText     =   "Click on Heading to Sort, Left Click to View/Edit, Right Click to Remove"
         Top             =   4200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2249
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^SampleID    |<Sample Date |<         "
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
      Begin VB.Label Label1 
         Caption         =   "C O A G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   0
         Left            =   9960
         TabIndex        =   68
         Top             =   4290
         Width           =   150
      End
      Begin VB.Label Label2 
         Caption         =   "H A E M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   3
         Left            =   9960
         TabIndex        =   67
         Top             =   2430
         Width           =   120
      End
      Begin VB.Label Label2 
         Caption         =   "B I O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   2
         Left            =   9960
         TabIndex        =   66
         Top             =   390
         Width           =   120
      End
      Begin VB.Label Label11 
         Caption         =   "AddOns Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10140
         TabIndex        =   57
         Top             =   120
         Width           =   2580
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   435
         Index           =   12
         Left            =   9300
         Shape           =   3  'Circle
         Top             =   300
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   12
         Left            =   9540
         TabIndex        =   53
         Top             =   420
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   435
         Index           =   11
         Left            =   6120
         Shape           =   3  'Circle
         Top             =   300
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   11
         Left            =   6360
         TabIndex        =   52
         Top             =   420
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   51
         Top             =   420
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   435
         Index           =   10
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   300
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   7
         Left            =   9300
         Shape           =   3  'Circle
         Top             =   1800
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   7
         Left            =   9540
         TabIndex        =   46
         Top             =   1920
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   8
         Left            =   9300
         Shape           =   3  'Circle
         Top             =   3780
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   8
         Left            =   9540
         TabIndex        =   45
         Top             =   3900
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   9
         Left            =   9300
         Shape           =   3  'Circle
         Top             =   5424
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   9
         Left            =   9540
         TabIndex        =   44
         Top             =   5550
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label2 
         Caption         =   "B I O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1932
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Urgent Samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3690
         TabIndex        =   42
         Top             =   5940
         Width           =   2745
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   6
         Left            =   6360
         TabIndex        =   40
         Top             =   5550
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   6
         Left            =   6120
         Shape           =   3  'Circle
         Top             =   5424
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   5
         Left            =   6360
         TabIndex        =   39
         Top             =   3900
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   5
         Left            =   6120
         Shape           =   3  'Circle
         Top             =   3780
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   4
         Left            =   6360
         TabIndex        =   38
         Top             =   1920
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   4
         Left            =   6120
         Shape           =   3  'Circle
         Top             =   1800
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Not Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   6720
         TabIndex        =   36
         Top             =   120
         Width           =   792
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Not Validated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   3540
         TabIndex        =   35
         Top             =   120
         Width           =   984
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   32
         Top             =   5550
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   31
         Top             =   5820
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   1
         Left            =   3120
         TabIndex        =   30
         Top             =   3900
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   3
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   5424
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   435
         Index           =   2
         Left            =   30
         Shape           =   3  'Circle
         Top             =   5700
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   1
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   3780
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Shape sCircle 
         BorderWidth     =   2
         Height          =   432
         Index           =   0
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   1800
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Index           =   0
         Left            =   3120
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Auto Validation Failures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   5940
         Width           =   2748
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Phone Alerts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13140
         TabIndex        =   14
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Outstanding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   300
         TabIndex        =   5
         Top             =   120
         Width           =   912
      End
      Begin VB.Label Label2 
         Caption         =   "H A E M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1812
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   120
      End
      Begin VB.Label Label1 
         Caption         =   "C O A G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   4260
         Width           =   150
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   7170
      TabIndex        =   1
      Top             =   750
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8880
      Width           =   16860
      _ExtentX        =   29739
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "20/06/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "18:58"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   23283
            Text            =   "Custom Software Ltd"
            TextSave        =   "Custom Software Ltd"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmMain.frx":6AAD
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmMain.frx":6D83
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblTestSystem 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caution - Test System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   13380
      TabIndex        =   48
      Top             =   60
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblTransfer 
      BackColor       =   &H000000FF&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13380
      TabIndex        =   47
      Top             =   540
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9450
      TabIndex        =   19
      Top             =   420
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Show details for last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7170
      TabIndex        =   17
      Top             =   420
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8895
      Top             =   -60
      Width           =   480
   End
   Begin VB.Image imgSearch 
      Height          =   510
      Index           =   3
      Left            =   5700
      Picture         =   "frmMain.frx":7059
      Stretch         =   -1  'True
      ToolTipText     =   "Search by Name & Date of Birth"
      Top             =   240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Index           =   2
      Left            =   5190
      Picture         =   "frmMain.frx":E55B
      ToolTipText     =   "Search by Date of Birth"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Index           =   1
      Left            =   4650
      Picture         =   "frmMain.frx":EB95
      ToolTipText     =   "Search by Chart"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSearch 
      Height          =   510
      Index           =   0
      Left            =   4050
      Picture         =   "frmMain.frx":F1CF
      Stretch         =   -1  'True
      ToolTipText     =   "Search by Name"
      Top             =   240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mLogOn 
         Caption         =   "&Log On"
      End
      Begin VB.Menu mLogOff 
         Caption         =   "Log &Off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewWards 
         Caption         =   "View &Ward Enquiries"
      End
      Begin VB.Menu mnuResetLastUsed 
         Caption         =   "&Reset 'Last Used'"
         Enabled         =   0   'False
         Begin VB.Menu mnuResetLastUsedGeneral 
            Caption         =   "&General Lab"
         End
         Begin VB.Menu mnuResetLastUsedMicro 
            Caption         =   "&Microbiology"
         End
         Begin VB.Menu mnuMaintenance 
            Caption         =   "&Record Maintenance"
         End
      End
      Begin VB.Menu mnuViewArchives 
         Caption         =   "&View Archives"
      End
      Begin VB.Menu mnuMergeClinsWards 
         Caption         =   "&Merge Clinicians/Wards"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnull 
         Caption         =   "-"
      End
      Begin VB.Menu exitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu mViewEdit 
         Caption         =   "&View/Edit Results"
         Begin VB.Menu mnuEditAll 
            Caption         =   "&General"
         End
         Begin VB.Menu mnuEditMicrobiology 
            Caption         =   "&Microbiology"
         End
         Begin VB.Menu mnuEditSemen 
            Caption         =   "&Semen Analysis"
         End
      End
      Begin VB.Menu morder 
         Caption         =   "&Order"
      End
      Begin VB.Menu mBatches 
         Caption         =   "&Batches"
         Begin VB.Menu mnuBatchHaem 
            Caption         =   "&Haematology"
         End
         Begin VB.Menu mnuBatchMicro 
            Caption         =   "&Microbiology"
            Begin VB.Menu mnuBatchOccult 
               Caption         =   "&Occult Blood"
            End
         End
         Begin VB.Menu mnuBatchExt 
            Caption         =   "&Externals"
         End
      End
   End
   Begin VB.Menu msearch 
      Caption         =   "&Search"
      Enabled         =   0   'False
      Begin VB.Menu msearchmore 
         Caption         =   "&Name"
         Index           =   0
      End
      Begin VB.Menu msearchmore 
         Caption         =   "&Chart"
         Index           =   1
      End
      Begin VB.Menu msearchmore 
         Caption         =   "&Date of Birth"
         Index           =   2
      End
      Begin VB.Menu msearchmore 
         Caption         =   "Name && Date of &Birth"
         Index           =   3
      End
   End
   Begin VB.Menu mlists 
      Caption         =   "&Lists"
      Enabled         =   0   'False
      Begin VB.Menu mnuOrderComms 
         Caption         =   "&Order Comms"
         Enabled         =   0   'False
         Begin VB.Menu mnuOCPanels 
            Caption         =   "Panels"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mLocations 
         Caption         =   "Locations"
         Begin VB.Menu mWards 
            Caption         =   "&Wards"
         End
         Begin VB.Menu mListHospitals 
            Caption         =   "&Hospitals"
         End
         Begin VB.Menu mClinicians 
            Caption         =   "&Clinicians"
         End
         Begin VB.Menu mGPs 
            Caption         =   "&G.P.'s"
         End
      End
      Begin VB.Menu mComments 
         Caption         =   "Co&mments"
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Biochemistry"
            Index           =   0
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Haematology"
            Index           =   1
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Coagulation"
            Index           =   2
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Demographics"
            Index           =   3
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Semen"
            Index           =   4
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "Clinical De&tails"
            Index           =   5
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Microbiology"
            Index           =   6
         End
         Begin VB.Menu zSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBioCommentTemplates 
            Caption         =   "Bio Comment &Tempates"
         End
         Begin VB.Menu mnuCoagCommentTemplate 
            Caption         =   "Coag Comment T&emplate"
         End
      End
      Begin VB.Menu mDefaults 
         Caption         =   "&Defaults"
         Begin VB.Menu mnuBGARanges 
            Caption         =   "Blood &Gas"
         End
         Begin VB.Menu mDefaultsBio 
            Caption         =   "&Biochemistry"
            Begin VB.Menu mnuBioControlDefinitions 
               Caption         =   "&QC Names"
            End
            Begin VB.Menu mBarCode 
               Caption         =   "&Bar Codes"
            End
            Begin VB.Menu mPanelsTop 
               Caption         =   "&Panels"
               Begin VB.Menu mPanels 
                  Caption         =   "&Define"
               End
               Begin VB.Menu mPanelBarCodes 
                  Caption         =   "&Barcodes"
               End
            End
         End
         Begin VB.Menu mnuMicroLists 
            Caption         =   "&Microbiology"
            Begin VB.Menu mnuDefaultsMicro 
               Caption         =   "&Microbiology"
            End
            Begin VB.Menu mnuMicroIdent 
               Caption         =   "&Identification"
               Begin VB.Menu mnuMicroIDGram 
                  Caption         =   "&Gram Stains"
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
                  Caption         =   "&Casts"
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
         Begin VB.Menu mnuDefaultsExternal 
            Caption         =   "&External"
            Begin VB.Menu mnuExternalTests 
               Caption         =   "Define &Tests"
            End
            Begin VB.Menu mnuExternalPanels 
               Caption         =   "Define &Panels"
            End
         End
         Begin VB.Menu mnuMisc 
            Caption         =   "&Miscellaneous"
            Begin VB.Menu mnuListErrors 
               Caption         =   "Errors"
            End
            Begin VB.Menu mnuListSampleTypes 
               Caption         =   "Sample Types"
            End
            Begin VB.Menu mnuListSpecimenSources 
               Caption         =   "Specimen Sources"
            End
            Begin VB.Menu mnuListResistanceMarkers 
               Caption         =   "&Resistance Markers"
            End
         End
         Begin VB.Menu mnuPhoneAlerts 
            Caption         =   "&Phone Alerts"
         End
         Begin VB.Menu mnuAutoGenerateComments 
            Caption         =   "&Auto-Generate Comments"
            Begin VB.Menu mnuAutoGenCommentBio 
               Caption         =   "&Biochemistry"
            End
            Begin VB.Menu mnuAutoGenCommentCoag 
               Caption         =   "&Coagulation"
            End
            Begin VB.Menu mnuAutoGenCommentMicro 
               Caption         =   "&Microbiology"
            End
         End
      End
      Begin VB.Menu mPrinters 
         Caption         =   "&Printers"
      End
      Begin VB.Menu mnuSounds 
         Caption         =   "&Sounds"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "System &Options"
      End
   End
   Begin VB.Menu mreports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Begin VB.Menu mnuActivityLog 
         Caption         =   "Activity Log"
      End
      Begin VB.Menu mnuDailyBiochemistry 
         Caption         =   "&Daily Biochemistry"
      End
      Begin VB.Menu mnuReportCollated 
         Caption         =   "&Collated Report"
      End
      Begin VB.Menu mnu24HrUrine 
         Caption         =   "&24 Hour Urine"
      End
      Begin VB.Menu mnuAbnormals 
         Caption         =   "&Abnormals"
         Begin VB.Menu mnuAbnormalBio 
            Caption         =   "&Biochemistry"
         End
         Begin VB.Menu mnuAbnormalHaem 
            Caption         =   "&Haematology"
         End
         Begin VB.Menu mnuAbnormalCoag 
            Caption         =   "&Coagulation"
         End
         Begin VB.Menu mnuAbnormalImm 
            Caption         =   "&Immunology"
         End
      End
      Begin VB.Menu mnuMicroReports 
         Caption         =   "&Microbiology"
         Begin VB.Menu mnuOutstandingMicro 
            Caption         =   "&Outstanding"
         End
         Begin VB.Menu mnuMicroUsage 
            Caption         =   "&Usage"
         End
         Begin VB.Menu mnuMicroUsageByDate 
            Caption         =   "Usage By &Date"
         End
         Begin VB.Menu mnuMicroListDemographicData 
            Caption         =   "List &Demographic Data"
         End
         Begin VB.Menu mnuMicroUnusedSIDs 
            Caption         =   "&Unused Sample ID's"
         End
      End
      Begin VB.Menu mCreatClear 
         Caption         =   "&Creatinine Clearance"
      End
      Begin VB.Menu mnuUPro 
         Caption         =   "&Urinary Protein"
      End
      Begin VB.Menu mworklist 
         Caption         =   "&Worklist"
      End
      Begin VB.Menu mstats 
         Caption         =   "&Statistics"
         Begin VB.Menu mtothaem 
            Caption         =   "Totals for &Haematology"
         End
         Begin VB.Menu mtotbio 
            Caption         =   "Totals for &Biochemistry"
         End
         Begin VB.Menu mnuTotCoag 
            Caption         =   "Totals for &Coagulation"
         End
         Begin VB.Menu mnuTotExt 
            Caption         =   "Totals for &Externals"
         End
      End
      Begin VB.Menu mtests 
         Caption         =   "&Test Count"
      End
      Begin VB.Menu mnuPhoneLog 
         Caption         =   "&Phone Log"
      End
      Begin VB.Menu mnuFaxLog 
         Caption         =   "&Fax Log"
      End
      Begin VB.Menu mnuExternalWorklist 
         Caption         =   "&External Worklist"
      End
   End
   Begin VB.Menu mqc 
      Caption         =   "&Q.C."
      Enabled         =   0   'False
      Begin VB.Menu mnuQCBio 
         Caption         =   "&Biochemistry"
         Begin VB.Menu mnuViewBioQCToday 
            Caption         =   "View &Today"
         End
         Begin VB.Menu mnuViewBioQCHistorical 
            Caption         =   "View &Historical"
         End
      End
      Begin VB.Menu mnuQCHaem 
         Caption         =   "&Haematology"
      End
      Begin VB.Menu mqclimits 
         Caption         =   "&Limits"
         Visible         =   0   'False
      End
      Begin VB.Menu mmeans 
         Caption         =   "&Running Means"
      End
      Begin VB.Menu mnuReagentLotNumbers 
         Caption         =   "Reagent &Lot Numbers"
         Begin VB.Menu mnuReagentLotMonoMalSick 
            Caption         =   "&Monospot"
            Index           =   0
         End
         Begin VB.Menu mnuReagentLotMonoMalSick 
            Caption         =   "Ma&laria"
            Index           =   1
         End
         Begin VB.Menu mnuReagentLotMonoMalSick 
            Caption         =   "&Sickledex"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Begin VB.Menu mBatch 
         Caption         =   "&Batch"
      End
      Begin VB.Menu mG 
         Caption         =   "&Glucose"
         Begin VB.Menu mglucose 
            Caption         =   "By &Date"
         End
         Begin VB.Menu mGluByName 
            Caption         =   "By &Name"
         End
      End
   End
   Begin VB.Menu mstock 
      Caption         =   "&Stock"
      Enabled         =   0   'False
   End
   Begin VB.Menu mS 
      Caption         =   "&Statistics"
      Enabled         =   0   'False
      Begin VB.Menu mViewStats 
         Caption         =   "View"
      End
      Begin VB.Menu mSetSourceNames 
         Caption         =   "Set Source Names"
      End
      Begin VB.Menu mnuStatsMicro 
         Caption         =   "&Microbiology"
      End
      Begin VB.Menu mnuSuperStats 
         Caption         =   "&By GP, Ward etc"
      End
      Begin VB.Menu mnuBioEndoTotals 
         Caption         =   "Bio/Endo Totals"
      End
      Begin VB.Menu mnuStatsExtInt 
         Caption         =   "&External/Internal"
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help"
      Begin VB.Menu mwinhelp 
         Caption         =   "&Windows Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mtechnical 
         Caption         =   "&Technical Assistance"
      End
      Begin VB.Menu mnull1 
         Caption         =   "-"
      End
      Begin VB.Menu mabout 
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

Option Compare Text

Private ClosingFlag As Boolean
Dim SortOrder As Boolean
Dim DashboardCounter As Integer
'Zyam 5-7-24
Dim trackTimer As Integer
'Zyam 5-7-24

'++++ Junaid 24-02-2024
Private Const fcLine_NO = 0
Private Const fcSr = 1
Private Const fcSDate = 2
Private Const fcSID = 3
Private Const fcISONo = 4
Private Const fcISO = 5
Private Const fcRpt = 6
'--- Junaid

Private Sub AdjustLIH()

          Dim tb As Recordset
          Dim sql As String
          Dim LIH As Integer

17750     On Error GoTo AdjustLIH_Error

17760     sql = "SELECT TOP 50 * FROM Masks WHERE LIH IS NULL"
17770     Set tb = New Recordset
17780     RecOpenServer 0, tb, sql
17790     Do While Not tb.EOF
17800         LIH = 0
17810         If tb!H Then LIH = 1
17820         If tb!S Then LIH = 3
17830         If tb!g Then LIH = 5
17840         If tb!l Then LIH = LIH + 30
17850         If tb!J Then LIH = LIH + 300
17860         tb!LIH = LIH
17870         tb.Update
17880         tb.MoveNext
17890     Loop

17900     Exit Sub

AdjustLIH_Error:

          Dim strES As String
          Dim intEL As Integer

17910     intEL = Erl
17920     strES = Err.Description
17930     LogError "frmMain", "AdjustLIH", intEL, strES, sql


End Sub

Private Sub CheckAutoValidation()

          Dim sql As String

17940     On Error GoTo CheckAutoValidation_Error

17950     sql = "SELECT Pass = 1 WHERE NOT EXISTS ( SELECT * FROM HaemResults R, HaemAutoVal A WHERE " & _
              "(A.Parameter = 'WBC' AND ((A.Include = 1 and r.wbc < a.low) or (a.include = 1 and r.wbc >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'RBC' AND ((A.Include = 1 and r.RBC < a.low) or (a.include = 1 and r.rbc >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'Hct' AND ((A.Include = 1 and r.Hct < a.low) or (a.include = 1 and r.Hct >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'Hgb' AND ((A.Include = 1 and r.Hgb < a.low) or (a.include = 1 and r.Hgb >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'MCH' AND ((A.Include = 1 and r.MCH < a.low) or (a.include = 1 and r.MCH >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'MCHC' AND ((A.Include = 1 and r.MCHC < a.low) or (a.include = 1 and r.MCHC >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'MCV' AND ((A.Include = 1 and r.MCV < a.low) or (a.include = 1 and r.MCV >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'Plt' AND ((A.Include = 1 and r.Plt < a.low) or (a.include = 1 and r.Plt >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'MPV' AND ((A.Include = 1 and r.MPV < a.low) or (a.include = 1 and r.MPV >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'PDW' AND ((A.Include = 1 and r.PDW < a.low) or (a.include = 1 and r.PDW >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'PLCR' AND ((A.Include = 1 and r.PLCR < a.low) or (a.include = 1 and r.PLCR >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'BasA' AND ((A.Include = 1 and r.BasA < a.low) or (a.include = 1 and r.BasA >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'BasP' AND ((A.Include = 1 and r.BasP < a.low) or (a.include = 1 and r.BasP >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'EosA' AND ((A.Include = 1 and r.EosA < a.low) or (a.include = 1 and r.EosA >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'EosP' AND ((A.Include = 1 and r.EosP < a.low) or (a.include = 1 and r.EosP >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'MonoA' AND ((A.Include = 1 and r.MonoA < a.low) or (a.include = 1 and r.MonoA >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'MonoP' AND ((A.Include = 1 and r.MonoP < a.low) or (a.include = 1 and r.MonoP >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'NeutA' AND ((A.Include = 1 and r.NeutA < a.low) or (a.include = 1 and r.NeutA >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'NeutP' AND ((A.Include = 1 and r.NeutP < a.low) or (a.include = 1 and r.NeutP >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'LymA' AND ((A.Include = 1 and r.LymA < a.low) or (a.include = 1 and r.LymA >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'LymP' AND ((A.Include = 1 and r.LymP < a.low) or (a.include = 1 and r.LymP >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'RDWSD' AND ((A.Include = 1 and r.RDWSD < a.low) or (a.include = 1 and r.RDWSD >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'RDWCV' AND ((A.Include = 1 and r.RDWCV < a.low) or (a.include = 1 and r.RDWCV >a.high)) and sampleid=1) " & _
              "or (A.Parameter = 'Ret' AND ((A.Include = 1 and r.RetA < a.low) or (a.include = 1 and r.RetA >a.high)) and sampleid=1) )"

17960     Exit Sub

CheckAutoValidation_Error:

          Dim strES As String
          Dim intEL As Integer

17970     intEL = Erl
17980     strES = Err.Description
17990     LogError "frmMain", "CheckAutoValidation", intEL, strES, sql


End Sub

Private Sub CheckPAS()

          Dim tb As Recordset
          Dim sql As String
          Dim S As String
          Dim lngNewDiff As Long
          Dim lngAmdDiff As Long

18000     On Error GoTo CheckPAS_Error

18010     lngNewDiff = 0
18020     lngAmdDiff = 0

          '40    sql = "DECLARE @AmendedHours int, @NewEntryHours int " & _
          '      "SET @AmendedHours = DATEDIFF(hour, (SELECT TOP 1 DateTimeAmended FROM PatientIFs " & _
          '      "                       WHERE NewEntry = 0 ORDER BY DateTimeAmended desc), " & _
          '      "                       getdate()) " & _
          '      "SET @NewEntryHours = DATEDIFF(hour, (SELECT TOP 1 DateTimeAmended FROM PatientIFs " & _
          '      "                       WHERE NewEntry = 1 ORDER BY DateTimeAmended desc), " & _
          '      "                       getdate()) " & _
          '      "SELECT 1,2, @AmendedHours AmendedHours, @NewEntryHours NewEntryHours, " & _
          '      "AmendedFlag = CASE " & _
          '      "  WHEN @AmendedHours > 48 " & _
          '      "  THEN 1 " & _
          '      "  ELSE 0 END, " & _
          '      "NewEntryFlag = CASE " & _
          '      "  WHEN @NewEntryHours > 48 " & _
          '      "  THEN 1 " & _
          '      "  ELSE 0 END"
          '
          '50    Set tb = New Recordset
          '60    RecOpenClient 0, tb, sql
          '
          '
          '70    If tb!AmendedFlag = 1 Or tb!NewEntryFlag = 1 Then
          '80      S = "There may be a problem with the PAS" & vbCrLf & vbCrLf
          '90      If tb!NewEntryFlag = 1 Then
          '100       S = S & "No New Records received for " & Format(tb!NewEntryHours) & " Hours." & vbCrLf & vbCrLf
          '110     End If
          '120     If tb!AmendedFlag = 1 Then
          '130       S = S & "No Amendedments received for " & Format(tb!AmendedHours) & " Hours."
          '140     End If
          '150     iMsg S, vbInformation, , vbRed
          '160   End If
          '
18030     sql = "Select Top 1 DateTimeAmended from PatientIFs " & _
              "where NewEntry = 1 " & _
              "order by DateTimeAmended desc"
18040     Set tb = New Recordset
18050     RecOpenServer 0, tb, sql
18060     If Not tb.EOF Then
18070         lngNewDiff = Abs(DateDiff("H", Now, tb!DateTimeAmended))
18080     End If

18090     sql = "Select Top 1 DateTimeAmended from PatientIFs " & _
              "where NewEntry = 0 " & _
              "order by DateTimeAmended desc"
18100     Set tb = New Recordset
18110     RecOpenServer 0, tb, sql
18120     If Not tb.EOF Then
18130         lngAmdDiff = Abs(DateDiff("H", Now, tb!DateTimeAmended))
18140     End If

18150     If lngNewDiff >= 48 Or lngAmdDiff >= 48 Then
18160         S = "There may be a problem with the PAS" & vbCrLf & vbCrLf
18170         If lngNewDiff >= 48 Then
18180             S = S & "No New Records received for " & Format(lngNewDiff) & " Hours." & vbCrLf & vbCrLf
18190         End If
18200         If lngAmdDiff >= 48 Then
18210             S = S & "No Amendedments received for " & Format(lngAmdDiff) & " Hours."
18220         End If
18230         iMsg S, vbInformation, , vbRed
18240     End If

18250     Exit Sub

CheckPAS_Error:

          Dim strES As String
          Dim intEL As Integer

18260     intEL = Erl
18270     strES = Err.Description
18280     LogError "frmMain", "CheckPAS", intEL, strES, sql

End Sub


Private Sub FillAutoValidation()

          Dim sql As String
          Dim tb As Recordset

18290     On Error GoTo FillAutoValidation_Error

          '20    CheckAutoVal

          '30    lstAutoValFail.Clear

18300     grdAutoValFail.Rows = 2
18310     grdAutoValFail.AddItem ""
18320     grdAutoValFail.RemoveItem 1

          '40    sql = BuildSelectAutoValSQL()
18330     sql = "SELECT * from AutoValFailures"

18340     Set tb = New Recordset
18350     RecOpenServer 0, tb, sql
18360     Do While Not tb.EOF
              '80      If tb!AutoVal = "Failure" Then
              '90        lstAutoValFail.AddItem tb!SampleID

18370         grdAutoValFail.AddItem tb!SampleID & vbTab & tb!Dept & vbTab & _
                  IIf(tb!DemNotValid, "x", "") & vbTab & _
                  IIf(tb!Outstanding, "x", "") & vbTab & _
                  IIf(tb!ResultRange, "x", "") & vbTab & _
                  IIf(tb!DeltaCheck, "x", "") & vbTab & _
                  IIf(tb!Flags, "x", "") & vbTab & _
                  IIf(tb!Age, "x", "") & vbTab & _
                  IIf(tb!Sample24HoursOld, "x", "")

              '100     Else
              '110       sql = "UPDATE HaemResults SET " & _
              '                "Valid = 1, " & _
              '                "Healthlink = 0, " & _
              '                "Operator = 'LIS', " & _
              '                "ValidateTime = '" & Format$(Now, "dd/MMM/yyyy HH:nn:ss") & "' " & _
              '                "WHERE SampleID = '" & tb!SampleID & "'"
              '120       Cnxn(0).Execute Sql
              '130     End If

18380         tb.MoveNext
18390     Loop

18400     If grdAutoValFail.Rows > 2 Then
18410         grdAutoValFail.RemoveItem 1
18420     End If

          '160   sql = "IF OBJECT_ID('tempdb..#TempAutoVal') IS NOT NULL  DROP TABLE #TempAutoVal"
          '170   Cnxn(0).Execute Sql

18430     Exit Sub

FillAutoValidation_Error:

          Dim strES As String
          Dim intEL As Integer

18440     intEL = Erl
18450     strES = Err.Description
18460     LogError "frmMain", "FillAutoValidation", intEL, strES, sql


End Sub

Private Sub FillBioNoResults()

          Dim tb As Recordset
          Dim sql As String

18470     On Error GoTo FillBioNoResults_Error

18480     With gBioNoResults
18490         .Rows = 2
18500         .FixedRows = 1
18510         .Rows = 1
18520     End With
18530     lblCount(0) = "0"
18540     lblCount(10) = "0"
18550     sql = "Select distinct R.SampleID, " & _
              "COALESCE(T.Analyser, '') AS Analyser, " & _
              "CASE datediff(day, datetime, getdate()) " & _
              "  WHEN 0 THEN 0 " & _
              "  WHEN 1 THEN 65535 " & _
              "  ELSE 33023 " & _
              "END AS Colour, " & _
              "CASE " & _
              "  WHEN CAST(SampleDate AS nvarchar(9)) IS NULL THEN 'Not Given' " & _
              "  ELSE CONVERT(nvarchar(9), SampleDate, 3) " & _
              "END SD, " & _
              "COALESCE(D.Ward, '') Ward " & _
              "FROM BioRequests R INNER JOIN Demographics D on D.SampleID = R.SampleID " & _
              "INNER JOIN BioTestDefinitions T ON R.Code = T.Code " & _
              "WHERE R.DateTime > DATEADD(day, -%resultdays, getdate())"
18560     If cmbBioNoResult.List(cmbBioNoResult.ListIndex) <> "All Analysers" Then
18570         sql = sql & " AND T.ANALYSER = '" & cmbBioNoResult.List(cmbBioNoResult.ListIndex) & "'"
18580     End If


          'sql = Replace(sql, "%resultdays", 5000)
18590     sql = Replace(sql, "%resultdays", cmbResultDays)

18600     Set tb = New Recordset
18610     Set tb = Cnxn(0).Execute(sql)
18620     With gBioNoResults
18630         .Visible = False
18640         .Col = 1
18650         Do While Not tb.EOF
18660             .AddItem tb!Analyser & vbTab & tb!SampleID    '& "" & vbTab & _
                  'tb!SD & vbTab & tb!Hosp
                  '      If tb!Colour > 0 Then
                  '          .row = .Rows - 1
                  '          .CellBackColor = tb!Colour
                  '      End If
18670             If IsWardInternal(tb!Ward) = False Then ' If UCase(tb!Ward) = "GP" Then
                      'if internal sample
18680                 lblCount(0) = Val(lblCount(0)) + 1
18690             Else
18700                 .Col = 0
18710                 .row = .Rows - 1
18720                 .CellBackColor = vbGreen
18730                 lblCount(10) = Val(lblCount(10)) + 1
18740             End If
18750             tb.MoveNext
18760         Loop

18770     End With
18780     With gBioNoResults
18790         .Visible = True
18800     End With
18810     ColorCodeCircles
18820     Exit Sub

FillBioNoResults_Error:

          Dim strES As String
          Dim intEL As Integer

18830     intEL = Erl
18840     strES = Err.Description
18850     LogError "frmMain", "FillBioNoResults", intEL, strES, sql
18860     gBioNoResults.Visible = True

End Sub
Private Sub FillCmbBioNoResults()

          Dim tb As Recordset
          Dim sql As String
          Dim DefaultAnylaser As String
          Dim FoundAnylaser As Boolean
          Dim i As Integer
18870     On Error GoTo FillcmbBioNoResults_Error



          'sql = "Select distinct Analyser, " & _
           "COALESCE(T.Analyser, '') AS Analyser, " & _
           "CASE datediff(day, datetime, getdate()) " & _
           "  WHEN 0 THEN 0 " & _
           "  WHEN 1 THEN 65535 " & _
           "  ELSE 33023 " & _
           "END AS Colour, " & _
           "CASE " & _
           "  WHEN CAST(SampleDate AS nvarchar(9)) IS NULL THEN 'Not Given' " & _
           "  ELSE CONVERT(nvarchar(9), SampleDate, 3) " & _
           "END SD, " & _
           "COALESCE(D.Ward, '') Ward " & _
           "FROM BioRequests R INNER JOIN Demographics D on D.SampleID = R.SampleID " & _
           "INNER JOIN BioTestDefinitions T ON R.Code = T.Code " & _
           "WHERE R.DateTime > DATEADD(day, -%resultdays, getdate())"


18880     sql = "Select Analyser " & _
              "FROM BioRequests R INNER JOIN Demographics D on D.SampleID = R.SampleID " & _
              "INNER JOIN BioTestDefinitions T ON R.Code = T.Code " & _
              "WHERE R.DateTime > DATEADD(day, -%resultdays, getdate())" & _
              " Group by Analyser "


18890     sql = Replace(sql, "%resultdays", cmbResultDays)

18900     Set tb = New Recordset
18910     Set tb = Cnxn(0).Execute(sql)
18920     With cmbBioNoResult
18930         .Clear
18940         .AddItem "All Analysers"
18950         Do While Not tb.EOF
18960             .AddItem tb!Analyser
                  '.ItemData = (.NewIndex)
18970             tb.MoveNext
18980         Loop

18990     End With
19000     DefaultAnylaser = GetOptionSetting("DefaultAnalyserBio", "")
19010     If DefaultAnylaser = "" Then
19020         cmbBioNoResult.ListIndex = 0
19030     Else
19040         With cmbBioNoResult
19050             For i = 0 To .ListCount
19060                 If UCase(.List(i)) = UCase(DefaultAnylaser) Then
19070                     FoundAnylaser = True
19080                 End If
19090             Next i

19100             If FoundAnylaser = False Then
19110                 cmbBioNoResult.AddItem (DefaultAnylaser)
19120             End If
19130             cmbBioNoResult = (DefaultAnylaser)
19140         End With
19150     End If

19160     Exit Sub

FillcmbBioNoResults_Error:

          Dim strES As String
          Dim intEL As Integer

19170     intEL = Erl
19180     strES = Err.Description
19190     LogError "frmMain", "FillCmbBioNoResults", intEL, strES, sql


End Sub

Private Sub FillForPhone()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

19200     On Error GoTo FillForPhone_Error

19210     For n = 0 To 2
19220         grdPhone(n).Rows = 2
19230         grdPhone(n).AddItem ""
19240         grdPhone(n).RemoveItem 1

19250         sql = "SELECT D.Ward, P.SampleID, P.Parameter " & _
                  "FROM PhoneAlert P JOIN Demographics D ON " & _
                  "D.SampleID = P.SampleID " & _
                  "WHERE Discipline = '" & _
                  Choose(n + 1, "Biochemistry", "Haematology", "Coagulation", "Microbiology") & "' " & _
                  "AND D.RunDate > DATEADD(day, -%resultdays, getdate())"
19260         sql = Replace(sql, "%resultdays", cmbResultDays)

19270         Set tb = New Recordset
19280         RecOpenClient 0, tb, sql
19290         Do While Not tb.EOF
19300             grdPhone(n).AddItem tb!SampleID & vbTab & tb!Parameter & vbTab & tb!Ward & ""
19310             tb.MoveNext
19320         Loop
19330         If grdPhone(n).Rows > 2 Then
19340             grdPhone(n).RemoveItem 1
19350         End If
19360     Next

19370     Exit Sub

FillForPhone_Error:

          Dim strES As String
          Dim intEL As Integer

19380     intEL = Erl
19390     strES = Err.Description
19400     LogError "frmMain", "FillForPhone", intEL, strES, sql

End Sub

Private Sub FillHaemNoResult()

          Dim tb As Recordset
          Dim sql As String

19410     On Error GoTo FillHaemNoResult_Error

19420     With gHaemNoResults
19430         .Visible = False
              '  .Rows = 2
              '40      .FixedRows = 1
19440         .Rows = 1
19450     End With

19460     sql = "SELECT DISTINCT R.SampleID, " & _
              "CASE " & _
              "  WHEN CAST(SampleDate AS nvarchar(9)) IS NULL THEN 'Not Given' " & _
              "  ELSE CONVERT(nvarchar(9),SampleDate,3) " & _
              "END SD " & _
              "FROM HaemRequests R LEFT JOIN Demographics D ON D.SampleID = R.SampleID " & _
              "WHERE RunDate > DATEADD(day, -%resultdays, getdate())"

          'sql = Replace(sql, "%resultdays", cmbResultDays)


          ''''''   Masood 18-Feb-2016   Fill HaeRequest

19470     sql = sql & vbNewLine & " UNION ALL " & vbNewLine

19480     sql = sql & " SELECT DISTINCT R.SampleID, " & _
              "CASE " & _
              "  WHEN CAST(SampleDate AS nvarchar(9)) IS NULL THEN 'Not Given' " & _
              "  ELSE CONVERT(nvarchar(9),SampleDate,3) " & _
              "END SD " & _
              "FROM HaeRequests R LEFT JOIN Demographics D ON D.SampleID = R.SampleID " & _
              "WHERE RunDate > DATEADD(day, -%resultdays, getdate())"

19490     sql = Replace(sql, "%resultdays", cmbResultDays)




19500     Set tb = New Recordset
19510     Set tb = Cnxn(0).Execute(sql)
19520     With gHaemNoResults
19530         .Col = 1
19540         Do While Not tb.EOF
19550             .AddItem tb!SampleID & vbTab & tb!SD
19560             tb.MoveNext
19570         Loop
              '  If .Rows > 2 Then
              '      .RemoveItem 1
              '  End If
19580         .Visible = True
19590     End With

19600     With gHaemNoResults
19610         lblCount(1) = .Rows - 1
19620         .Visible = True
19630     End With

19640     Exit Sub

FillHaemNoResult_Error:

          Dim strES As String
          Dim intEL As Integer

19650     intEL = Erl
19660     strES = Err.Description
19670     LogError "frmMain", "FillHaemNoResult", intEL, strES, sql
19680     gHaemNoResults.Visible = True

End Sub
Private Sub FillBioNotValid()

          Dim tb As Recordset
          Dim sql As String
          Dim VorP As String

19690     On Error GoTo FillBioNotvalid_Error

19700     With gBioNotValid
19710         .Visible = False
19720         .Rows = 2
19730         .FixedRows = 1
19740         .Rows = 1
19750     End With
19760     lblCount(4) = "0"
19770     lblCount(11) = "0"


          '=========BioPrint
19780     sql = "SELECT DISTINCT CONVERT(numeric,R.SampleID) SampleID, " & _
              "COALESCE(T.Analyser, '') AS AN, " & _
              "COALESCE(D.Ward, '') Ward " & _
              "FROM BioResults AS R INNER JOIN Demographics D on D.SampleID = R.SampleID " & _
              "INNER JOIN BioTestDefinitions T ON R.Code = T.Code WHERE   " & _
              "R.RunDate > DATEADD(day, -%resultdays, getdate()) " & _
              "AND ( R.Valid = 0 ) AND  ISNULL(T.Printable,0)=1  ORDER BY CONVERT(numeric,R.SampleID) "

19790     sql = Replace(sql, "%resultdays", cmbResultDays)
19800     Set tb = New Recordset
19810     RecOpenServer 0, tb, sql
19820     Do While Not tb.EOF
19830         gBioNotValid.AddItem tb!AN & vbTab & tb!SampleID & ""
19840         With gBioNotValid
19850             If IsWardInternal(tb!Ward) = False Then ' If UCase(tb!Ward) = "GP" Then
                      'if internal sample
19860                 lblCount(4) = Val(lblCount(4)) + 1
19870             Else
19880                 .Col = 0
19890                 .row = .Rows - 1
19900                 .CellBackColor = vbGreen
19910                 lblCount(11) = Val(lblCount(11)) + 1
19920             End If
19930         End With
19940         tb.MoveNext
19950     Loop

19960     With gBioNotValid

19970         .Visible = True
19980     End With

19990     ColorCodeCircles

20000     Exit Sub


FillBioNotvalid_Error:

          Dim strES As String
          Dim intEL As Integer

20010     intEL = Erl
20020     strES = Err.Description
20030     LogError "frmMain", "FillBioNotvalid", intEL, strES, sql
20040     gBioNotValid.Visible = True


End Sub
Private Sub FillBioNotPrinted()

          Dim tb As Recordset
          Dim sql As String
          Dim VorP As String

20050     On Error GoTo FillBioNotPrinted_Error

20060     With gBioNotPrinted
20070         .Visible = False
20080         .Rows = 2
20090         .FixedRows = 1
20100         .Rows = 1
20110     End With
20120     lblCount(7) = "0"
20130     lblCount(12) = "0"

          '=========BioPrint
20140     sql = "SELECT DISTINCT CONVERT(numeric,R.SampleID) SampleID, " & _
              "COALESCE(T.Analyser, '') AS AN, " & _
              "COALESCE(D.Ward, '') Ward " & _
              "FROM BioResults AS R JOIN Demographics D on D.SampleID = R.SampleID " & _
              "INNER JOIN BioTestDefinitions T ON R.Code = T.Code WHERE   " & _
              "R.RunDate > DATEADD(day, -%resultdays, getdate()) " & _
              "AND (R.Valid = 1 AND R.Printed = 0) AND  ISNULL(T.Printable,0)=1  ORDER BY CONVERT(numeric,R.SampleID) "

20150     sql = Replace(sql, "%resultdays", cmbResultDays)
20160     Set tb = New Recordset
20170     RecOpenServer 0, tb, sql
20180     Do While Not tb.EOF
20190         gBioNotPrinted.AddItem tb!AN & vbTab & tb!SampleID & ""
20200         With gBioNotPrinted
20210             If IsWardInternal(tb!Ward) = False Then 'If UCase(tb!Ward) = "GP" Then
                      'if internal sample
20220                 lblCount(7) = Val(lblCount(7)) + 1
20230             Else
20240                 .Col = 0
20250                 .row = .Rows - 1
20260                 .CellBackColor = vbGreen
20270                 lblCount(12) = Val(lblCount(12)) + 1
20280             End If
20290         End With
20300         tb.MoveNext
20310     Loop

20320     With gBioNotPrinted
20330         .Visible = True
20340     End With

20350     ColorCodeCircles

20360     Exit Sub

FillBioNotPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

20370     intEL = Erl
20380     strES = Err.Description
20390     LogError "frmMain", "FillBioNotPrinted", intEL, strES, sql
20400     gBioNotPrinted.Visible = True

End Sub

Private Sub ColorCodeCircles()

          Dim i As Integer

20410     On Error GoTo ColorCodeCircles_Error

20420     For i = 0 To 12
20430         sCircle(i).Visible = True
20440         lblCount(i).Visible = True
20450         If i = 0 Or i = 4 Or i = 7 Or i = 10 Or i = 11 Or i = 12 Then       'bio
20460             If lblCount(i) = 0 Then
20470                 sCircle(i).BorderColor = vbGreen
20480             ElseIf lblCount(i) > 0 And lblCount(i) < 200 Then
20490                 sCircle(i).BorderColor = &H80FF&
20500             ElseIf lblCount(i) > 200 Then
20510                 sCircle(i).BorderColor = vbRed
20520             End If
20530         ElseIf i = 1 Or i = 5 Or i = 8 Then    ' Haem
20540             If lblCount(i) = 0 Then
20550                 sCircle(i).BorderColor = vbGreen
20560             ElseIf lblCount(i) > 0 And lblCount(i) < 100 Then
20570                 sCircle(i).BorderColor = &H80FF&
20580             ElseIf lblCount(i) > 100 Then
20590                 sCircle(i).BorderColor = vbRed
20600             End If
20610         ElseIf i = 2 Or i = 6 Or i = 9 Then    ' Coagulation
20620             If lblCount(i) = 0 Then
20630                 sCircle(i).BorderColor = vbGreen
20640             ElseIf lblCount(i) > 0 And lblCount(i) < 30 Then
20650                 sCircle(i).BorderColor = &H80FF&
20660             ElseIf lblCount(i) > 30 Then
20670                 sCircle(i).BorderColor = vbRed
20680             End If
20690         Else
20700             If lblCount(i) = 0 Then
20710                 sCircle(i).BorderColor = vbGreen
20720             ElseIf lblCount(i) > 0 And lblCount(i) < 20 Then
20730                 sCircle(i).BorderColor = &H80FF&
20740             ElseIf lblCount(i) > 20 Then
20750                 sCircle(i).BorderColor = vbRed
20760             End If
20770         End If
20780     Next i

20790     Exit Sub

ColorCodeCircles_Error:

          Dim strES As String
          Dim intEL As Integer

20800     intEL = Erl
20810     strES = Err.Description
20820     LogError "frmMain", "ColorCodeCircles", intEL, strES

End Sub

Private Sub FillCoagRequests()

          Dim sql As String
          Dim tb As Recordset

20830     On Error GoTo FillCoagRequests_Error

20840     sql = "Select distinct SampleID from CoagRequests"
20850     Set tb = New Recordset
20860     Set tb = Cnxn(0).Execute(sql)
20870     lstCoagNoResults.Clear
20880     Do While Not tb.EOF
20890         lstCoagNoResults.AddItem tb!SampleID & ""
20900         tb.MoveNext
20910     Loop
20920     tb.Close
20930     lblCount(3) = lstCoagNoResults.ListCount
20940     ColorCodeCircles
20950     Exit Sub

FillCoagRequests_Error:

          Dim strES As String
          Dim intEL As Integer

20960     intEL = Erl
20970     strES = Err.Description
20980     LogError "frmMain", "FillCoagRequests", intEL, strES, sql

End Sub

Private Sub FillHaemNotPrinted()

          Dim sql As String
          Dim tb As Recordset
          Dim VorP As String

20990     On Error GoTo FillHaemNotPrinted_Error



21000     sql = "Select distinct CONVERT(numeric,SampleID) SampleID from HaemResults where " & _
              "COALESCE( Printed , 0) = 0 And Valid = 1" & _
              "AND RunDate > DATEADD(day, -%resultdays, getdate()) ORDER BY CONVERT(numeric,SampleID) "

21010     sql = Replace(sql, "%resultdays", cmbResultDays)
21020     Set tb = New Recordset
21030     Set tb = Cnxn(0).Execute(sql)
21040     lstHaemNotPrinted.Clear
21050     Do While Not tb.EOF
21060         lstHaemNotPrinted.AddItem tb!SampleID & ""
21070         tb.MoveNext
21080     Loop
21090     tb.Close
21100     lblCount(8) = lstHaemNotPrinted.ListCount
21110     Exit Sub

FillHaemNotPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

21120     intEL = Erl
21130     strES = Err.Description
21140     LogError "frmMain", "FillHaemNotPrinted", intEL, strES, sql

End Sub
Private Sub FillHaemNotValid()

          Dim sql As String
          Dim tb As Recordset
          Dim VorP As String

21150     On Error GoTo FillHaemNotValid_Error

          '    If optNotPrintedVal(0) Then
          '        VorP = "Printed"
          '    Else
21160     VorP = "Valid"
          '    End If

          '30    sql = "Select distinct CONVERT(numeric,SampleID) SampleID from HaemResults where " & _
          '          "COALESCE(" & VorP & ", 0) = 0 " & _
          '          "" & _
          '         "AND RunDate > DATEADD(day, -%resultdays, getdate()) ORDER BY CONVERT(numeric,SampleID) "


          ' Masood
21170     sql = "Select distinct CONVERT(numeric,SampleID) SampleID from HaemResults where " & _
              "COALESCE(" & VorP & ", 0) = 0 " & _
              " AND (COALESCE(cESR, 0) = 0 OR COALESCE(WBC, '') <> '') " & _
              "AND RunDate > DATEADD(day, -%resultdays, getdate()) ORDER BY CONVERT(numeric,SampleID) "

21180     sql = Replace(sql, "%resultdays", cmbResultDays)
21190     Set tb = New Recordset
21200     Set tb = Cnxn(0).Execute(sql)
21210     lstHaemNotValid.Clear
21220     Do While Not tb.EOF
21230         lstHaemNotValid.AddItem tb!SampleID & ""
21240         tb.MoveNext
21250     Loop
21260     tb.Close

21270     lblCount(5) = lstHaemNotValid.ListCount

21280     Exit Sub

FillHaemNotValid_Error:

          Dim strES As String
          Dim intEL As Integer

21290     intEL = Erl
21300     strES = Err.Description
21310     LogError "frmMain", "FillHaemNotValid", intEL, strES, sql

End Sub
Private Sub FillCoagNotPrinted()

          Dim sql As String
          Dim tb As Recordset

21320     On Error GoTo FillCoagNotPrinted_Error

21330     sql = "SELECT DISTINCT CONVERT(numeric,SampleID) SampleID FROM CoagResults WHERE " & _
              " Valid = 1 AND Printed = 0 " & _
              "AND RunDate > DATEADD(day, -%resultdays, getdate()) " & _
              "ORDER BY CONVERT(numeric, SampleID)"

21340     sql = Replace(sql, "%resultdays", cmbResultDays)

21350     Set tb = New Recordset
21360     Set tb = Cnxn(0).Execute(sql)
21370     lstCoagNotPrinted.Clear
21380     Do While Not tb.EOF
21390         lstCoagNotPrinted.AddItem tb!SampleID & ""
21400         tb.MoveNext
21410     Loop

21420     lblCount(9) = lstCoagNotPrinted.ListCount
21430     ColorCodeCircles

21440     Exit Sub

FillCoagNotPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

21450     intEL = Erl
21460     strES = Err.Description
21470     LogError "frmMain", "FillCoagNotPrinted", intEL, strES, sql

End Sub
Private Sub FillCoagNotValid()

          Dim sql As String
          Dim tb As Recordset

21480     On Error GoTo FillCoagNotValid_Error

21490     sql = "SELECT DISTINCT CONVERT(numeric,SampleID) SampleID FROM CoagResults WHERE " & _
              "COALESCE(Valid, 0) = 0 " & _
              "AND RunDate > DATEADD(day, -%resultdays, getdate()) " & _
              "ORDER BY CONVERT(numeric, SampleID)"

21500     sql = Replace(sql, "%resultdays", cmbResultDays)

21510     Set tb = New Recordset
21520     Set tb = Cnxn(0).Execute(sql)
21530     lstCoagNotValid.Clear
21540     Do While Not tb.EOF
21550         lstCoagNotValid.AddItem tb!SampleID & ""
21560         tb.MoveNext
21570     Loop

21580     lblCount(6) = lstCoagNotValid.ListCount
21590     ColorCodeCircles

21600     Exit Sub

FillCoagNotValid_Error:

          Dim strES As String
          Dim intEL As Integer

21610     intEL = Erl
21620     strES = Err.Description
21630     LogError "frmMain", "FillCoagNotValid", intEL, strES, sql

End Sub


Private Sub btnComparison_Click()
21640     If InputBox("Please Enter Password", "Comparison") = "gort77" Then
21650         fmeComparison.Visible = True
21660         dtoFrom.Value = Date
21670         dtpTo.Value = Date
21680         flxComp.Rows = 1
21690         flxComp.row = 0
21700     Else
21710         MsgBox "Invalid Password.", vbInformation
21720     End If
End Sub

Private Sub btnHide_Click()
21730     fmeComparison.Visible = False
End Sub

Private Sub btnProcess_Click()
21740     Call GetComparison
21750     Call GetComparisonWithReport
End Sub

Private Sub chkAutoRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

21760     If chkAutoRefresh.Value Then
21770         FillOutstandinNotPrinted
21780     End If

End Sub

Private Sub FillOutstandinNotPrinted()
21790     On Error GoTo FillOutstandinNotPrinted_Error

21800     FillCmbBioNoResults
          ' FillBioNoResults
21810     FillHaemNoResult
21820     FillCoagRequests

21830     FillBioNotValid
21840     FillHaemNotValid
21850     FillCoagNotValid

21860     FillBioNotPrinted
21870     FillHaemNotPrinted
21880     FillCoagNotPrinted

21890     FillForPhone
21900     FillAutoValidation

21910     Exit Sub

FillOutstandinNotPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

21920     intEL = Erl
21930     strES = Err.Description
21940     LogError "frmMain", "FillOutstandinNotPrinted", intEL, strES

End Sub


Private Sub cmbBioNoResult_Click()
21950     FillBioNoResults
End Sub

Private Sub cmbResultDays_Click()

21960     On Error GoTo cmbResultDays_Click_Error

21970     If chkAutoRefresh.Value Then
21980         FillOutstandinNotPrinted
21990     End If

22000     Exit Sub

cmbResultDays_Click_Error:

          Dim strES As String
          Dim intEL As Integer

22010     intEL = Erl
22020     strES = Err.Description
22030     LogError "frmMain", "cmbResultDays_Click", intEL, strES

End Sub

Private Sub cmdGreenTick_Click(Index As Integer)
          Dim n As Integer
22040     On Error GoTo cmdGreenTick_Click_Error

22050     If grdAddOns.Rows > 2 Then
22060         grdAddOns.Col = 2
22070         For n = 1 To grdAddOns.Rows - 1
22080             grdAddOns.row = n
22090             Set grdAddOns.CellPicture = imgGreenTick.Picture
22100         Next
              '80        cmdOrderAddOns.Enabled = True
22110     End If

22120     Exit Sub

cmdGreenTick_Click_Error:
          Dim strES As String
          Dim intEL As Integer

22130     intEL = Erl
22140     strES = Err.Description
22150     LogError "frmMain", "cmdGreenTick_Click", intEL, strES

End Sub

Private Sub cmdOrderAddOns_Click()


          Dim n As Integer


22160     On Error GoTo cmdOrderAddOns_Click_Error

22170     For n = 1 To grdAddOnsCOAG.Rows - 1
22180         grdAddOnsCOAG.Col = 2
22190         grdAddOnsCOAG.row = n
22200         If grdAddOnsCOAG.CellPicture = imgGreenTick.Picture Then
22210             Call frmScaneSample.ShowRecords(grdAddOnsCOAG.TextMatrix(n, 0), "1")
22220         End If
22230     Next

22240     For n = 1 To grdAddOns.Rows - 1
22250         grdAddOns.Col = 2
22260         grdAddOns.row = n
22270         If grdAddOns.CellPicture = imgGreenTick.Picture Then
22280             Call frmScaneSample.ShowRecords(grdAddOns.TextMatrix(n, 0), "1")
22290         End If
22300     Next

22310     For n = 1 To grdAddOnsHAEM.Rows - 1
22320         grdAddOnsHAEM.Col = 2
22330         grdAddOnsHAEM.row = n
22340         If grdAddOnsHAEM.CellPicture = imgGreenTick.Picture Then
22350             Call frmScaneSample.ShowRecords(grdAddOnsHAEM.TextMatrix(n, 0), "1")
22360         End If
22370     Next
22380     frmScaneSample.Show 1

22390     FillAddOns


22400     Exit Sub

cmdOrderAddOns_Click_Error:
          Dim strES As String
          Dim intEL As Integer

22410     intEL = Erl
22420     strES = Err.Description
22430     LogError "frmMain", "cmdOrderAddOns_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

          Dim n As Integer
          Dim max As Integer

22440     chkAutoRefresh = 0

22450     max = gBioNotPrinted.Rows - 1
22460     If gBioNoResults.Rows - 1 > max Then
22470         max = gBioNoResults.Rows - 1
22480     End If

22490     Printer.Print "Biochemistry"
22500     Printer.Print "Not Printed"; Tab(25); "No Results"
22510     For n = 1 To max
22520         If gBioNotPrinted.Rows - 1 >= n Then
22530             Printer.Print gBioNotPrinted.TextMatrix(n, 0); Tab(10);
22540             Printer.Print gBioNotPrinted.TextMatrix(n, 1);
22550         End If
22560         If gBioNoResults.Rows - 1 >= n Then
22570             Printer.Print Tab(25); gBioNoResults.TextMatrix(n, 0); Tab(35);
22580             Printer.Print gBioNoResults.TextMatrix(n, 1);
22590         End If
22600         Printer.Print
22610     Next

22620     Printer.Print
22630     Printer.Print

22640     Printer.Print "Haematology"
22650     Printer.Print "Not Printed"
22660     For n = 0 To lstHaemNotPrinted.ListCount - 1
22670         Printer.Print lstHaemNotPrinted.List(n)
22680     Next

22690     Printer.Print
22700     Printer.Print

22710     max = lstCoagNotPrinted.ListCount - 1
22720     If lstCoagNoResults.ListCount - 1 > max Then
22730         max = lstCoagNoResults.ListCount - 1
22740     End If
22750     Printer.Print "Coagulation"
22760     Printer.Print "Not Printed"; Tab(25); "No Results"
22770     For n = 0 To max
22780         If lstCoagNotPrinted.ListCount - 1 >= n Then
22790             Printer.Print lstCoagNotPrinted.List(n);
22800         End If
22810         If lstCoagNoResults.ListCount - 1 >= n Then
22820             Printer.Print Tab(25); lstCoagNoResults.List(n);
22830         End If
22840         Printer.Print
22850     Next

22860     Printer.Print "Printed by " & UserName

22870     Printer.EndDoc

22880     chkAutoRefresh = 1

End Sub

Private Sub cmdRedCross_Click(Index As Integer)
          Dim n As Integer
22890     On Error GoTo cmdRedCross_Click_Error

22900     If grdAddOns.Rows > 2 Then
22910         grdAddOns.Col = 2
22920         For n = 1 To grdAddOns.Rows - 1
22930             grdAddOns.row = n
22940             Set grdAddOns.CellPicture = imgRedCross.Picture
22950         Next
              '80        cmdOrderAddOns.Enabled = False
22960     End If
22970     Exit Sub

cmdRedCross_Click_Error:
          Dim strES As String
          Dim intEL As Integer

22980     intEL = Erl
22990     strES = Err.Description
23000     LogError "frmMain", "cmdRedCross_Click", intEL, strES

End Sub

Private Sub cmdScaneSample_Click()
23010     On Error GoTo cmdScaneSample_Click_Error

          '20    If UserHasAuthority(UserMemberOf, "MainViewUnvalidated") = False Then
          '30        iMsg "You do not have authority to view unvalidated samples" & vbCrLf & "Please contact system administrator"
          '40        Exit Sub
          '50    End If


23020     frmScaneSample.Show 1

23030     Exit Sub

cmdScaneSample_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23040     intEL = Erl
23050     strES = Err.Description
23060     LogError "frmMain", "cmdUnvalidated_Click", intEL, strES
End Sub

Private Sub cmdTransfer_Click()

          Dim sql As String
          Dim RecsAff As Integer
          Dim t As Date

23070 On Error GoTo cmdTransfer_Click_Error

23080 sql = "INSERT INTO [CavanTest].[dbo].[demographics] " & _
          "([SampleID],[Chart],[PatName],[Age],[Sex],[ForHaem],[ForBio],[TimeTaken] " & _
          ",[ForHba1c],[ForFerritin],[ForPSA],[Source],[RunDate],[DoB],[Addr0] " & _
          ",[Addr1],[Ward],[Clinician],[GP],[SampleDate],[HaemComments] " & _
          ",[BioComments],[HaemComments2],[HaemComments3],[HaemComments4] " & _
          ",[ClDetails],[Hospital],[BioComments1],[RooH],[FAXed],[ForCoag] " & _
          ",[ForESR],[NinNumber],[Fasting],[OnWarfarin],[DateTimeDemographics] " & _
          ",[DateTimeHaemPrinted],[DateTimeBioPrinted],[DateTimeCoagPrinted] " & _
          ",[Pregnant],[AandE],[RecDate],[ForImm],[RecordDateTime] " & _
          ",[Operator],[ForBGA],[Category],[ForHisto],[ForCyto],[HistoValid] " & _
          ",[CytoValid],[Mrn],[ForExt],[ForEnd],[ForPgp],[Username],[Urgent] " & _
          ",[Valid],[HYear],[ForMicro],[ForSemen],[SentToEMedRenal] " & _
      ",[AssID],[SurName],[ForeName],[ExtSampleID],[Healthlink]) "
23090 sql = sql & "SELECT " & _
          "[SampleID] , [Chart], [PatName], [Age], [Sex], [ForHaem], [ForBio], [TimeTaken] " & _
          ",[ForHba1c],[ForFerritin],[ForPSA],[Source],[RunDate],[DoB],[Addr0] " & _
          ",[Addr1],[Ward],[Clinician],[GP],[SampleDate],[HaemComments] " & _
          ",[BioComments],[HaemComments2],[HaemComments3],[HaemComments4] " & _
          ",[ClDetails],[Hospital],[BioComments1],[RooH],[FAXed],[ForCoag] " & _
          ",[ForESR],[NinNumber],[Fasting],[OnWarfarin],[DateTimeDemographics] " & _
          ",[DateTimeHaemPrinted],[DateTimeBioPrinted],[DateTimeCoagPrinted] " & _
          ",[Pregnant],[AandE],[RecDate],[ForImm],[RecordDateTime] " & _
          ",[Operator],[ForBGA],[Category],[ForHisto],[ForCyto],[HistoValid] " & _
          ",[CytoValid],[Mrn],[ForExt],[ForEnd],[ForPgp],[Username],[Urgent] " & _
          ", 0,[HYear],[ForMicro],[ForSemen],[SentToEMedRenal] " & _
          ",[AssID],[SurName],[ForeName],[ExtSampleID],[Healthlink] " & _
          "From [Cavan].[dbo].[Demographics] "
23100 sql = sql & "Where Rundate = DateAdd(dd, 0, DateDiff(dd, 0, GETDATE())) " & _
          "AND SampleID NOT IN(SELECT Test.SampleID  FROM [CavanTest].[dbo].[demographics] Test " & _
        "                    JOIN [Cavan].[dbo].[demographics] Live " & _
        "                    ON Test.SampleID = Live.SampleID " & _
        "                    AND Test.RunDate =DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"

23110 Cnxn(0).Execute sql, RecsAff
       
23120 lblTransfer = RecsAff & " Demographic records copied"
23130 lblTransfer.Refresh
23140 t = Now: Do While DateDiff("s", t, Now) < 2: Loop

      '+++ Need To Check
23150 sql = "INSERT INTO [CavanTest].[dbo].[BioResults] " & _
          "([sampleid],[Code],[result],[valid],[printed],[RunTime],[RunDate],[Operator], " & _
          "[Flags],[Units],[SampleType],[Analyser],[Faxed],[Authorised], " & _
          "[Healthlink],[Comment],[DefIndex]) " & _
          "SELECT " & _
          "[SampleID] , [Code], [Result], 0, 0, [RunTime], [Rundate], [Operator], " & _
          "[Flags],[Units],[SampleType],[Analyser],[Faxed],[Authorised], " & _
          "[Healthlink],[Comment],[DefIndex] " & _
          "From [Cavan].[dbo].[BioResults] " & _
          "Where Rundate = DateAdd(dd, 0, DateDiff(dd, 0, GETDATE())) " & _
          "AND SampleID NOT IN(SELECT Test.SampleID  FROM [CavanTest].[dbo].[BioResults] Test " & _
        "                    JOIN [Cavan].[dbo].[BioResults] Live " & _
        "                    ON Test.SampleID = Live.SampleID " & _
        "                    AND Test.RunDate =DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"

23160 Cnxn(0).Execute sql, RecsAff
       
23170 lblTransfer = RecsAff & " Biochemistry records copied"
23180 lblTransfer.Refresh
23190 t = Now: Do While DateDiff("s", t, Now) < 2: Loop


23200 sql = "INSERT INTO [CavanTest].[dbo].[CoagResults] " & _
          "([RunDate],[SampleID],[RunTime],[Code],[Result],[Printed],[Valid],[Units]," & _
          "[UserName],[Authorised],[Released],[FAXed],[Analyser],[Healthlink]) " & _
          "SELECT " & _
          "[Rundate] , [SampleID], [RunTime], [Code], [Result], 0, 0, [Units], " & _
          "[UserName],[Authorised],[Released],[FAXed],[Analyser],[Healthlink] " & _
          "From [Cavan].[dbo].[CoagResults] " & _
          "Where Rundate = DateAdd(dd, 0, DateDiff(dd, 0, GETDATE())) " & _
          "AND SampleID NOT IN(SELECT Test.SampleID  FROM [CavanTest].[dbo].[CoagResults] Test " & _
        "                JOIN [Cavan].[dbo].[CoagResults] Live " & _
        "                    ON Test.SampleID = Live.SampleID " & _
        "                    AND Test.RunDate =DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"

23210 Cnxn(0).Execute sql, RecsAff
       
23220 lblTransfer = RecsAff & " Coagulation records copied"
23230 lblTransfer.Refresh
23240 t = Now: Do While DateDiff("s", t, Now) < 2: Loop


23250 sql = "INSERT INTO [CavanTest].[dbo].[HaemResults] " & _
          "([sampleid],[analysiserror],[negposerror],[posdiff],[posmorph],[poscount],[err_f],[err_r],[ipmessage] " & _
          ",[wbc],[rbc],[hgb],[hct],[mcv],[mch],[mchc],[plt],[lymp],[monop],[neutP],[eosp],[basp],[lyma],[monoa] " & _
          ",[neuta],[eosa],[basa],[rdwcv],[rdwsd],[pdw],[mpv],[plcr],[valid],[printed],[retics],[monospot],[wbccomment] " & _
          ",[cesr],[cretics],[cmonospot],[ccoag],[md0],[md1],[md2],[md3],[md4],[md5],[RunDate],[RunDateTime] " & _
          ",[ESR],[PT],[PTControl],[APTT],[APTTControl],[INR],[FDP],[FIB],[Operator],[FAXed],[Warfarin],[DDimers] " & _
          ",[TransmitTime],[Pct],[WIC],[WOC],[gWB1],[gWB2],[gRBC],[gPlt],[gWIC],[LongError],[cFilm],[RetA] " & _
          ",[RetP],[nrbcA],[nrbcP],[cMalaria],[Malaria],[cSickledex],[Sickledex],[RA],[cRA],[Val1],[Val2] " & _
          ",[Val3],[Val4],[Val5],[gRBCH],[gPLTH],[gPLTF],[gV],[gC],[gS],[DF1],[IRF],[Image],[mi] " & _
          ",[an],[ca],[va],[ho],[he],[ls],[at],[bl],[pp],[nl],[mn],[wp],[ch],[wb],[hdw],[LUCP],[LUCA] " & _
          ",[LI],[MPXI],[ANALYSER],[cAsot],[tAsot],[tRa],[hyp],[rbcf],[rbcg],[mpo],[ig],[lplt],[pclm],[ValidateTime] " & _
          ",[Healthlink],[CD3A],[CD4A],[CD8A],[CD3P],[CD4P],[CD8P],[CD48],[WVF],[AnalyserMessage]) " & _
          "SELECT "
23260 sql = sql & "[SampleID] , [analysiserror], [negposerror], [posdiff], [posmorph], [poscount], [err_f], [err_r], [ipmessage] " & _
          ",[wbc],[rbc],[hgb],[hct],[mcv],[mch],[mchc],[plt],[lymp],[monop],[neutP],[eosp],[basp],[lyma],[monoa] " & _
          ",[neuta],[eosa],[basa],[rdwcv],[rdwsd],[pdw],[mpv],[plcr],0 ,0 ,[retics],[monospot],[wbccomment] " & _
          ",[cesr],[cretics],[cmonospot],[ccoag],[md0],[md1],[md2],[md3],[md4],[md5],[RunDate],[RunDateTime] " & _
          ",[ESR],[PT],[PTControl],[APTT],[APTTControl],[INR],[FDP],[FIB],[Operator],[FAXed],[Warfarin],[DDimers] " & _
          ",[TransmitTime],[Pct],[WIC],[WOC],[gWB1],[gWB2],[gRBC],[gPlt],[gWIC],[LongError],[cFilm],[RetA] " & _
          ",[RetP],[nrbcA],[nrbcP],[cMalaria],[Malaria],[cSickledex],[Sickledex],[RA],[cRA],[Val1],[Val2] " & _
          ",[Val3],[Val4],[Val5],[gRBCH],[gPLTH],[gPLTF],[gV],[gC],[gS],[DF1],[IRF],[Image],[mi] " & _
          ",[an],[ca],[va],[ho],[he],[ls],[at],[bl],[pp],[nl],[mn],[wp],[ch],[wb],[hdw],[LUCP],[LUCA] " & _
          ",[LI],[MPXI],[ANALYSER],[cAsot],[tAsot],[tRa],[hyp],[rbcf],[rbcg],[mpo],[ig],[lplt],[pclm],[ValidateTime] " & _
          ",[Healthlink],[CD3A],[CD4A],[CD8A],[CD3P],[CD4P],[CD8P],[CD48],[WVF],[AnalyserMessage] " & _
          "From [Cavan].[dbo].[HaemResults] "
23270 sql = sql & "Where Rundate = DateAdd(dd, 0, DateDiff(dd, 0, GETDATE())) " & _
          "AND SampleID NOT IN(SELECT Test.SampleID  FROM [CavanTest].[dbo].[HaemResults] Test " & _
        "                    JOIN [Cavan].[dbo].[HaemResults] Live " & _
        "                    ON Test.SampleID = Live.SampleID " & _
        "                    AND Test.RunDate =DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"

23280 Cnxn(0).Execute sql, RecsAff
       
23290 lblTransfer = RecsAff & " Haematology records copied"
23300 lblTransfer.Refresh
23310 t = Now: Do While DateDiff("s", t, Now) < 2: Loop
23320 lblTransfer = ""

23330 Exit Sub

cmdTransfer_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23340 intEL = Erl
23350 strES = Err.Description
23360 LogError "frmMain", "cmdTransfer_Click", intEL, strES, sql

End Sub

Private Sub cmdUnvalidated_Click()

23370 On Error GoTo cmdUnvalidated_Click_Error

23380 If UserHasAuthority(UserMemberOf, "MainViewUnvalidated") = False Then
23390     iMsg "You do not have authority to view unvalidated samples" & vbCrLf & "Please contact system administrator"
23400     Exit Sub
23410 End If


23420 frmNotValidatedPrinted.Show 1

23430 Exit Sub

cmdUnvalidated_Click_Error:

      Dim strES As String
      Dim intEL As Integer

23440 intEL = Erl
23450 strES = Err.Description
23460 LogError "frmMain", "cmdUnvalidated_Click", intEL, strES

End Sub

Private Sub cmdUnvalidatedSamples_Click()

23470 On Error GoTo cmdUnvalidatedSamples_Click_Error

23480 If UserHasAuthority(UserMemberOf, "MainMicroViewUnvalidated") = False Then
23490     iMsg "You do not have authority to view microbiology unvalidated samples" & vbCrLf & "Please contact system administrator"
23500     Exit Sub
23510 End If
23520 frmUnvalidatedSamples.Show 1




23530 Exit Sub

cmdUnvalidatedSamples_Click_Error:
      Dim strES As String
      Dim intEL As Integer

23540 intEL = Erl
23550 strES = Err.Description
23560 LogError "frmMain", "cmdUnvalidatedSamples_Click", intEL, strES

End Sub

Private Sub cmdValidationList_Click()

          Dim f As New frmDemographicValidation
23570 f.Show 1
23580 Set f = Nothing

End Sub



'---------------------------------------------------------------------------------------
' Procedure : cmdResetLabNo_Click
' Author    : XPMUser
' Date      : 10/Feb/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdResetLabNo_Click()

23590 On Error GoTo cmdResetLabNo_Click_Error

      Dim sql As String

23600 sql = "UPDATE Demographics SET LABNO = NULL "
23610 Cnxn(0).Execute (sql)
23620 Call iMsg("Lab no is reset", vbExclamation)

23630 Exit Sub


cmdResetLabNo_Click_Error:

      Dim strES As String
      Dim intEL As Integer

23640 intEL = Erl
23650 strES = Err.Description
23660 LogError "frmMain", "cmdResetLabNo_Click", intEL, strES, sql
End Sub



Private Sub Command1_Click()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdResetLabNo_Click
' Author    : XPMUser
' Date      : 13/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'


Private Sub exitmenu_Click()

          Dim f As Form
          Dim S As String

23670 For Each f In Forms
23680   Debug.Print f.Name
23690 Next

23700 S = "Are you sure you want to close?"
23710 If iMsg(S, vbYesNo + vbQuestion) = vbYes Then
23720   ClosingFlag = True
23730 End If

End Sub

Private Sub Form_Activate()

      Dim Path As String
      Dim strVersion As String

23740 If ClosingFlag Then
23750     Unload Me
23760     Exit Sub
23770 End If

23780 If Not IsIDE Then
23790     Path = CheckNewEXE("NetAcquire")
23800     If Path <> "" Then
23810         Shell App.Path & "\CustomStart.exe NetAcquire"
23820         End
23830         Exit Sub
23840     End If
23850 End If

23860 strVersion = App.Major & "." & App.Minor & "." & App.Revision
23870 Me.Caption = "NetAcquire - Custom Software. Version " & strVersion & " (" & HospName(0) & ")"
23880 lblTestSystem.Visible = False
23890 cmdTransfer.Visible = False
23900 lblTransfer.Visible = False
23910 cmdResetLabNo.Visible = False
23920 If InStr(UCase$(App.Path), "TEST") Then
23930     lblTestSystem.Visible = True
23940     cmdTransfer.Visible = True
23950     lblTransfer.Visible = True
23960     cmdResetLabNo.Visible = True
23970 End If

23980 tmrNotPrinted.Enabled = True

23990 mnuDailyBiochemistry.Visible = True
24000 mnuDailyBiochemistry.Visible = False
      '+++ Junaid 24-02-2024
24010     Call FormatGrid
      '--- Junaid

          'Zyam 6-5-24
        'checkMicroReport
          'Zyam 6-5-24
      'frmAutoGenerateCommentsMicro.Show 1
End Sub
'Zyam 6-5-24
Private Sub checkMicroReport()
          Dim sql As String
          Dim tb As Recordset
          Dim sql1 As String
          Dim tb1 As Recordset
          Dim sampleIds() As String
          Dim i As Integer
          Dim J As Integer
24020     i = 0
24030     J = 0
24040 On Error GoTo checkMicroReport_Click_Error
          
24050 sql = "SELECT sampleid from PrintValidLog WHERE Printed = 0 AND PrintedBy <> '' AND Valid = 1 AND ValidatedDateTime < GETDATE() - 1 AND ValidatedDateTime > DATEADD(months, -3, GETDATE())"
24060 Set tb = New Recordset
24070 RecOpenClient 0, tb, sql
          
24080 Do While Not tb.EOF
24090   sampleIds(i) = ConvertNull(tb!SampleID, "")
24100   i = i + 1
24110     Loop
          
24120 Set tb = Nothing
24130 sql = ""
          
24140     For J = 0 To UBound(sampleIds) - 1
24150   sql1 = "SELECT ward, Clinician, GP FROM Demographics WHERE sampleID = '" & Trim(sampleIds(J)) & "'"
24160   Set tb1 = New Recordset
24170   RecOpenClient 0, tb1, sql1
        
24180   sql = "SELECT * FROM PrintPending WHERE SampleID = '" & Trim(sampleIds(J)) & "'"
24190   RecOpenClient 0, tb, sql
24200   If tb.EOF Then
24210      tb.AddNew
24220   End If

24230   tb!SampleID = Trim(sampleIds(J))
24240   tb!Ward = Trim(tb1!Ward)
24250   tb!Clinician = Trim(tb1!Clinician)
24260   tb!GP = Trim(tb1!GP)
24270   tb!Department = "M"
24280   tb!Initiator = GetValidatorUser(Trim(sampleIds(J)), "M") 'UserName
24290   tb!UsePrinter = ""

24300   tb.Update
24310   J = J + 1


        

24320     Next

24330 Exit Sub

checkMicroReport_Click_Error:

          Dim strES As String
          Dim intEL As Integer

24340 intEL = Erl
24350 strES = Err.Description
24360 LogError "frmMain", "checkMicroReport_Click", intEL, strES


End Sub
'Zyam 6-5-24

Private Sub Form_Click()

          Dim x As Variant
24370 x = CDec(200000000000#)

End Sub




Private Sub gBioNotValid_Click()

          Dim SampleID As String

24380 On Error GoTo gBioNotValid_Click_Error



24390 With gBioNotValid
24400   If .MouseRow = 0 Then
24410       If .MouseCol = 0 Or .MouseCol = 1 Then
24420           If SortOrder Then
24430               .Sort = flexSortGenericAscending
24440           Else
24450               .Sort = flexSortGenericDescending
24460           End If
24470           SortOrder = Not SortOrder
24480       ElseIf .MouseCol = 1 Then

24490           .Sort = 9
24500           SortOrder = Not SortOrder
24510       End If
24520   Else

24530       SampleID = .TextMatrix(.row, 1)
24540       If Val(SampleID) > 0 Then

24550           If UserName = "" Then
24560               iMsg "Not Logged On!", vbExclamation
24570           Else

24580               SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
24590               frmEditAll.StartInDepartment = "B"
24600               frmEditAll.Show 1

24610               FillBioNotValid

24620           End If
24630       End If
24640   End If
24650 End With

24660 Exit Sub

gBioNotValid_Click_Error:

          Dim strES As String
          Dim intEL As Integer

24670 intEL = Erl
24680 strES = Err.Description
24690 LogError "frmMain", "gBioNotValid_Click", intEL, strES

End Sub

Private Sub grdAddOns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      Dim n As Integer

24700    On Error GoTo grdAddOns_MouseUp_Error

      '20    cmdOrderAddOns.Enabled = False
24710 If grdAddOns.MouseRow = 0 Or Trim(grdAddOns.TextMatrix(grdAddOns.MouseRow, 1)) = "" Then Exit Sub


24720 grdAddOns.row = grdAddOns.MouseRow

24730 If grdAddOns.MouseCol = 2 Then
24740     If grdAddOns.CellPicture = imgGreenTick.Picture Then
24750         Set grdAddOns.CellPicture = imgRedCross.Picture
24760     Else
24770         Set grdAddOns.CellPicture = imgGreenTick.Picture
24780     End If
24790 End If

24800 grdAddOns.Col = 2
24810 For n = 1 To grdAddOns.Rows - 1
24820     grdAddOns.row = n
24830     If grdAddOns.CellPicture = imgGreenTick.Picture Then
      '160           cmdOrderAddOns.Enabled = True
24840     End If
24850 Next


24860    Exit Sub

grdAddOns_MouseUp_Error:
      Dim strES As String
      Dim intEL As Integer

24870 intEL = Erl
24880 strES = Err.Description
24890 LogError "frmMain", "grdAddOns_MouseUp", intEL, strES

End Sub

Private Sub grdAddOnsCOAG_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      Dim n As Integer

24900    On Error GoTo grdAddOnsCOAG_MouseUp_Error

      '20    cmdOrderAddOns.Enabled = False
24910 If grdAddOnsCOAG.MouseRow = 0 Or Trim(grdAddOnsCOAG.TextMatrix(grdAddOnsCOAG.MouseRow, 1)) = "" Then Exit Sub


24920 grdAddOnsCOAG.row = grdAddOnsCOAG.MouseRow

24930 If grdAddOnsCOAG.MouseCol = 2 Then
24940     If grdAddOnsCOAG.CellPicture = imgGreenTick.Picture Then
24950         Set grdAddOnsCOAG.CellPicture = imgRedCross.Picture
24960     Else
24970         Set grdAddOnsCOAG.CellPicture = imgGreenTick.Picture
24980     End If
24990 End If

25000 grdAddOnsCOAG.Col = 2
25010 For n = 1 To grdAddOnsCOAG.Rows - 1
25020     grdAddOnsCOAG.row = n
25030     If grdAddOnsCOAG.CellPicture = imgGreenTick.Picture Then
      '160           cmdOrderAddOns.Enabled = True
25040     End If
25050 Next


25060    Exit Sub

grdAddOnsCOAG_MouseUp_Error:
      Dim strES As String
      Dim intEL As Integer

25070 intEL = Erl
25080 strES = Err.Description
25090 LogError "frmMain", "grdAddOnsCOAG_MouseUp", intEL, strES
End Sub

Private Sub grdAddOnsHAEM_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      Dim n As Integer

25100    On Error GoTo grdAddOnsHAEM_MouseUp_Error

      '20    cmdOrderAddOns.Enabled = False
25110 If grdAddOnsHAEM.MouseRow = 0 Or Trim(grdAddOnsHAEM.TextMatrix(grdAddOnsHAEM.MouseRow, 1)) = "" Then Exit Sub


25120 grdAddOnsHAEM.row = grdAddOnsHAEM.MouseRow

25130 If grdAddOnsHAEM.MouseCol = 2 Then
25140     If grdAddOnsHAEM.CellPicture = imgGreenTick.Picture Then
25150         Set grdAddOnsHAEM.CellPicture = imgRedCross.Picture
25160     Else
25170         Set grdAddOnsHAEM.CellPicture = imgGreenTick.Picture
25180     End If
25190 End If

25200 grdAddOnsHAEM.Col = 2
25210 For n = 1 To grdAddOnsHAEM.Rows - 1
25220     grdAddOnsHAEM.row = n
25230     If grdAddOnsHAEM.CellPicture = imgGreenTick.Picture Then
      '160           cmdOrderAddOns.Enabled = True
25240     End If
25250 Next


25260    Exit Sub

grdAddOnsHAEM_MouseUp_Error:
      Dim strES As String
      Dim intEL As Integer

25270 intEL = Erl
25280 strES = Err.Description
25290 LogError "frmMain", "grdAddOnsHAEM_MouseUp", intEL, strES
End Sub

Private Sub grdAutoValFail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

          Dim SampleID As String
          Static SortOrder As Boolean

25300 If UserName = "" Then
25310   iMsg "Not Logged On!", vbExclamation
25320   Exit Sub
25330 End If

25340 If grdAutoValFail.MouseRow > 0 Then

25350   SampleID = Format$(Val(grdAutoValFail.TextMatrix(grdAutoValFail.row, 0)))

25360   SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
25370   frmEditAll.StartInDepartment = grdAutoValFail.TextMatrix(grdAutoValFail.row, 1)
25380   frmEditAll.Show 1

25390   FillAutoValidation

25400 Else

25410   If SortOrder Then
25420       grdAutoValFail.Sort = flexSortGenericAscending
25430   Else
25440       grdAutoValFail.Sort = flexSortGenericDescending
25450   End If
25460   SortOrder = Not SortOrder

25470 End If

End Sub


Private Sub lstCoagNoResults_DblClick()
'Dim SampleID As String
'
'10    On Error GoTo lstCoagNoResults_Click_Error
'
'20    With lstCoagNoResults
'
'30      SampleID = .Text
'40      If Val(SampleID) > 0 Then
'
'50          If UserName = "" Then
'60              iMsg "Not Logged On!", vbExclamation
'70          Else
'
''80              SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
'90              frmEditAll.StartInDepartment = "C"
'                'Zyam 20-04-24
'                frmEditAll.txtSampleID = SampleID
'                Call frmEditAll.txtsampleid_LostFocus
'                'Zyam 20-04-24
'100             frmEditAll.Show 1
'
'
'110             FillCoagRequests
'
'120         End If
'130     End If
'140   End With
'
'150   Exit Sub
'
'lstCoagNoResults_Click_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'160   intEL = Erl
'170   strES = Err.Description
'180   LogError "frmMain", "lstCoagNoResults_DblClick", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lstCoagNotValid_Click
' Author    : XPMUser
' Date      : 09/Dec/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lstCoagNotValid_Click()
25480 On Error GoTo lstCoagNotValid_Click_Error

          Dim SampleID As String
25490 With lstCoagNotValid

25500   SampleID = .Text
25510   If Val(SampleID) > 0 Then

25520       If UserName = "" Then
25530           iMsg "Not Logged On!", vbExclamation
25540       Else

25550           SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
25560           frmEditAll.StartInDepartment = "C"
25570           frmEditAll.Show 1

                'FillHaemNotValid
25580           FillCoagNotPrinted
25590       End If
25600   End If
25610 End With


25620 Exit Sub


lstCoagNotValid_Click_Error:

          Dim strES As String
          Dim intEL As Integer

25630 intEL = Erl
25640 strES = Err.Description
25650 LogError "frmMain", "lstCoagNotValid_Click", intEL, strES

End Sub

Private Sub lstHaemNotValid_Click()

          Dim SampleID As String

25660 On Error GoTo lstHaemNotValid_Click_Error

25670 With lstHaemNotValid

25680   SampleID = .Text
25690   If Val(SampleID) > 0 Then

25700       If UserName = "" Then
25710           iMsg "Not Logged On!", vbExclamation
25720       Else

25730           SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
25740           frmEditAll.StartInDepartment = "H"
25750           frmEditAll.Show 1

25760           FillHaemNotValid

25770       End If
25780   End If
25790 End With

25800 Exit Sub

lstHaemNotValid_Click_Error:

          Dim strES As String
          Dim intEL As Integer

25810 intEL = Erl
25820 strES = Err.Description
25830 LogError "frmMain", "lstHaemNotValid_Click", intEL, strES

End Sub

'Zyam 5-7-24
Private Sub microRepTim_Timer()
25840 If Val(trackTimer) < 3600000 Then
25850         trackTimer = trackTimer + Val(microRepTim.Interval)
25860     Else
25870         checkMicroReport
25880         trackTimer = 0
25890     End If
          
End Sub
'Zyam 5-7-24

Private Sub mnuActivityLog_Click()
25900 frmActivityLog.Show 1
End Sub

Private Sub mnuAutoGenCommentBio_Click()

25910 With frmAutoGenerateComments
25920   .Discipline = "Biochemistry"
25930   .Show 1
25940 End With

End Sub

Private Sub mnuAutoGenCommentCoag_Click()

25950 With frmAutoGenerateComments
25960   .Discipline = "Coagulation"
25970   .Show 1
25980 End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAutoGenCommentMicro_Click
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAutoGenCommentMicro_Click()
25990     On Error GoTo mnuAutoGenCommentMicro_Click_Error


26000     With frmAutoGenerateCommentsMicro
26010         .Show 1
26020     End With


26030     Exit Sub


mnuAutoGenCommentMicro_Click_Error:

          Dim strES As String
          Dim intEL As Integer

26040     intEL = Erl
26050     strES = Err.Description
26060     LogError "frmMain", "mnuAutoGenCommentMicro_Click", intEL, strES
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

26070 On Error GoTo mnuChangePassword_Click_Error

26080 Current = iBOX("Enter your current Password", , , True)
26090 sql = "SELECT * FROM Users WHERE " & _
          "Name = '" & AddTicks(UserName) & "' " & _
          "AND Password = '" & AddTicks(Current) & "' "
26100 If GetOptionSetting("LogOnUpperLower", False) Then
26110   sql = sql & "COLLATE SQL_Latin1_General_CP1_CS_AS"
26120 End If
26130 Set tb = New Recordset
26140 RecOpenServer 0, tb, sql
26150 If Not tb.EOF Then

26160   NewPass = iBOX("Enter new password", , , True)

26170   MinLength = Val(GetOptionSetting("LogOnMinPassLen", "1"))
26180   If Len(NewPass) < MinLength Then
26190       iMsg "Passwords must have a minimum of " & Format(MinLength) & " characters!", vbExclamation
26200       Exit Sub
26210   End If

26220   If GetOptionSetting("LogOnUpperLower", False) Then
26230       If AllLowerCase(NewPass) Or AllUpperCase(NewPass) Then
26240           iMsg "Passwords must have a mixture of UPPER CASE and lower case letters!", vbExclamation
26250           Exit Sub
26260       End If
26270   End If

26280   If GetOptionSetting("LogOnNumeric", False) Then
26290       If Not ContainsNumeric(NewPass) Then
26300           iMsg "Passwords must contain a numeric character!", vbExclamation
26310           Exit Sub
26320       End If
26330   End If

26340   If GetOptionSetting("LogOnAlpha", False) Then
26350       If Not ContainsAlpha(NewPass) Then
26360           iMsg "Passwords must contain an alphabetic character!", vbExclamation
26370           Exit Sub
26380       End If
26390   End If

26400   AllowReUse = GetOptionSetting("PasswordReUse", "No")
26410   If AllowReUse = "No" Then
26420       If PasswordHasBeenUsed(NewPass) Then
26430           iMsg "Password has been used!", vbExclamation
26440           Exit Sub
26450       End If
26460   End If

26470   Confirm = iBOX("Confirm password", , , True)

26480   If NewPass <> Confirm Then
26490       iMsg "Passwords don't match!", vbExclamation
26500       Exit Sub
26510   End If

26520   Cnxn(0).Execute sql

26530   PasswordExpiry = Val(GetOptionSetting("PasswordExpiry", "90"))

26540   sql = "UPDATE Users SET " & _
              "PassWord = '" & NewPass & "', " & _
              "PassDate = '" & Format$(Now + PasswordExpiry, "dd/MMM/yyyy") & "', " & _
              "ExpiryDate = '" & Format$(Now + PasswordExpiry, "dd/MMM/yyyy") & "' " & _
              "WHERE " & _
              "Name = '" & AddTicks(UserName) & "'"
26550   Cnxn(0).Execute sql

26560   iMsg "Your Password has been changed.", vbInformation

26570 End If

26580 Exit Sub

mnuChangePassword_Click_Error:

          Dim strES As String
          Dim intEL As Integer

26590 intEL = Erl
26600 strES = Err.Description
26610 LogError "frmMain", "mnuChangePassword_Click", intEL, strES, sql

End Sub


Private Sub Form_Deactivate()

26620 tmrNotPrinted.Enabled = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

26630 pBar = 0
26640 pbCounter = 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : XPMUser
' Date      : 19/Oct/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()


26650     ReDim Cnxn(0 To 0) As Connection
26660     ReDim CnxnBB(0 To 0) As Connection
26670     ReDim CnxnRemoteBB(0 To 0) As Connection
26680     ReDim HospName(0 To 0) As String
          Dim Con As String
          Dim ConBB As String

26690     If App.PrevInstance Then End

26700     On Error Resume Next

26710     GetINI    'ConnectToDatabase

          'If IsIDE Then
          '  CheckForHaemUpdate
          'End If
      '+++ Junaid 27-05-2024 For Quick Load
      '80        EnsureColumnExists "CoagTestDefinitions", "PrintRefRange", "tinyint NOT NULL DEFAULT 1"
      '
      '90        EnsureColumnExists "GenericResults", "UserName", "nvarchar(50)"
      '100       EnsureColumnExists "GenericResults", "DateTimeOfRecord", "datetime DEFAULT getdate()"
      '110       EnsureColumnExists "GenericResults", "Valid", "tinyint NULL"
      '120       EnsureColumnExists "GenericResults", "ValidatedBy", "nvarchar(50) NULL"
      '130       EnsureColumnExists "GenericResults", "ValidatedDateTime", "datetime NULL"
      '140       EnsureColumnExists "ExternalDefinitions", "UserName", "nvarchar(50)"
      '150       EnsureColumnExists "ExternalDefinitions", "DateTimeOfRecord", "datetime DEFAULT getdate()"
      '160       EnsureColumnExists "ExternalDefinitions", "Department", "nvarchar(50)"
      '170       EnsureColumnExists "eAddress", "ListOrder", "int"
      '180       EnsureColumnExists "eAddress", "DateTimeOfRecord", "datetime DEFAULT getdate()"
      '190       EnsureColumnExists "eAddress", "UserName", "nvarchar(50)"
      '200       EnsureColumnExists "Sensitivities", "DateTimeOfRecord", "datetime DEFAULT getdate()"
      '
      '210       EnsureColumnExists "Reports", "Hidden", "tinyint NOT NULL DEFAULT 0"
      '
      '220       EnsureColumnExists "FaecesResults50", "ValidatedDateTime", "datetime NULL"
      '230       EnsureColumnExists "FaecesResults50", "PrintedDateTime", "datetime NULL"
      '240       EnsureColumnExists "FaecesResults50", "ValidatedBy", "nvarchar(50) NULL"
      '250       EnsureColumnExists "FaecesResults50", "PrintedBy", "nvarchar(50) NULL"
      '
      '260       EnsureColumnExists "BioTestDefinitions", "HealthLinkPanel", "nvarchar(50) NULL"
      '
      '270       EnsureColumnExists "BioRequests", "AddOn", "bit NULL"
      '280       EnsureColumnExists "PhoneLog", "Direction", "nvarchar(50) NULL"
      '
      '290       EnsureColumnExists "Antibiotics", "ViewInGrid", "bit NOT NULL DEFAULT 1"
      '300       EnsureColumnExists "GPs", "PrintReport", "bit NOT NULL DEFAULT 1"
      '310       EnsureColumnExists "GPs", "Interim", "bit NOT NULL DEFAULT 0"

      '320       EnsureColumnExists "MedibridgeResults", "Counter", "numeric(18, 0) IDENTITY(1,1) NOT NULL"
      '330       EnsureColumnExists "MedibridgeResults", "Analyte", "nvarchar(50) NULL"
      '340       EnsureColumnExists "MedibridgeResults", "Chart", "nvarchar(50) NULL"
      '350       EnsureColumnExists "MedibridgeResults", "FileName", "nvarchar(500) NULL"
      '360       EnsureColumnExists "MedibridgeResults", "Units", "nvarchar(50) NULL"
      '
      '370       EnsureColumnExists "Demographics", "MicroHealthLinkReleaseTime", "datetime NULL"
      '380       EnsureColumnExists "DemographicsAudit", "MicroHealthLinkReleaseTime", "datetime NULL"
      '390       EnsureColumnExists "ExternalDefinitions", "BiomnisCode", "nvarchar(50) NULL"
      '400       EnsureColumnExists "AnalyserTestCodeMapping", "Department", "nvarchar(50) NULL"
      '
      '          ' Masood - 26-07-2014
      '410       EnsureColumnExists "BIOTestDefinitions", "DeltaDaysBackLimit", "Int"
      '420       EnsureColumnExists "BioTestDefinitionsArc", "DeltaDaysBackLimit", "Int"
      '430       EnsureColumnExists "CoagTestDefinitions", "DeltaDaysBackLimit", "Int"
      '440       EnsureColumnExists "EndTestDefinitions", "DeltaDaysBackLimit", "Int"
      '450       EnsureColumnExists "EndTestDefinitionsArc", "DeltaDaysBackLimit", "Int"
      '460       EnsureColumnExists "HaemTestDefinitions", "DeltaDaysBackLimit", "Int"
      '          'EnsureColumnExists "HBIETestDefinitions", "DeltaDaysBackLimit", "Int"
      '470       EnsureColumnExists "ImmTestDefinitions", "DeltaDaysBackLimit", "Int"
      '480       EnsureColumnExists "ImmTestDefinitionsArc", "DeltaDaysBackLimit", "Int"
      '          ' Masood - 26-07-2014
      '          'Ensure options exist or listentry exists
      '490       EnsureListEntryExists "B", "Biomnis", "AC"

          ' Masood - 12-08-2014
      '500       EnsureColumnExists "CoagResults", "ValidateTime", "datetime NULL"
      '510       EnsureColumnExists "BioResults", "ValidateTime", "datetime NULL"
      '520       EnsureColumnExists "HaemRepeats50", "ValidateTime", "datetime NULL"
      '530       EnsureColumnExists "BioResultsAudit", "ValidateTime", "datetime NULL"
      '540       EnsureColumnExists "BioResultsAudit", "ValidateTime", "datetime NULL"
      '550       EnsureColumnExists "BioRequestsAudit", "ValidateTime", "datetime NULL"
      '560       EnsureColumnExists "BIOTestDefinitions", "HealthLink", "Int"
      '570       EnsureColumnExists "BIOTestDefinitions", "Accredited", "Int"
      '580       EnsureColumnExists "HaemResults50", "ValidateTime", "datetime NULL"
      '          ' Masood - 12-08-2014
      '590       EnsureColumnExists "demographics", "LabNo", "nvarchar(50) NULL"
      '600       EnsureColumnExists "demographicsAudit", "LabNo", "nvarchar(50) NULL"
      '          ' Masood - 02-10-2014
      '610       EnsureColumnExists "BioRepeats", "MedRenal", "tinyint NULL"
      '620       EnsureColumnExists "BioRepeats", "ValidatedDateTime", "datetime NULL"
      '630       EnsureColumnExists "BioResults", "ValidatedDateTime", "datetime NULL"
      '640       EnsureColumnExists "BioRepeats", "ValidateTime", "datetime NULL"
      '          ' Masood - 02-10-2014
      '650       EnsureColumnExists "BioRequests", "GBottle", "Int"

      '660       EnsureColumnExists "HaemResults", "SignOff", "bit NULL"
      '670       EnsureColumnExists "HaemResults", "SignOffBy", "nvarchar(50) NULL"
      '680       EnsureColumnExists "HaemResults", "SignOffDateTime", "DateTime NULL"
      '690       EnsureColumnExists "HaemResults50", "SignOff", "bit NULL"
      '700       EnsureColumnExists "HaemResults50", "SignOffBy", "nvarchar(50) NULL"
      '710       EnsureColumnExists "HaemResults50", "SignOffDateTime", "DateTime NULL"
      '720       EnsureColumnExists "HaemRepeats", "SignOff", "bit NULL"
      '730       EnsureColumnExists "HaemRepeats", "SignOffBy", "nvarchar(50) NULL"
      '740       EnsureColumnExists "HaemRepeats", "SignOffDateTime", "DateTime NULL"
      '750       EnsureColumnExists "HaemRepeats50", "SignOff", "bit NULL"
      '760       EnsureColumnExists "HaemRepeats50", "SignOffBy", "nvarchar(50) NULL"
      '770       EnsureColumnExists "HaemRepeats50", "SignOffDateTime", "DateTime NULL"
      '780       EnsureColumnExists "TempHaem", "SignOff", "bit NULL"
      '790       EnsureColumnExists "TempHaem", "SignOffBy", "nvarchar(50) NULL"
      '800       EnsureColumnExists "TempHaem", "SignOffDateTime", "DateTime NULL"

      '810       EnsureColumnExists "BioResults", "SignOff", "bit NULL"
      '820       EnsureColumnExists "BioResults", "SignOffBy", "nvarchar(50) NULL"
      '830       EnsureColumnExists "BioResults", "SignOffDateTime", "DateTime NULL"
      '840       EnsureColumnExists "BioRepeats", "SignOff", "bit NULL"
      '850       EnsureColumnExists "BioRepeats", "SignOffBy", "nvarchar(50) NULL"
      '860       EnsureColumnExists "BioRepeats", "SignOffDateTime", "DateTime NULL"
      '870       EnsureColumnExists "BioResultsAudit", "SignOff", "bit NULL"
      '880       EnsureColumnExists "BioResultsAudit", "SignOffBy", "nvarchar(50) NULL"
      '890       EnsureColumnExists "BioResultsAudit", "SignOffDateTime", "DateTime NULL"
      '
      '900       EnsureColumnExists "CoagResults", "SignOff", "bit NULL"
      '910       EnsureColumnExists "CoagResults", "SignOffBy", "nvarchar(50) NULL"
      '920       EnsureColumnExists "CoagResults", "SignOffDateTime", "DateTime NULL"
      '930       EnsureColumnExists "CoagRepeats", "SignOff", "bit NULL"
      '940       EnsureColumnExists "CoagRepeats", "SignOffBy", "nvarchar(50) NULL"
      '950       EnsureColumnExists "CoagRepeats", "SignOffDateTime", "DateTime NULL"
      '          '920   EnsureColumnExists "CoagResultsAudit", "SignOff", "bit NULL"
      '          '930   EnsureColumnExists "CoagResultsAudit", "SignOffBy", "nvarchar(50) NULL"
      '          '940   EnsureColumnExists "CoagResultsAudit", "SignOffDateTime", "DateTime NULL"

      '960       EnsureColumnExists "ImmResults", "SignOff", "bit NULL"
      '970       EnsureColumnExists "ImmResults", "SignOffBy", "nvarchar(50) NULL"
      '980       EnsureColumnExists "ImmResults", "SignOffDateTime", "DateTime NULL"
      '990       EnsureColumnExists "ImmRepeats", "SignOff", "bit NULL"
      '1000      EnsureColumnExists "ImmRepeats", "SignOffBy", "nvarchar(50) NULL"
      '1010      EnsureColumnExists "ImmRepeats", "SignOffDateTime", "DateTime NULL"
      '          'EnsureColumnExists "ImmResultsAudit", "SignOff", "bit NULL"
      '          'EnsureColumnExists "ImmResultsAudit", "SignOffBy", "nvarchar(50) NULL"
      '          'EnsureColumnExists "ImmResultsAudit", "SignOffDateTime", "DateTime NULL"
      '
      '
      '
      '1020      EnsureColumnExists "PrintValidLog", "SignOff", "bit NULL"
      '1030      EnsureColumnExists "PrintValidLog", "SignOffBy", "nvarchar(50) NULL"
      '1040      EnsureColumnExists "PrintValidLog", "SignOffDateTime", "DateTime NULL"
      '
      '1050      EnsureColumnExists "PrintValidLogArc", "SignOff", "bit NULL"
      '1060      EnsureColumnExists "PrintValidLogArc", "SignOffBy", "nvarchar(50) NULL"
      '1070      EnsureColumnExists "PrintValidLogArc", "SignOffDateTime", "DateTime NULL"
      '1080      EnsureColumnExists "FaecesRequests50", "Analyser", "nvarchar(20) NULL"
      '1090      EnsureColumnExists "FaecesRequests50", "Programmed", "bit NULL"
      '1100      EnsureColumnExists "ViewedReports", "Usercode", "nvarchar(5)"
      '
      '1110      EnsureColumnExists "HaemTestDefinitions", "KnownToAnalyser", "bit NULL"    ' Masood - 04-11-2015
          
      '1120      EnsureColumnExists "PrintPending", "PrintAction", "nvarchar(50) NULL"   ' Masood - 05-11-2015
      '1130      EnsureColumnExists "GPOrders", "Programmed", "bit NULL"
      '1140      EnsureColumnExists "GPOrders", "SampleDate", "datetime NULL"
      '1150      EnsureColumnExists "PhoneAlert", "PhoneAlertDateTime", "DateTime NULL"    ' Masood 07-Oct-2015
      '1160      EnsureColumnExists "Reports", "ReportType", "nvarchar(20) NULL"
          


      '1170      CheckScanViewLogInDb
      '1180      CheckUserProfilesInDb
      '
      '1190      CheckPOCTResultsInDb
      '1200      CheckPOCTPatientLiveInDb
      '1210      CheckPOCTPatientsInDb
      '1220      CheckPOCTPatientTempInDb
      '
      '1230      CheckLIHValuesInDb
      '1240      CheckIncludeAutoValUrineInDb
      '1250      CheckUserRoleInDb
      '
      '1260      CheckGPOrdersInDb
      '1270      CheckGPOrderPatientInDb
      '1280      CheckGpordersProfileInDb
      '
      '
      '          '    EnsureColumnExists "GpordersProfile", "Panel", "bit NULL"
      '1290      CheckHaePanelsInDb
      '1300      CheckDemographicsUniLabNoInDb
          
26720     Entity = IIf(UCase$(HospName(0)) = "CAVAN", "01", "31")
26730     RemoteEntity = IIf(UCase$(HospName(0)) = "CAVAN", "31", "01")

26740     cmbResultDays.ListIndex = 0

26750     AdjustLIH

26760     LoadOptions

26770     FillInterpTable

26780     SetOptions

          'fraUrgent.Visible = sysOptUrgent(0)

26790     CheckAndUpdateLockStatus
26800     CheckAndUpdatePrintedStatus

26810     CheckDemogValidationInDb
26820     CheckBiomnisRequestInDb
26830     CheckConsultantListLogInDb
26840     CheckMicroAutoCommentAlertInDb
26850     trackTimer = 0

26860     Exit Sub


Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

26870     intEL = Erl
26880     strES = Err.Description
26890     LogError "frmMain", "Form_Load", intEL, strES

End Sub
Private Sub UpdateLoggedOnUser()
On Error Resume Next
          Dim sql As String
          Dim MachineName As String

26900 On Error GoTo UpdateLoggedOnUser_Error

26910 MachineName = UCase$(vbGetComputerName())

26920 sql = "IF EXISTS (SELECT * FROM LoggedOnUsers WHERE " & _
        "           MachineName = '" & MachineName & "' " & _
        "           AND AppName = 'NetAcquire') " & _
        "  UPDATE LoggedOnUsers " & _
        "  SET UserName = '" & AddTicks(UserName) & "' " & _
        "  WHERE  MachineName = '" & MachineName & "' " & _
        "  AND AppName = 'NetAcquire'" & _
          "ELSE " & _
        "  INSERT INTO LoggedOnUsers " & _
        "  (MachineName, UserName, AppName) " & _
        "  VALUES " & _
        "  ('" & MachineName & "', " & _
        "   '" & AddTicks(UserName) & "', " & _
        "   'NetAcquire')"
26930 Cnxn(0).Execute sql

26940 Exit Sub

UpdateLoggedOnUser_Error:

          Dim strES As String
          Dim intEL As Integer

26950 intEL = Erl
26960 strES = Err.Description
26970 LogError "frmMain", "UpdateLoggedOnUser", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      '
      'Static oButton As Integer
      'Static oShift As Integer
      'Static Ox As Single
      'Static oY As Single
      '
      'If oButton <> Button Or oShift <> Shift Or Ox <> x Or oY <> y Then
26980 pBar = 0
26990 pbCounter = 0
          'End If
          '
          'oButton = Button
          'oShift = Shift
          'Ox = x
          'oY = y

End Sub

Private Sub Form_Unload(Cancel As Integer)

          Dim f As Form

27000 For Each f In Forms
27010   If f.Name <> Me.Name Then
27020       Unload f
27030   End If
27040 Next

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

27050 pBar = 0
27060 pbCounter = 0

End Sub


Private Sub gBioNoResults_Click()


          Dim SampleID As String

27070 With gBioNoResults
27080   If .MouseRow = 0 Then
27090       If .MouseCol = 0 Or .MouseCol = 3 Then
27100           If SortOrder Then
27110               .Sort = flexSortGenericAscending
27120           Else
27130               .Sort = flexSortGenericDescending
27140           End If
27150           SortOrder = Not SortOrder
27160       ElseIf .MouseCol = 2 Then

27170           .Sort = 9
27180           SortOrder = Not SortOrder
27190       End If
27200   Else

27210       SampleID = .TextMatrix(.row, 1)
27220       If Val(SampleID) > 0 Then

27230           If UserName = "" Then
27240               iMsg "Not Logged On!", vbExclamation
27250           Else

27260               SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
27270               frmEditAll.StartInDepartment = "B"
27280               frmEditAll.Show 1

27290               FillBioNoResults

27300           End If
27310       End If
27320   End If
27330 End With

End Sub



Private Sub gBioNoResults_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

27340 With gBioNoResults
27350   If Not IsDate(.TextMatrix(Row1, .Col)) Then
27360       Cmp = 0
27370       Exit Sub
27380   End If

27390   If Not IsDate(.TextMatrix(Row2, .Col)) Then
27400       Cmp = 0
27410       Exit Sub
27420   End If

27430   d1 = Format(.TextMatrix(Row1, .Col), "General Date")
27440   d2 = Format(.TextMatrix(Row2, .Col), "General Date")
27450 End With

27460 If SortOrder Then
27470   Cmp = Sgn(DateDiff("s", d1, d2))
27480 Else
27490   Cmp = Sgn(DateDiff("s", d2, d1))
27500 End If

End Sub

Private Sub gBioNotPrinted_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

          Dim sql As String
          Dim SampleID As String

27510 On Error GoTo gBioNotPrinted_MouseUp_Error

27520 If UserName = "" Then
27530   iMsg "Not Logged On!", vbExclamation
27540   Exit Sub
27550 End If

27560 With gBioNotPrinted

27570   If .MouseRow = 0 Then
27580       If SortOrder Then
27590           .Sort = flexSortGenericAscending
27600       Else
27610           .Sort = flexSortGenericDescending
27620       End If
27630       SortOrder = Not SortOrder
27640   Else
27650       SampleID = .TextMatrix(.MouseRow, 1)
27660       If Button = vbRightButton Then
27670           If .MouseRow > 0 Then
27680               If iMsg("Mark " & SampleID & " as Printed?", vbQuestion + vbYesNo) = vbYes Then
27690                   sql = "Update BioResults " & _
                              "Set Printed = 1 where " & _
                              "SampleID = '" & SampleID & "'"
27700                   Cnxn(0).Execute (sql)
27710                   FillBioNotPrinted
27720               End If
27730           End If
27740       Else
27750           If Val(SampleID) > 0 Then
27760               If UserName = "" Then
27770                   iMsg "Not Logged On!", vbExclamation
27780               Else
27790                   SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID

27800                   frmEditAll.StartInDepartment = "B"
27810                   frmEditAll.Show 1

27820                   FillBioNotPrinted
27830               End If
27840           End If
                'SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
                'frmEditAll.Show 1
27850       End If
27860   End If
27870 End With

          'FillBioNotPrinted

27880 Exit Sub

gBioNotPrinted_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

27890 intEL = Erl
27900 strES = Err.Description
27910 LogError "frmMain", "gBioNotPrinted_MouseUp", intEL, strES, sql


End Sub


Private Sub gHaemNoResults_Click()
          Dim SampleID As String

27920 With gHaemNoResults
27930   If .MouseRow = 0 Then
27940       If .MouseCol = 0 Then
27950           If SortOrder Then
27960               .Sort = flexSortGenericAscending
27970           Else
27980               .Sort = flexSortGenericDescending
27990           End If
28000           SortOrder = Not SortOrder
28010       ElseIf .MouseCol = 1 Then

28020           .Sort = 9
28030           SortOrder = Not SortOrder
28040       End If
28050   Else

28060       SampleID = .TextMatrix(.row, 0)
28070       If Val(SampleID) > 0 Then

28080           If UserName = "" Then
28090               iMsg "Not Logged On!", vbExclamation
28100           Else

28110               SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
28120               frmEditAll.StartInDepartment = "H"
28130               frmEditAll.Show 1

28140               FillHaemNoResult

28150           End If
28160       End If
28170   End If
28180 End With
End Sub

Private Sub gHaemNoResults_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
          Dim d1 As String
          Dim d2 As String

28190 With gHaemNoResults
28200   If Not IsDate(.TextMatrix(Row1, .Col)) Then
28210       Cmp = 0
28220       Exit Sub
28230   End If

28240   If Not IsDate(.TextMatrix(Row2, .Col)) Then
28250       Cmp = 0
28260       Exit Sub
28270   End If

28280   d1 = Format(.TextMatrix(Row1, .Col), "General Date")
28290   d2 = Format(.TextMatrix(Row2, .Col), "General Date")
28300 End With

28310 If SortOrder Then
28320   Cmp = Sgn(DateDiff("s", d1, d2))
28330 Else
28340   Cmp = Sgn(DateDiff("s", d2, d1))
28350 End If
End Sub

Private Sub grdPhone_Click(Index As Integer)

          Dim SampleID As String

28360 With grdPhone(Index)
28370   If .MouseRow = 0 Then
28380       If SortOrder Then
28390           .Sort = flexSortGenericAscending
28400       Else
28410           .Sort = flexSortGenericDescending
28420       End If
28430       SortOrder = Not SortOrder
28440       Exit Sub
28450   End If
28460   SampleID = .TextMatrix(.row, 0)
28470 End With

28480 If Val(SampleID) > 0 Then

28490   If UserName = "" Then
28500       iMsg "Not Logged On!", vbExclamation
28510   Else
28520       SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
28530       Select Case Index
            Case 0: frmEditAll.StartInDepartment = "B"
28540       Case 1: frmEditAll.StartInDepartment = "H"
28550       Case 2: frmEditAll.StartInDepartment = "C"
28560       End Select
28570       frmEditAll.Show 1
28580       FillForPhone
28590   End If

28600 End If


End Sub

Private Sub grdUrg_Click()

          Dim TempTab As Long

28610 If grdUrg.MouseRow = 0 Then Exit Sub

28620 With frmEditAll
28630   .ClearDemographics
28640   .txtSampleID = grdUrg.TextMatrix(grdUrg.row, 0)
28650   .LoadDemographics
28660   .LoadBiochemistry
28670   .LoadCoagulation
28680   .LoadHaematology
28690   TempTab = IIf(sysOptDefaultTab(0) = "", 0, sysOptDefaultTab(0))
28700   sysOptDefaultTab(0) = Val(grdUrg.Col)
28710   .Show 1
28720   sysOptDefaultTab(0) = TempTab
28730 End With

End Sub

Private Sub Image1_Click()

          Dim PW As String

28740 PW = UCase$(iBOX("Password?", , , True))

28750 If PW = "TEMO" Then
28760   frmUpdatePrinted.Show 1
28770 ElseIf PW = "FREDOL2" Then
28780   frmDuplicates.Show 1
28790 End If

End Sub



Private Sub cmdBio_Click()

28800 frmEditAll.Show 1

End Sub

Private Sub cmdMicro_Click()

          Dim f As Form

28810 Set f = New frmEditMicrobiology
28820 f.Show 1
28830 Unload f
28840 Set f = Nothing

End Sub


Private Sub imgSearch_Click(Index As Integer)

28850 With frmPatHistory
28860   .oFor(Index) = True
28870   .FromEdit = False
28880   .chkShort.Value = 0
28890   .Show 1
28900 End With

End Sub

Private Sub cmdSemen_Click()

28910 If UserHasAuthority(UserMemberOf, "AndrologyTab") = False Then
28920   iMsg "You do not have authority to access semen analysis" & vbCrLf & "Please contact system administrator"
28930   Exit Sub
28940 End If
28950 frmEditSemen.Show 1

End Sub


Private Sub lstAutoValFail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

          Dim SampleID As String

28960 On Error GoTo lstAutoValFail_MouseUp_Error

28970 If Val(lstAutoValFail.Text) = 0 Then
28980   Exit Sub
28990 End If
29000 SampleID = Format$(Val(lstAutoValFail.Text))

29010 If UserName = "" Then
29020   iMsg "Not Logged On!", vbExclamation
29030   Exit Sub
29040 End If

29050 SampleID = lstAutoValFail.Text
29060 SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
29070 frmEditAll.Show 1

29080 FillAutoValidation

29090 Exit Sub

lstAutoValFail_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

29100 intEL = Erl
29110 strES = Err.Description
29120 LogError "frmMain", "lstAutoValFail_MouseUp", intEL, strES

End Sub


Private Sub lstCoagNoResults_Click()

          Dim SampleID As String

29130 On Error GoTo lstCoagNoResults_Click_Error

29140 With lstCoagNoResults

29150   SampleID = .Text
29160   If Val(SampleID) > 0 Then

29170       If UserName = "" Then
29180           iMsg "Not Logged On!", vbExclamation
29190       Else

      '80              SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
29200           frmEditAll.StartInDepartment = "C"
                'Zyam 20-04-24
29210           frmEditAll.txtSampleID = SampleID
29220           Call frmEditAll.txtsampleid_LostFocus
                'Zyam 20-04-24
29230           frmEditAll.Show 1
                

29240           FillCoagRequests

29250       End If
29260   End If
29270 End With

29280 Exit Sub

lstCoagNoResults_Click_Error:

          Dim strES As String
          Dim intEL As Integer

29290 intEL = Erl
29300 strES = Err.Description
29310 LogError "frmMain", "lstCoagNoResults_Click", intEL, strES

End Sub

Private Sub lstCoagNotPrinted_Click()

          Dim SampleID As String

29320 If UserName = "" Then
29330   iMsg "Not Logged On!", vbExclamation
29340   Exit Sub
29350 End If

29360 With lstCoagNotPrinted
29370   SampleID = .Text
29380   SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
29390   frmEditAll.StartInDepartment = "C"
29400   frmEditAll.Show 1
29410 End With

29420 FillCoagNotPrinted

End Sub

Private Sub lstHaemNotPrinted_Click()

          Dim SampleID As String

29430 On Error GoTo lstHaemNotPrinted_Click_Error

29440 If Val(lstHaemNotPrinted.Text) = 0 Then
29450   Exit Sub
29460 End If
29470 SampleID = Format$(Val(lstHaemNotPrinted.Text))

29480 If UserName = "" Then
29490   iMsg "Not Logged On!", vbExclamation
29500   Exit Sub
29510 End If

29520 SampleID = lstHaemNotPrinted.Text
29530 SaveSetting "NetAcquire", "StartUp", "LastUsed", SampleID
29540 frmEditAll.StartInDepartment = "H"
29550 frmEditAll.Show 1

29560 FillHaemNotPrinted

29570 Exit Sub

lstHaemNotPrinted_Click_Error:

          Dim strES As String
          Dim intEL As Integer

29580 intEL = Erl
29590 strES = Err.Description
29600 LogError "frmMain", "lstHaemNotPrinted_Click", intEL, strES

End Sub

Private Sub mabout_Click()
29610 frmAbout.Show 1
      'frmPatientNotePad.Show 1
End Sub

Private Sub mbarcode_Click()

29620 frmBarCodes.Show 1

End Sub

Private Sub mbatch_Click()

29630 frmPrintOptions.Show 1

End Sub

Private Sub mClinicians_Click()

29640 frmClinicians.Show 1

End Sub

Private Sub mnu24HrUrine_Click()

29650 frm24hrUrine.Show 1

End Sub

Private Sub mnuAbnormalBio_Click()

29660 frmBioAbnormals.Show 1

End Sub

Private Sub mnuAbnormalCoag_Click()

29670 iMsg "Not Available"

End Sub


Private Sub mnuAbnormalHaem_Click()

29680 iMsg "Not Available"

End Sub


Private Sub mnuAbnormalImm_Click()

29690 iMsg "Not Available"

End Sub


Private Sub mnuBatchExt_Click()

29700 frmExternalBatch.Show 1

End Sub

Private Sub mnuBatchHaem_Click()

29710 frmASOT.Show 1

End Sub

Private Sub mnuBatchOccult_Click()

29720 frmBatchOccult.Show 1

End Sub



Private Sub mnuBGARanges_Click()

29730 frmBGARanges.Show 1

End Sub

Private Sub mnuBioCommentTemplates_Click()
29740 frmCommentsTemplate.CommentDepartment = "B"
29750 frmCommentsTemplate.Show 1
End Sub

Private Sub mnuBioControlDefinitions_Click()

29760 frmBioControlDefinitions.Show 1

End Sub

Private Sub mnuBioEndoTotals_Click()

29770 frmBioEndoTotals.Show 1

End Sub

Private Sub mnuCoagCommentTemplate_Click()
29780 frmCommentsTemplate.CommentDepartment = "C"
29790 frmCommentsTemplate.Show 1
End Sub

Private Sub mnuCommentList_Click(Index As Integer)

29800 With frmComments
29810   .optType(Index) = True
29820   .Show 1
29830 End With

End Sub

Private Sub mCreatClear_Click()

29840 frmCreatClearance.Show 1

End Sub



Private Sub mnuDailyBiochemistry_Click()

29850 frmListingBioCD4TMImm.Show 1

End Sub

Private Sub mnuEditAll_Click()

29860 frmEditAll.Show 1

End Sub

Private Sub mnuEditMicrobiology_Click()

          Dim f As Form

29870 Set f = New frmEditMicrobiology
29880 f.Show 1
29890 Unload f
29900 Set f = Nothing

End Sub


Private Sub mnuEditSemen_Click()

29910 If UserHasAuthority(UserMemberOf, "AndrologyTab") = False Then
29920   iMsg "You do not have authority to access semen analysis" & vbCrLf & "Please contact system administrator"
29930   Exit Sub
29940 End If
29950 frmEditSemen.Show 1

End Sub


Private Sub mGluByName_Click()

29960 frmGlucoseByName.Show 1

End Sub

Private Sub mglucose_Click()

29970 frmGlucose.Show 1

End Sub

Private Sub mGPs_Click()

29980 frmGPs.Show 1

End Sub

Private Sub mListHospitals_Click()

29990 frmHospital.Show 1

End Sub

Private Sub mLogOff_Click()

30000 AddActivity "", "NetAcquire LogOff", "LogOff", "", "", "", ""

30010 UserName = ""
30020 UserCode = ""

30030 EnableMenus False

30040 UpdateLoggedOnUser



End Sub

Private Sub mLogOn_Click()
On Error Resume Next
30050 GetLogOn

30060 EnableMenus UserCode <> ""

30070 UpdateLoggedOnUser



End Sub

Private Sub EnableMenus(ByVal Enable As Boolean)
On Error Resume Next
      Dim n As Integer

30080 For n = 0 To 3
30090     imgSearch(n).Visible = Enable
30100 Next

30110 cmdValidationList.Enabled = Enable
30120 cmdScaneSample.Enabled = Enable
30130 cmdOrderAddOns.Enabled = Enable
30140 cmdUnvalidated.Enabled = True
30150 cmdUnvalidatedSamples.Enabled = True
30160 fraSelectOrder(0).Enabled = Enable
30170 mLogOn.Enabled = Not Enable
30180 mLogOff.Enabled = Enable
30190 mnuResetLastUsed.Enabled = Enable
30200 mEdit.Enabled = Enable
30210 msearch.Enabled = Enable
30220 mlists.Enabled = False
30230 mreports.Enabled = Enable
30240 mstock.Enabled = Enable
30250 mPrint.Enabled = Enable
30260 mqc.Enabled = Enable
30270 mnuMergeClinsWards.Enabled = False
30280 mS.Enabled = False
30290 cmdBio.Visible = Enable
30300 If sysOptDeptMicro(0) Then
30310     cmdMicro.Visible = Enable
30320     cmdSemen.Visible = Enable
30330 Else
30340     cmdMicro.Visible = False
30350     cmdSemen.Visible = False
30360 End If

30370 If Not Enable Then
30380     UserName = ""
30390     UserCode = ""
30400     UserMemberOf = ""
30410 End If

30420 If UserMemberOf = "Managers" Then
30430     mnuMergeClinsWards.Enabled = True
30440     mlists.Enabled = True
30450     mS.Enabled = True
30460 End If

30470 If UserHasAuthority(UserMemberOf, "MicroListMenu") = False Then
30480     mnuMicroLists.Enabled = False
30490 Else
30500     mnuMicroLists.Enabled = True
30510 End If

30520 If IsIDE Or sysOptOrderComms(0) Then
30530     mnuOrderComms.Enabled = True
30540 End If

30550 StatusBar1.Panels(3).Text = UserName

End Sub


Public Sub GetLogOn()
On Error Resume Next
30560 frmLogOn.Show 1

End Sub


Private Sub mmeans_Click()

30570 frmViewRM.Show 1

End Sub

Private Sub mnuDefaultsMicro_Click()

30580 frmAntibioticLists.Show 1

End Sub

Private Sub mnuExternalPanels_Click()

30590 frmExtPanels.Show 1

End Sub

Private Sub mnuExternalTests_Click()

30600 frmExternalTests.Show 1

End Sub

Private Sub mnuExternalWorklist_Click()

30610 frmExternalWorkList.Show 1

End Sub

Private Sub mnuFaxLog_Click()

30620 frmPhoneLogHistory.PhoneOrFAX = "FAX"
30630 frmPhoneLogHistory.SampleID = ""
30640 frmPhoneLogHistory.Show 1

End Sub

Private Sub mnuListErrors_Click()

30650 With frmListsGeneric
30660   .ListType = "ER"
30670   .ListTypeNames = "Errors"
30680   .ListTypeName = "Error"
30690   .Show 1
30700 End With

End Sub

Private Sub mnuListPrestonCCDA_Click()

30710 With frmListsFaeces
30720   .o(2) = True
30730   .Show 1
30740 End With

End Sub

Private Sub mnuListResistanceMarkers_Click()

30750 With frmListsGeneric
30760   .ListType = "RM"
30770   .ListTypeNames = "Resistance Markers"
30780   .ListTypeName = "Resistance Marker"
30790   .Show 1
30800 End With

End Sub

Private Sub mnuListSampleTypes_Click()

30810 With frmListsGeneric
30820   .ListType = "ST"
30830   .ListTypeName = "Sample Type"
30840   .ListTypeNames = "Sample Types"
30850   .Show 1
30860 End With

End Sub

Private Sub mnuListSMAC_Click()

30870 With frmListsFaeces
30880   .o(1) = True
30890   .Show 1
30900 End With

End Sub

Private Sub mnuListSpecimenSources_Click()

30910 With frmListsGeneric
30920   .ListType = "MB"
30930   .ListTypeNames = "Specimen Sources"
30940   .ListTypeName = "Specimen Source"
30950   .Show 1
30960 End With

End Sub

Private Sub mnuListXLDDCA_Click()

30970 With frmListsFaeces
30980   .o(0) = True
30990   .Show 1
31000 End With

End Sub

Private Sub mnuMaintenance_Click()

31010 If UCase$(iBOX("Password Required", , , True)) <> "TEMO" Then Exit Sub

31020 frmMaintenance.Show 1

End Sub

Private Sub mnuMergeClinsWards_Click()

31030 frmMergeClinsWards.Show 1

End Sub

Private Sub mnuMicroGWQuantity_Click()
31040 With frmMicroLists
31050   .o(10).Value = True
31060   .Show 1
31070 End With
End Sub

Private Sub mnuMicroIDGram_Click()

31080 With frmMicroLists
31090   .o(3).Value = True
31100   .Show 1
31110 End With

End Sub

Private Sub mnuMicroListDemographicData_Click()

31120 frmMicroListDemographics.Show 1

End Sub

Private Sub mnuMicroUnusedSIDs_Click()

31130 frmMicroUnusedSampleIDs.Show 1

End Sub

Private Sub mnuMicroUsage_Click()

31140 frmMicroUsage.Show 1

End Sub

Private Sub mnuMicroUsageByDate_Click()

31150 frmMicroByDate.Show 1

End Sub

Private Sub mnuMicroWetPrep_Click()

31160 With frmMicroLists
31170   .o(4).Value = True
31180   .Show 1
31190 End With

End Sub


Private Sub mnuOCPanels_Click()

31200 frmOCPanels.Show 1

End Sub

Private Sub mnuOptions_Click()

31210 frmOption.Show 1

End Sub

Private Sub mnuOutstandingMicro_Click()

31220 frmUnfinished.Show 1

End Sub

Private Sub mnuPhoneAlerts_Click()

31230 frmPhoneAlertLevel.Show 1

End Sub

Private Sub mnuPhoneLog_Click()

31240 frmPhoneLogHistory.SampleID = ""
31250 frmPhoneLogHistory.Show 1

End Sub

Private Sub mnuQCHaem_Click()

31260 frmQCHaem.Show 1

End Sub

Private Sub mnuReagentLotMonoMalSick_Click(Index As Integer)

31270 frmReagentLotNumberReport.optAnalyte(Index) = True
31280 frmReagentLotNumberReport.Show 1

End Sub

Private Sub mnuReportCollated_Click()

31290 frmReportCollated.Show 1

End Sub

Private Sub mnuResetLastUsedMicro_Click()

          Dim LU As String
          Dim NewLU As String

31300 LU = GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")

31310 NewLU = iBOX("Enter new 'Last Used' Number.", , LU)

31320 If Val(NewLU) > 0 Then
31330   SaveSetting "NetAcquire", "StartUp", "LastUsedMicro", Format$(Val(NewLU))
31340   iMsg "'Last Used' Number changed to " & Format$(Val(NewLU)), vbInformation
31350 Else
31360   iMsg "'Last Used' Number not changed!", vbExclamation
31370 End If

End Sub

Private Sub mnuSounds_Click()

31380 frmSound.Show 1

End Sub

Private Sub mnuStatsExtInt_Click()

31390 frmStatsExtInt.Show 1

End Sub

Private Sub mnuStatsMicro_Click()

31400 frmStatsMicro.Show 1

End Sub

Private Sub mnuSuperStats_Click()

31410 frmSuperStats.Show 1

End Sub

Private Sub mnuTotCoag_Click()

31420 frmCoagTotals.Show 1

End Sub

Private Sub mnuTotExt_Click()

31430 frmExternalStats.Show 1

End Sub

Private Sub mnuUPro_Click()

31440 frmUPro.Show 1

End Sub

Private Sub mnuUrineCasts_Click()

31450 With frmMicroLists
31460   .o(5).Value = True
31470   .Show 1
31480 End With

End Sub

Private Sub mnuUrineCrystals_Click()

31490 With frmMicroLists
31500   .o(6).Value = True
31510   .Show 1
31520 End With

End Sub


Private Sub mnuUrineMisc_Click()

31530 With frmMicroLists
31540   .o(7).Value = True
31550   .Show 1
31560 End With

End Sub


Private Sub mnuViewArchives_Click()

31570 frmArchives.Show 1

End Sub

Private Sub mnuViewBioQCHistorical_Click()

31580 frmBioQC.Show 1

End Sub

Private Sub mnuViewBioQCToday_Click()

31590 frmBioTodayQC.Show 1

End Sub

Private Sub mnuViewWards_Click()

31600 frmViewWards.Show 1

End Sub

Private Sub morder_Click()

31610 With frmNewOrder
31620   .FromEdit = False
31630   .Show 1
31640 End With

End Sub

Private Sub mPanelBarCodes_Click()

31650 frmPanelBarCodes.Show 1

End Sub

Private Sub mpanels_Click()

31660 frmPanels.Show 1

End Sub

Private Sub mPrinters_Click()

31670 frmPrinters.Show 1

End Sub

Private Sub mqclimits_Click()

'FVERIFIER.Show 1

End Sub

Private Sub mnuResetLastUsedGeneral_Click()

          Dim LU As String
          Dim NewLU As String

31680 LU = GetSetting("NetAcquire", "StartUp", "LastUsed", "1")

31690 NewLU = iBOX("Enter new 'Last Used' Number.", , LU)

31700 If Val(NewLU) > 0 Then
31710   SaveSetting "NetAcquire", "StartUp", "LastUsed", Format$(Val(NewLU))
31720   iMsg "'Last Used' Number changed to " & Format$(Val(NewLU)), vbInformation
31730 Else
31740   iMsg "'Last Used' Number not changed!", vbExclamation
31750 End If

End Sub

Private Sub msearchmore_Click(Index As Integer)

31760 With frmPatHistory
31770   .oFor(Index) = True
31780   .FromEdit = False
31790   .Show 1
31800 End With

End Sub

Private Sub mSetSourceNames_Click()

31810 frmSetSources.Show 1

End Sub

Private Sub mstock_Click()

31820 frmStockControl.Show 1

End Sub

Private Sub mtechnical_Click()

          Dim S As String

31830 S = "Technical Assistance can be obtained by calling" & vbCrLf & _
        "+353 57 860 1230, +353 87 906 0389" & vbCrLf & _
        "or email info@customsoftware.ie"
31840 iMsg S, vbInformation

End Sub

Private Sub mtests_Click()

31850 frmTests.Show 1

End Sub

Private Sub mtotbio_Click()

31860 frmTotals.Show 1

End Sub

Private Sub mtothaem_Click()

31870 frmTotHaem.Show 1

End Sub

Private Sub mViewStats_Click()

31880 frmStatSources.Show 1

End Sub

Private Sub mWards_Click()

31890 frmWardList.Show 1

End Sub

Private Sub mwinhelp_Click()
'
'Dim dummyval As String
'Dim temp As Integer
'
'dummyval = " "
'temp = winhelp(acquire.hWnd, "c:\windows\help\windows.hlp", HELP_PARTIALKEY, dummyval)

End Sub

Private Sub mworklist_Click()

31900 frmDaily.Show 1

End Sub







Private Sub optImmunology_Click()

End Sub


Private Sub FillUrgent()

      Dim sql As String
      Dim tb As Recordset
      Dim rs As Recordset
      Dim n As Integer
      Dim Found As Boolean

31910 On Error GoTo FillUrgent_Error

31920 If sysOptUrgent(0) Then
31930     grdUrg.Visible = False
31940     grdUrg.Rows = 2
31950     grdUrg.AddItem ""
31960     grdUrg.RemoveItem 1

          '    sql = "update demographics set urgent = 0 where " & _
               '          "recdate < '" & Format(Now - 0.25, "dd/MMM/yyyy hh:mm:ss") & "' and urgent = 1"
          '    Cnxn(0).Execute Sql
31970     sql = "Select * from demographics where urgent = 1 " & _
                "and (recdate = '" & Format(Now - sysOptUrgentRef(0), "dd/MMM/yyyy hh:mm:ss") & "' or recdate " & _
                "is null or recdate <> '') order by rundate desc"
31980     Set tb = New Recordset
31990     Set tb = Cnxn(0).Execute(sql)
32000     Do While Not tb.EOF
32010         grdUrg.AddItem tb!SampleID
32020         grdUrg.row = grdUrg.Rows - 1
32030         sql = "SElect valid from haemresults where sampleid = '" & tb!SampleID & "'"
32040         Set rs = New Recordset
32050         RecOpenServer 0, rs, sql
32060         If Not rs.EOF Then
32070             If Not rs!Valid Then
32080                 grdUrg.Col = 1
32090                 grdUrg.CellBackColor = vbRed
32100             End If
32110         Else
32120             grdUrg.Col = 1
32130             grdUrg.CellBackColor = vbRed
32140         End If
32150         sql = "Select valid from bioresults where sampleid = '" & tb!SampleID & "'"
32160         Set rs = New Recordset
32170         RecOpenServer 0, rs, sql
32180         If Not rs.EOF Then
32190             Do While Not rs.EOF
32200                 If Not rs!Valid Then
32210                     grdUrg.Col = 2
32220                     grdUrg.CellBackColor = vbRed
32230                 End If
32240                 rs.MoveNext
32250             Loop
32260         Else
32270             grdUrg.Col = 2
32280             grdUrg.CellBackColor = vbRed
32290         End If
32300         sql = "SElect valid from coagresults where sampleid = '" & tb!SampleID & "'"
32310         Set rs = New Recordset
32320         RecOpenServer 0, rs, sql
32330         If Not rs.EOF Then
32340             Do While Not rs.EOF
32350                 If Not rs!Valid Then
32360                     grdUrg.Col = 3
32370                     grdUrg.CellBackColor = vbRed
32380                 End If
32390                 rs.MoveNext
32400             Loop
32410         Else
32420             grdUrg.Col = 3
32430             grdUrg.CellBackColor = vbRed
32440         End If



32450         For n = 1 To 6
32460             grdUrg.Col = n
32470             If grdUrg.CellBackColor = vbRed Then
32480                 Found = True
32490             End If
32500         Next
32510         If Found = False Then grdUrg.RemoveItem grdUrg.Rows - 1
32520         Found = False
32530         tb.MoveNext
32540     Loop

32550     If grdUrg.Rows > 2 Then
32560         grdUrg.RemoveItem 1
32570     End If
32580     grdUrg.Visible = True
32590 End If

32600 Exit Sub

FillUrgent_Error:
      Dim strES As String
      Dim intEL As Integer

32610 intEL = Erl
32620 strES = Err.Description
32630 LogError "frmMain", "FillUrgent", intEL, strES

End Sub
Private Sub FillAddOns()

      Dim sql As String
      Dim tb As Recordset
      Dim tbH As Recordset
      Dim tbC As Recordset
      Dim rs As Recordset
      Dim n As Integer
      Dim Found As Boolean



      '20    If sysOptUrgent(0) Then
32640 On Error GoTo fillAddOns_Error
      ' Add ADDOns Bio
32650 grdAddOns.Visible = False
32660 grdAddOns.Rows = 2
32670 grdAddOns.AddItem ""
32680 grdAddOns.RemoveItem 1


      '60    sql = "  SELECT DISTINCT sampleid,departmentid FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1"

      '      Sql = "  SELECT DISTINCT sampleid,left(SampleDate,11) SampleDate FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1 And DepartmentID = 'Bio'"
       
       '+++ Abubaker Siddique 15-11-2023
32690  sql = "  SELECT DISTINCT sampleid,left(SampleDate,11) SampleDate FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1 And DepartmentID = 'Bio'" & _
             "  AND SampleDate > DATEADD(day, -%resultdays, getdate()) "
             
32700  sql = Replace(sql, "%resultdays", cmbResultDays)
       '--- Abubaker Siddique 15-11-2023
       
32710 Set tb = New Recordset
32720 Set tb = Cnxn(0).Execute(sql)
32730 Do While Not tb.EOF
32740     grdAddOns.AddItem tb!SampleID & vbTab & tb!SampleDate & ""
32750     grdAddOns.row = grdAddOns.Rows - 1

32760     tb.MoveNext
32770 Loop



32780 If grdAddOns.Rows > 2 Then
32790     grdAddOns.Col = 2
32800     For n = 1 To grdAddOns.Rows - 1
32810         grdAddOns.row = n
32820         Set grdAddOns.CellPicture = imgRedCross.Picture
32830     Next
32840     grdAddOns.RemoveItem 1
32850 End If
32860 grdAddOns.Visible = True


      ' Add ADDOns HAEM
32870     grdAddOnsHAEM.Visible = False
32880     grdAddOnsHAEM.Rows = 2
32890     grdAddOnsHAEM.AddItem ""
32900     grdAddOnsHAEM.RemoveItem 1
          
          '+++ Abubaker Siddique 15-11-2023
          'sql = "  SELECT DISTINCT sampleid,departmentid FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1"
32910     sql = "  SELECT DISTINCT sampleid,left(SampleDate,11) SampleDate FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1 And DepartmentID = 'HAEM'" & _
          "  AND SampleDate > DATEADD(day, -%resultdays, getdate()) "
             
32920  sql = Replace(sql, "%resultdays", cmbResultDays)
          '--- Abubaker Siddique 15-11-2023
32930     Set tbH = New Recordset
32940     Set tbH = Cnxn(0).Execute(sql)
32950     Do While Not tbH.EOF
32960   grdAddOnsHAEM.AddItem tbH!SampleID & vbTab & tbH!SampleDate & ""
32970   grdAddOnsHAEM.row = grdAddOnsHAEM.Rows - 1
        
32980   tbH.MoveNext
32990     Loop
          
          
          
33000     If grdAddOnsHAEM.Rows > 2 Then
33010   grdAddOnsHAEM.Col = 2
33020   For n = 1 To grdAddOnsHAEM.Rows - 1
33030       grdAddOnsHAEM.row = n
33040       Set grdAddOnsHAEM.CellPicture = imgRedCross.Picture
33050   Next
33060   grdAddOnsHAEM.RemoveItem 1
33070     End If
33080     grdAddOnsHAEM.Visible = True

      ' Add ADDOns COAG
33090     grdAddOnsCOAG.Visible = False
33100     grdAddOnsCOAG.Rows = 2
33110     grdAddOnsCOAG.AddItem ""
33120     grdAddOnsCOAG.RemoveItem 1
          
          '+++ Abubaker Siddique 15-11-2023
          'sql = "  SELECT DISTINCT sampleid,departmentid FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1"
33130     sql = "  SELECT DISTINCT sampleid,left(SampleDate,11) SampleDate FROM ocmrequestDetails WHERE Programmed =0 AND Addon = 1 And DepartmentID = 'COAG'" & _
          "  AND SampleDate > DATEADD(day, -%resultdays, getdate()) "
             
33140     sql = Replace(sql, "%resultdays", cmbResultDays)
          '--- Abubaker Siddique 15-11-2023
          
33150     Set tbC = New Recordset
33160     Set tbC = Cnxn(0).Execute(sql)
33170     Do While Not tbC.EOF
33180   grdAddOnsCOAG.AddItem tbC!SampleID & vbTab & tbC!SampleDate & ""
33190   grdAddOnsCOAG.row = grdAddOnsCOAG.Rows - 1
        
33200   tbC.MoveNext
33210     Loop
          
          
          
33220     If grdAddOnsCOAG.Rows > 2 Then
33230   grdAddOnsCOAG.Col = 2
33240   For n = 1 To grdAddOnsCOAG.Rows - 1
33250       grdAddOnsCOAG.row = n
33260       Set grdAddOnsCOAG.CellPicture = imgRedCross.Picture
33270   Next
33280   grdAddOnsCOAG.RemoveItem 1
33290     End If
33300     grdAddOnsCOAG.Visible = True



33310 Exit Sub

fillAddOns_Error:
      Dim strES As String
      Dim intEL As Integer

33320 intEL = Erl
33330 strES = Err.Description
33340 LogError "frmMain", "fillAddOns", intEL, strES, sql

End Sub

'Private Sub Timer1_Timer()
'    If Val(trackTimer) < 10000 Then
'        trackTimer = trackTimer + Val(microRepTim.Interval)
'    Else
'        MsgBox trackTimer
'        trackTimer = 0
'    End If
'
'End Sub

Private Sub tmrNotPrinted_Timer()

      Dim sql As String

33350 On Error GoTo tmrNotPrinted_Timer_Error


33360 DashboardCounter = DashboardCounter + 1
33370 If mLogOff.Enabled Then
33380     StatusBar1.Panels(3).Text = UserName
33390     pbCounter = pbCounter + 1
33400     If pbCounter < pBar.max Then
33410         pBar = pbCounter
33420     Else
33430         AddActivity "", "NetAcquire Logoff by timeout", "LogOff", "", "", "", ""
33440         EnableMenus False
33450         pbCounter = 0
33460         pBar = 0

33470     End If
33480 Else
33490     StatusBar1.Panels(3).Text = ""
33500 End If

33510 fMainCounter = fMainCounter + 1
33520 fmainImageCounter = fmainImageCounter + 1
33530 If fmainImageCounter > 4 Then
33540     fmainImageCounter = 1
          'Zyam
      '    If Format$("04/05/2001", "dd/mmm/yyyy") <> "04/May/2001" Then
      '        MsgBox "Date/Time Format in" & vbCrLf & _
      '               "International Settings" & vbCrLf & _
      '               "is not set correctly." & vbCrLf & vbCrLf & _
      '               "Cannot proceed!", vbCritical
      '        End
      '    End If
33550 End If

33560 Image1.Picture = ImageList1.ListImages(fmainImageCounter).Picture

33570 If fMainCounter < 5 Then
33580     Exit Sub
33590 End If
33600 fMainCounter = 0

33610 If chkAutoRefresh = 1 And DashboardCounter > 20 Then
          'refresh after every 2 minutes
33620     DashboardCounter = 1
33630     FillOutstandinNotPrinted
33640     FillUrgent
33650     FillAddOns
33660 ElseIf DashboardCounter > 120 Then
33670     DashboardCounter = 1
33680 End If

33690 Exit Sub

tmrNotPrinted_Timer_Error:

      Dim strES As String
      Dim intEL As Integer

33700 intEL = Erl
33710 strES = Err.Description
33720 LogError "frmMain", "tmrNotPrinted_Timer", intEL, strES, sql

End Sub



Private Sub tmrRefresh_Timer()

          Static Counter As Long

          'tmrRefresh.Interval set to 30 seconds

33730 Counter = Counter + 1

33740 If Counter = 30 Then    '15 minutes
33750   CheckPAS
33760   Counter = 0
33770 End If

End Sub

'+++ Junaid 24-02-2024
Private Sub FormatGrid()
33780     On Error GoTo ERROR_FormatGrid
          
33790     flxComp.Rows = 1
33800     flxComp.row = 0
          
33810     flxComp.ColWidth(fcLine_NO) = 100
          
33820     flxComp.TextMatrix(0, fcSr) = "Sr.#"
33830     flxComp.ColWidth(fcSr) = 500
33840     flxComp.ColAlignment(fcSr) = flexAlignLeftCenter
          
33850     flxComp.TextMatrix(0, fcSDate) = "Date"
33860     flxComp.ColWidth(fcSDate) = 1500
33870     flxComp.ColAlignment(fcSDate) = flexAlignLeftCenter
          
33880     flxComp.TextMatrix(0, fcSID) = "Sample ID"
33890     flxComp.ColWidth(fcSID) = 1500
33900     flxComp.ColAlignment(fcSID) = flexAlignLeftCenter
          
33910     flxComp.TextMatrix(0, fcISONo) = "I.No"
33920     flxComp.ColWidth(fcISONo) = 500
33930     flxComp.ColAlignment(fcISONo) = flexAlignLeftCenter
          
33940     flxComp.TextMatrix(0, fcISO) = "Isolates"
33950     flxComp.ColWidth(fcISO) = 2500
33960     flxComp.ColAlignment(fcISO) = flexAlignLeftCenter
          
33970     flxComp.TextMatrix(0, fcRpt) = "Reports"
33980     flxComp.ColWidth(fcRpt) = 2500
33990     flxComp.ColAlignment(fcRpt) = flexAlignLeftCenter
          
          
34000     Exit Sub
ERROR_FormatGrid:
          Dim strES As String
          Dim intEL As Integer

34010     intEL = Erl
34020     strES = Err.Description
34030     LogError "frmMain", "FormatGrid", intEL, strES
End Sub

Private Sub GetComparison()
34040     On Error GoTo ERROR_GetComparison
          
          Dim l_SQL As String
          Dim l_rs As Recordset
          Dim l_rsI As Recordset
          Dim l_Count As Integer
          Dim l_str As String
          
34050     flxComp.Rows = 1
34060     flxComp.row = 0
34070     l_Count = 0
          
34080     l_SQL = "Select SampleID From Isolates" 'Where RecordDateTime Between '" & Format(dtoFrom.Value, "yyyy/MM/dd") & "' And '" & Format(dtpTo.Value, "yyyy/MM/dd") & "'"
34090     Set l_rs = New Recordset
34100     RecOpenServer 0, l_rs, l_SQL
34110     If Not l_rs Is Nothing Then
34120         If Not l_rs.EOF Then
34130             While Not l_rs.EOF
34140                 l_SQL = "Select * from Isolates Where SampleID = '" & ConvertNull(l_rs!SampleID, "") & "'"
34150                 Set l_rsI = New Recordset
34160                 RecOpenServer 0, l_rsI, l_SQL
34170                 If Not l_rsI Is Nothing Then
34180                     If Not l_rsI.EOF Then
34190                         While Not l_rsI.EOF
34200                             l_Count = l_Count + 1
34210                             l_str = "" & vbTab & l_Count & vbTab & ConvertNull(l_rsI!RecordDateTime, 0) & vbTab & ConvertNull(l_rsI!SampleID, "") & vbTab & ConvertNull(l_rsI!IsolateNumber, "") & vbTab & ConvertNull(l_rsI!OrganismName, "")
34220                             flxComp.AddItem (l_str)
34230                             l_rsI.MoveNext
34240                             DoEvents
34250                         Wend
34260                     End If
34270                 End If
34280                 l_rs.MoveNext
34290             Wend
34300         End If
34310     End If
          
34320     Exit Sub
ERROR_GetComparison:
          Dim strES As String
          Dim intEL As Integer

34330     intEL = Erl
34340     strES = Err.Description
34350     LogError "frmMain", "GetComparison", intEL, strES
End Sub

Private Sub GetComparisonWithReport()
34360     On Error GoTo ERROR_GetComparison
          
          Dim l_SQL As String
          Dim l_rs As Recordset
          Dim l_CountLen As Integer
          Dim l_Count As Integer
          Dim l_str As String
          
34370     For l_Count = 1 To flxComp.Rows - 1
34380         l_str = ""
34390         l_str = Trim(flxComp.TextMatrix(l_Count, fcISONo)) & ": " & Trim(flxComp.TextMatrix(l_Count, fcISO))
34400         l_CountLen = Len(l_str)
34410         l_SQL = "Select * from Reports Where SampleID = '" & flxComp.TextMatrix(l_Count, fcSID) & "' And ReportType = 'Final Report'"
34420         Set l_rs = New Recordset
34430         RecOpenServer 0, l_rs, l_SQL
34440         If Not l_rs Is Nothing Then
34450             If Not l_rs.EOF Then
34460                 If InStr(1, ConvertNull(l_rs!Report, ""), l_str, vbTextCompare) > 0 Then
34470                     flxComp.TextMatrix(l_Count, fcRpt) = Left(Right(ConvertNull(l_rs!Report, ""), Len(ConvertNull(l_rs!Report, "")) - InStr(1, ConvertNull(l_rs!Report, ""), l_str, vbTextCompare) + 1), l_CountLen)
34480                 Else
34490                     flxComp.TextMatrix(l_Count, fcRpt) = "Not Found"
34500                 End If
34510                 DoEvents
34520             End If
34530         End If
34540     Next
          
34550     Exit Sub
ERROR_GetComparison:
          Dim strES As String
          Dim intEL As Integer

34560     intEL = Erl
34570     strES = Err.Description
34580     LogError "frmMain", "GetComparisonWithReport", intEL, strES
End Sub
'--- Junaid
